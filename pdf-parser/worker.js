// lf-pdf-parser — Cloudflare Worker that turns a financier's PDF amortisation
// schedule into the plain rows the amortisation tool's paste box accepts.
//
// Flow: amortisation.html POSTs { pdf: <base64> } with the caller's Microsoft
// Graph token in Authorization. The worker verifies the caller is an LF
// Advisory tenant user via Graph /me (so the endpoint can't be used by the
// public — no secret ever ships in the page), then sends the PDF to Claude
// and returns { principal, balloon, rows }.
//
// Secrets: ANTHROPIC_API_KEY (secret_text binding).
// Deploy: REST API upload, same method as register-mailer (see CLAUDE.md).

const ALLOWED_ORIGINS = ['https://www.lfadvisory.com.au', 'https://lfadvisory.com.au'];
const MODEL = 'claude-sonnet-5';

const PROMPT = `You are given a finance amortisation schedule PDF from a financier (hire purchase, chattel mortgage or equipment loan).

Extract the repayment schedule and reply with EXACTLY this plain-text format and nothing else:

PRINCIPAL: <amount financed at day one, plain number, or UNKNOWN>
BALLOON: <balloon/residual payable at the end of the term, plain number, or 0>
ROWS:
<one line per scheduled payment: date,payment,interest  OR  date,payment,interest,fee,gst>

Column mapping (layouts and names vary between financiers):
- date: the repayment/due date, as dd/mm/yyyy. If only month/year is shown (e.g. 07/2026), use day 01.
- payment: the TOTAL cash paid that period, INCLUDING any per-payment fee or GST shown for the row. The column may be named Repayment, Repayment Amount, Instalment, Payment Amount, etc. If the schedule shows no total column, or the shown total excludes the fee, output payment = principal + interest + fee (+ gst) instead of the shown figure.
- interest: the interest / term charge for that period (0 if none shown).
- fee: per-payment administration/account-keeping fees, if the schedule has such a column (e.g. Administration Fees). gst: per-payment GST credits if shown. Use the 5-column form only when the schedule has these columns.

Verify before replying (fix the rows, do not explain):
- Every row must satisfy: payment - fee - gst - interest = that row's principal reduction (it must match the movement in the balance column when one is shown).
- PRINCIPAL: use the stated amount financed. If not stated, derive it from the schedule (e.g. the first row's closing balance plus that row's principal reduction). Use UNKNOWN only if it genuinely cannot be determined.
- The rows must fully repay PRINCIPAL:
  - SKIP settlement/drawdown rows where both payment and interest are zero.
  - If the balloon is shown as (or included in) the final scheduled payment, include it there.
  - If the schedule instead ENDS with a remaining balance still owing (a positive final balance), that balance is the balloon: append one extra row dated the same as the final repayment, with payment = that remaining balance and interest = 0. A trivial residual of a few cents (rounding) needs no extra row.
- Amounts as plain numbers with no currency symbols, commas, or spaces (e.g. 1623.10).
- Include every repayment row. Do not summarise, skip repayment rows, or add totals, headers, or commentary.
- If the document is not an amortisation/repayment schedule (e.g. it is a loan contract or invoice), reply exactly: ERROR: not an amortisation schedule`;

export default {
  async fetch(request, env) {
    const origin = request.headers.get('Origin') || '';
    const cors = {
      'Access-Control-Allow-Origin': ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0],
      'Access-Control-Allow-Methods': 'POST, OPTIONS',
      'Access-Control-Allow-Headers': 'Authorization, Content-Type',
      'Access-Control-Max-Age': '86400',
    };
    if (request.method === 'OPTIONS') return new Response(null, { status: 204, headers: cors });
    if (request.method !== 'POST') return json({ error: 'POST only' }, 405, cors);

    // 1. Caller must be a signed-in LF Advisory user (portal's shared Graph token)
    const auth = request.headers.get('Authorization') || '';
    if (!auth.startsWith('Bearer ')) return json({ error: 'No sign-in token — sign in at the portal first.' }, 401, cors);
    const me = await fetch('https://graph.microsoft.com/v1.0/me', { headers: { Authorization: auth } });
    if (!me.ok) return json({ error: 'Microsoft sign-in rejected — go back to the portal and sign in again.' }, 401, cors);
    const user = await me.json();
    const upn = (user.userPrincipalName || user.mail || '').toLowerCase();
    if (!upn.endsWith('@lfadvisory.com.au')) return json({ error: 'Not an LF Advisory account.' }, 403, cors);

    // 2. Read the PDF (base64)
    let body;
    try { body = await request.json(); } catch { return json({ error: 'Bad request body.' }, 400, cors); }
    const pdf = String(body.pdf || '').replace(/^data:application\/pdf;base64,/, '');
    if (!pdf) return json({ error: 'No PDF supplied.' }, 400, cors);
    if (pdf.length > 25 * 1024 * 1024) return json({ error: 'PDF too large (max ~18MB).' }, 413, cors);

    // 3. Have Claude extract the schedule
    const resp = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key': env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
        'content-type': 'application/json',
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: 8192,
        messages: [{
          role: 'user',
          content: [
            { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: pdf } },
            { type: 'text', text: PROMPT },
          ],
        }],
      }),
    });
    if (!resp.ok) {
      console.log('Anthropic error', resp.status, await resp.text());
      return json({ error: 'Claude API error ' + resp.status + ' — try again shortly.' }, 502, cors);
    }
    const data = await resp.json();
    if (data.stop_reason === 'refusal') return json({ error: 'The model declined to process this document.' }, 422, cors);
    let text = '';
    for (const block of data.content || []) if (block.type === 'text') text += block.text;
    if (data.stop_reason === 'max_tokens') return json({ error: 'Schedule too long for one pass — split the PDF and upload the parts.' }, 422, cors);
    if (text.trim().startsWith('ERROR:')) return json({ error: text.trim() }, 422, cors);

    // 4. Parse the structured reply into fields the page can use directly
    const principalMatch = text.match(/PRINCIPAL:\s*([\d.]+)/i);
    const balloonMatch = text.match(/BALLOON:\s*([\d.]+)/i);
    const rowsIdx = text.search(/ROWS:/i);
    const rows = rowsIdx >= 0 ? text.slice(rowsIdx + 5).trim() : '';
    if (!rows) return json({ error: 'No schedule rows found in the document — check the PDF and try again.' }, 422, cors);

    // arithmetic sanity: rows should repay the principal (payment - fee - gst - interest summed)
    let drift = null;
    const pNum = principalMatch ? +principalMatch[1] : NaN;
    if (isFinite(pNum) && pNum > 0) {
      let paid = 0;
      for (const line of rows.split('\n')) {
        const c = line.split(',').map(x => x.trim());
        const pay = +c[1], int = +c[2], fee = c[3] !== undefined ? +c[3] : 0, gst = c[4] !== undefined ? +c[4] : 0;
        if (isFinite(pay) && isFinite(int)) paid += pay - (isFinite(fee) ? fee : 0) - (isFinite(gst) ? gst : 0) - int;
      }
      drift = Math.round((pNum - paid) * 100) / 100;
    }

    return json({
      principal: principalMatch ? principalMatch[1] : '',
      balloon: balloonMatch ? balloonMatch[1] : '',
      rows,
      drift,
      usage: data.usage,
    }, 200, cors);
  },
};

function json(obj, status, cors) {
  return new Response(JSON.stringify(obj), { status, headers: { 'content-type': 'application/json', ...cors } });
}
