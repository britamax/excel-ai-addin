/**
 * AI Custom Functions for Excel — Shared Runtime
 * Namespace: AI => =AI(), =AI.EXPLAIN(), =AI.TABLE()
 * Config shared via localStorage (set in task pane Settings).
 */
/* global CustomFunctions, Excel */

function getCfg() {
  try { return JSON.parse(localStorage.getItem('aiexcel_config')) || {}; }
  catch (_) { return {}; }
}

async function callAI(prompt, sysMsg) {
  const cfg = getCfg();
  if (!cfg.apiKey) throw new Error('Set Endpoint + API Key in task pane ⚙️ Settings first');
  if (!cfg.endpoint) throw new Error('Set Endpoint in task pane ⚙️ Settings first');

  const body = {
    messages: [
      { role: 'system', content: sysMsg || 'You are an AI assistant inside Excel. Be concise.' },
      { role: 'user', content: prompt }
    ],
    temperature: cfg.temperature || 0.7,
    stream: false,
  };
  if (cfg.model) body.model = cfg.model;

  const res = await fetch(cfg.endpoint, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + cfg.apiKey },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => '');
    throw new Error('API ' + res.status + ': ' + txt.slice(0, 150));
  }

  const d = await res.json();
  return d.choices?.[0]?.message?.content || d.choices?.[0]?.message?.reasoning_content || '(empty)';
}

// =AI(query) — simple ask
async function aiQuery(query) {
  return await callAI(query, 'You are Excel AI. Answer concisely.');
}

// =AI.EXPLAIN([cellRef]) — explain selected cell or text
async function aiExplain(ref) {
  let target = ref;
  if (ref && typeof ref === 'string' && /^[A-Z]+\d*$/i.test(ref.replace(/[^a-zA-Z0-9]/g,''))) {
    try {
      await Excel.run(async (ctx) => {
        const rng = ctx.workbook.worksheets.getActiveWorksheet().getRange(ref);
        rng.load('text');
        await ctx.sync();
        target = rng.text[0][0] || ref;
      });
    } catch (_) { target = ref; }
  } else if (!ref || ref === '') {
    try {
      await Excel.run(async (ctx) => {
        const sel = ctx.workbook.getSelectedRange();
        sel.load('text');
        await ctx.sync();
        target = sel.text[0][0] || '(empty cell)';
      });
    } catch (_) { target = '(unknown)'; }
  }
  return await callAI('Explain this: ' + target, 'You explain Excel data and formulas concisely.');
}

// =AI.TABLE([range], [instruction]) — analyze table data
async function aiTable(range, instruction) {
  let data = 'N/A';
  try {
    await Excel.run(async (ctx) => {
      let rng;
      if (range && typeof range === 'string') {
        rng = ctx.workbook.worksheets.getActiveWorksheet().getRange(range);
      } else {
        rng = ctx.workbook.getSelectedRange();
      }
      rng.load('values');
      await ctx.sync();
      data = JSON.stringify(rng.values);
    });
  } catch (_) { data = '(could not read range)'; }
  return await callAI(
    'Data: ' + data + '\n\nInstruction: ' + (instruction || 'Analyze this data'),
    'Analyze Excel table data, return concise insights.'
  );
}

// Register
CustomFunctions.associate('AI', aiQuery);
CustomFunctions.associate('AI.EXPLAIN', aiExplain);
CustomFunctions.associate('AI.TABLE', aiTable);
