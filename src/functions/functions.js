/**
 * AI Custom Functions for Excel — Shared Runtime
 * Namespace: AI => =AI(), =AI.EXPLAIN(), =AI.TABLE()
 */
/* global CustomFunctions, Excel */

function getCfg() {
  try { return JSON.parse(localStorage.getItem('aiexcel_config')) || {}; }
  catch (_) { return {}; }
}

async function callAI(prompt, sysMsg) {
  const cfg = getCfg();
  if (!cfg.apiKey) throw new Error('Set API Key di task pane > Settings');

  const ep = cfg.endpoint || 'https://9router.britamax.my.id/v1/chat/completions';
  const mdl = cfg.model || 'openrouter/deepseek-v4-flash';

  const res = await fetch(ep, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + cfg.apiKey },
    body: JSON.stringify({
      model: mdl,
      messages: [
        { role: 'system', content: sysMsg || 'You are an AI assistant inside Excel. Be concise.' },
        { role: 'user', content: prompt }
      ],
      temperature: cfg.temperature || 0.7,
      stream: false,
    }),
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
  // If cell reference provided, read its value
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
    // No arg → read selected range
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
