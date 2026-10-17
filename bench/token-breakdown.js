// usage: node tokens.js <run.jsonl>...
// Where the context tokens of a headless run come from. Each tool result stays in the conversation and is re-read
// (cache read) on every later turn, so its cost is roughly size x remaining turns. Sizes: text ~ chars/4, images
// ~ width*height/750 (Anthropic's estimate), from the PNG/JPEG header.
const fs = require('fs');
function imgDims(b64) {
  const b = Buffer.from(b64.slice(0, 400), 'base64');
  if (b[0] === 0x89 && b[1] === 0x50) return [b.readUInt32BE(16), b.readUInt32BE(20)];
  if (b[0] === 0xff && b[1] === 0xd8) { // JPEG: scan for SOF0/2
    const full = Buffer.from(b64, 'base64');
    for (let i = 2; i < full.length - 9;) {
      if (full[i] !== 0xff) { i++; continue; }
      const m = full[i + 1]; const len = full.readUInt16BE(i + 2);
      if (m >= 0xc0 && m <= 0xc3) return [full.readUInt16BE(i + 7), full.readUInt16BE(i + 5)];
      i += 2 + len;
    }
  }
  return [0, 0];
}
const imgTokens = (w, h) => { // images are scaled to fit 1568 px long side / ~1.15 MP
  let s = Math.min(1, 1568 / Math.max(w, h)); w *= s; h *= s;
  const mp = w * h; if (mp > 1150000) { s = Math.sqrt(1150000 / mp); w *= s; h *= s; }
  return Math.round(w * h / 750);
};
for (const file of process.argv.slice(2)) {
  const lines = fs.readFileSync(file, 'utf8').split('\n').filter(Boolean).map(l => { try { return JSON.parse(l); } catch { return null; } }).filter(Boolean);
  const toolName = {}; const items = []; let turn = 0; let usage = { in: 0, cr: 0, cc: 0, out: 0 }; const seenMsg = new Set();
  for (const m of lines) {
    if (m.type === 'assistant') {
      const id = m.message.id; if (!seenMsg.has(id)) { seenMsg.add(id); turn++; const u = m.message.usage || {}; usage.in += u.input_tokens || 0; usage.cr += u.cache_read_input_tokens || 0; usage.cc += u.cache_creation_input_tokens || 0; usage.out += u.output_tokens || 0; }
      for (const c of m.message.content || []) {
        if (c.type === 'tool_use') { toolName[c.id] = c.name.replace('mcp__vb6__', ''); items.push({ turn, kind: 'call:' + c.name.replace('mcp__vb6__', ''), tok: JSON.stringify(c.input).length / 4 }); }
        if (c.type === 'text') items.push({ turn, kind: 'assistant text', tok: c.text.length / 4 });
        if (c.type === 'thinking') items.push({ turn, kind: 'thinking', tok: (c.thinking || '').length / 4 });
      }
    }
    if (m.type === 'user') for (const c of (m.message.content || [])) {
      if (c.type !== 'tool_result') continue;
      const name = toolName[c.tool_use_id] || '?';
      const parts = Array.isArray(c.content) ? c.content : [{ type: 'text', text: String(c.content) }];
      for (const p of parts) {
        if (p.type === 'image') { const [w, h] = imgDims(p.source.data); items.push({ turn, kind: 'image:' + name, tok: imgTokens(w, h), w, h }); }
        else items.push({ turn, kind: 'text:' + name, tok: (p.text || '').length / 4 });
      }
    }
  }
  const total = turn; const agg = {};
  for (const it of items) {
    const a = agg[it.kind] ??= { n: 0, tok: 0, reread: 0 };
    a.n++; a.tok += it.tok; a.reread += it.tok * Math.max(0, total - it.turn);
  }
  const sum = Object.values(agg).reduce((s, a) => s + a.reread, 0);
  const imgs = items.filter(i => i.kind.startsWith('image:'));
  console.log(`\n${file}: ${total} turns; usage in=${usage.in} cacheRead=${usage.cr} cacheCreate=${usage.cc} out=${usage.out}`);
  console.log(`estimated re-read of tool traffic: ${(sum / 1e6).toFixed(1)} M tokens; images ${imgs.length}, avg ${Math.round(imgs.reduce((s, i) => s + i.tok, 0) / Math.max(1, imgs.length))} tokens, sizes e.g. ${imgs.slice(0, 4).map(i => i.w + 'x' + i.h).join(', ')}`);
  const groups = {};
  for (const [k, a] of Object.entries(agg)) { const g = k.startsWith('image:') ? 'images' : k.startsWith('text:') ? 'tool results (text)' : k.startsWith('call:') ? 'tool call arguments' : k; const x = groups[g] ??= { tok: 0, reread: 0 }; x.tok += a.tok; x.reread += a.reread; }
  for (const [g, a] of Object.entries(groups).sort((x, y) => y[1].reread - x[1].reread)) console.log(`  ${g.padEnd(24)} size ${(a.tok / 1000).toFixed(0).padStart(5)}k  re-read ${(a.reread / 1e6).toFixed(1).padStart(5)} M  ${(100 * a.reread / sum).toFixed(0)}%`);
  console.log('  top kinds by re-read:');
  for (const [k, a] of Object.entries(agg).sort((x, y) => y[1].reread - x[1].reread).slice(0, 14)) console.log(`    ${k.padEnd(34)} n=${String(a.n).padStart(3)} size ${(a.tok / 1000).toFixed(1).padStart(6)}k avg ${Math.round(a.tok / a.n).toString().padStart(5)}  re-read ${(a.reread / 1e6).toFixed(2)} M`);
}
