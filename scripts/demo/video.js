// The README demo: a scripted bug hunt on a copy of the Cart fixture through the vb6 MCP server, recorded as a
// timeline of events and screenshots for scripts/demo/videomaker, which renders the MP4 and the GIF:
//   node scripts/demo/video.js
//   dotnet run --project scripts/demo/videomaker -c Release -- build/demo/video build/demo/demo.mp4 docs/media/demo.gif
// The MP4 is attached to the GitHub release (it's too big for the repo); the GIF is the one in the README.
// Every screenshot is real (the add-in's PrintWindow of the IDE and of the program's form). What the renderer adds is
// drawn: the chat panel with the prompt and the tool calls, and a cursor that moves to the control each call acts on
// (the tools send input to the control itself; the real mouse doesn't move). The title card says so.
// Needs the add-in registered (build-addin.ps1), the server installed (install-server.ps1), and best a running
// watchdog (build-addin.ps1 -WatchdogOnly -Stay), so VB6 starts without a UAC prompt. The IDE is quit afterwards
// without saving. VIDEOMAKER_PNG="20,47" makes the renderer also save those seconds as PNGs, to check the layout.
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
const REPO = path.resolve(__dirname, '..', '..');
const EXE = path.join(REPO, 'build', 'server', 'current', 'vb6mcp.exe');
const FIXTURES = path.join(REPO, 'tests', 'fixtures');
// a fresh copy per run (tests/fixtures/scratch is git-ignored): an IDE left over from an earlier run may still hold
// the previous copy's files
const COPY = path.join(FIXTURES, 'scratch', `demo-${Date.now()}`);
const VBP = path.join(COPY, 'Cart.vbp');
const OUT = path.join(REPO, 'build', 'demo', 'video');
fs.rmSync(OUT, { recursive: true, force: true }); fs.mkdirSync(OUT, { recursive: true });
fs.rmSync(COPY, { recursive: true, force: true }); fs.mkdirSync(COPY, { recursive: true });
for (const f of ['Cart.vbp', 'frmCart.frm', 'modPricing.bas', 'clsLine.cls']) fs.copyFileSync(path.join(FIXTURES, 'Cart', f), path.join(COPY, f));

const p = spawn(EXE, [], { stdio: ['pipe', 'pipe', 'pipe'], env: { ...process.env, VB6MCP_ROOTS: FIXTURES } });
let buf = '', id = 0;
const waiting = new Map();
p.stdout.on('data', d => {
  buf += d;
  let i;
  while ((i = buf.indexOf('\n')) >= 0) {
    const line = buf.slice(0, i); buf = buf.slice(i + 1);
    if (!line.trim()) continue;
    const m = JSON.parse(line);
    if (waiting.has(m.id)) { waiting.get(m.id)(m); waiting.delete(m.id); }
  }
});
p.stderr.on('data', () => {});
const rpc = (method, params) => new Promise(res => { const n = ++id; waiting.set(n, res); p.stdin.write(JSON.stringify({ jsonrpc: '2.0', id: n, method, params }) + '\n'); });
let pid;
async function call(name, args = {}) {
  if (pid && !['ide_launch', 'symbol_refs'].includes(name)) args = { pid, ...args };
  const m = await rpc('tools/call', { name, arguments: args });
  const c = m.result?.content || [];
  const text = c.filter(x => x.type === 'text').map(x => x.text).join('\n');
  const image = c.find(x => x.type === 'image')?.data;
  const err = m.result?.isError || !!m.error;
  console.log(`== ${name} ${JSON.stringify(args).slice(0, 140)}${err ? ' [ERROR]' : ''}\n   ${(text || JSON.stringify(m.error)).slice(0, 200)}`);
  let json = null; try { json = JSON.parse(text); } catch { }
  if (err) throw new Error(`${name} failed: ${text}`);
  return { json, text, image };
}
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ---------- the timeline ----------
const events = [];
let shots = 0;
const ev = e => events.push(e);
/** A tool call as the chat panel shows it, then the call itself. */
async function tool(name, args, shown, result) {
  ev({ t: 'tool', call: shown, result: '' });
  const r = await call(name, args);
  events[events.length - 1].result = typeof result === 'function' ? result(r) : (result ?? '');
  return r;
}
async function shot(target) {
  const r = await call('app_screenshot', target === 'ide' ? { window: 'ide', max_width: 1600 } : { max_width: 1600 });
  const file = `s${String(++shots).padStart(2, '0')}.png`;
  fs.writeFileSync(path.join(OUT, file), Buffer.from(r.image, 'base64'));
  ev({ t: 'shot', target, file, clientOrigin: r.json?.clientOrigin });
}
/** Control rectangles of the running form (client coordinates), by name: text0, text1, Add, Clear. */
async function layout() {
  const r = await call('app_tree');
  const controls = {};
  for (const c of r.json.controls) {
    const key = c.kind === 'TextBox' ? `text${c.index}` : c.text.replace('&', '');
    controls[key] = { x: c.left, y: c.top, w: c.width, h: c.height };
  }
  ev({ t: 'layout', controls });
}
async function type(index, text) {
  ev({ t: 'cursor', target: `text${index}` });
  ev({ t: 'click' });
  await tool('app_set_text', { index, text }, `app_set_text  index=${index}  "${text}"`, r => `textNow: "${r.json.textNow}"`);
  await shot('app');
}
async function click(caption, breaks = false) {
  ev({ t: 'cursor', target: caption });
  ev({ t: 'click' });
  await tool('app_click', { text: caption }, `app_click  text="${caption}"`, r => r.json.mode === 'break' ? 'mode: break (breakpoint)' : 'clicked; mode: run');
  if (breaks) await shot('ide'); else await shot('app');
}
const caption = (text, tone = 'info', hold = 2.5) => ev({ t: 'caption', text, tone, hold });

(async () => {
  await rpc('initialize', { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'demo-video', version: '1' } });
  ev({ t: 'prompt', text: 'In the Cart project, adding 2 x 10.00 and then 1 x 5.00 shows a total of 10.00 instead of 25.00. Find the bug and fix it.' });

  const launch = await tool('ide_launch', { project: VBP }, 'ide_launch  project="Cart.vbp"', r => `started VB6 (pid ${r.json.pid}), add-in answers`);
  pid = launch.json.pid;
  // startup dialogs, answered off camera: another add-in's error before ours loads (the watchdog clicks it), then
  // SourceSafe. ide_status fails until the add-in has loaded, so this loop doesn't stop on errors.
  for (let k = 0; k < 15; k++) {
    let d = launch.json.dialogs?.[0];
    if (k > 0) {
      const s = await rpc('tools/call', { name: 'ide_status', arguments: { pid } });
      let j = null; try { j = JSON.parse(s.result?.content?.[0]?.text ?? ''); } catch { }
      if (j && !j.dialogs?.length && j.addin?.startsWith('answers')) break;
      d = j?.dialogs?.[0];
    }
    if (d) { try { await call('ide_dialog_respond', { pid, button: d.buttons.includes('No') ? 'No' : 'OK' }); } catch (e) { console.log(e.message); } }
    else await sleep(1500);
  }
  await call('form_screenshot', { component: 'frmCart', maximize: false });
  await shot('ide');
  caption('VB6 opens the project; the vb6-mcp add-in answers from inside the IDE');

  // reproduce
  await tool('run_start', {}, 'run_start', r => `mode: ${r.json.mode}`);
  await shot('ide');
  await layout();
  await shot('app');
  caption('First, reproduce it in the running program');
  await type(0, '10'); await type(1, '2'); await click('Add');
  await type(0, '5'); await type(1, '1'); await click('Add');
  caption('Reproduced: the list is right, but the total is 10.00, not 25.00', 'bug', 3.5);
  await tool('run_end', {}, 'run_end', r => `mode: ${r.json.mode}`);
  ev({ t: 'hide-app' });

  // investigate
  const code = await tool('code_read', { component: 'frmCart' }, 'code_read  frmCart', r => `${r.json.countOfLines} lines, hash ${r.json.hash.slice(0, 8)}...`);
  const at = code.json.text.split('\r\n').findIndex(l => l.trim() === 'lstLines.AddItem m_line.Describe') + 1;
  await tool('bp_toggle', { component: 'frmCart', line: at }, `bp_toggle  frmCart  line=${at}`, () => 'breakpoint set');
  await shot('ide');
  caption('A breakpoint right after a line is added to the cart');
  await tool('run_start', {}, 'run_start', r => `mode: ${r.json.mode}`);
  await shot('ide');
  await layout();
  await shot('app');
  await type(0, '10'); await type(1, '2'); await click('Add', true);
  await tool('run_continue', {}, 'run_continue', r => `mode: ${r.json.mode}`);
  await shot('app');
  await type(0, '5'); await type(1, '1'); await click('Add', true);
  const evr = await tool('debug_eval', { expression: 'm_lines(1).Describe & " | " & m_lines(2).Describe & " | same object: " & (m_lines(1) Is m_lines(2))' },
    'debug_eval  m_lines(1).Describe, m_lines(2).Describe, m_lines(1) Is m_lines(2)', r => r.json.result);
  await shot('ide');
  caption(evr.json.result.includes('True') ? 'Both cart entries are the same object: the second Add overwrote the first' : evr.json.result, 'bug', 4);
  await tool('run_end', { force: true }, 'run_end  force=true', r => `mode: ${r.json.mode}`);
  ev({ t: 'hide-app' });
  await tool('bp_clear_all', {}, 'bp_clear_all', () => 'cleared');

  // fix
  const code2 = await call('code_read', { component: 'frmCart' });
  const priceLine = code2.json.text.split('\r\n').findIndex(l => l.trim().startsWith('m_line.Price =')) + 1;
  await tool('code_replace_lines', { component: 'frmCart', start_line: priceLine, count: 0, new_text: '    Set m_line = New clsLine', expected_hash: code2.json.hash },
    `code_replace_lines  frmCart  line=${priceLine}  + "Set m_line = New clsLine"`, r => `new hash ${r.json.hash.slice(0, 8)}..., unsaved in the IDE`);
  await shot('ide');
  caption('The fix: a new object for every line, edited in the IDE (unsaved, undoable)', 'info', 3.5);

  // verify
  await tool('run_start', {}, 'run_start', r => `mode: ${r.json.mode}`);
  await shot('ide');
  await layout();
  await shot('app');
  await type(0, '10'); await type(1, '2'); await click('Add');
  await type(0, '5'); await type(1, '1'); await click('Add');
  caption('Fixed: 2 x 10.00 and 1 x 5.00, total 25.00', 'ok', 4);
  await tool('run_end', {}, 'run_end', r => `mode: ${r.json.mode}`);
  ev({ t: 'hide-app' });
  await shot('ide');
  ev({ t: 'answer', text: 'Fixed. cmdAdd_Click reused one clsLine object (m_line), so every Add put the same object into the cart and the last price and quantity counted for all lines. It now creates a new clsLine per Add; the total for your example is 25.00. The change is in the IDE and not saved yet.' });
  await call('ide_quit', { discard: true });

  fs.writeFileSync(path.join(OUT, 'timeline.json'), JSON.stringify(events, null, 2));
  console.log(`${events.length} events, ${shots} screenshots in ${OUT}`);
  p.stdin.end();
  setTimeout(() => process.exit(0), 1000);
})().catch(e => { console.error(e); process.exit(1); });
