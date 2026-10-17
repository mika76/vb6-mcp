// Writes the tool index for docs/tools.md from the server's own tools/list, so it can't drift from the code.
// Usage: node scripts/tool-index.js <vb6mcp.exe> <out.md>
const { spawn } = require('child_process');
const fs = require('fs');
const [exe, outFile] = process.argv.slice(2);
const p = spawn(exe, [], { env: { ...process.env, VB6MCP_ROOTS: process.cwd() } });
let buf = '';
const send = o => p.stdin.write(JSON.stringify(o) + '\n');
p.stdout.on('data', d => {
  buf += d; let i;
  while ((i = buf.indexOf('\n')) >= 0) {
    const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1);
    if (m.id === 1) { send({ jsonrpc: '2.0', method: 'notifications/initialized' }); send({ jsonrpc: '2.0', id: 2, method: 'tools/list' }); continue; }
    if (m.id !== 2) continue;
    const area = n => /^ide_|^server_/.test(n) ? 'IDEs' : /^project_/.test(n) ? 'Projects' : /^(code_|symbol_|proc_)/.test(n) ? 'Code'
      : /^(run_|step_|bp_|debug_|syntax_)/.test(n) ? 'Run and debug' : /^(app_|exe_)/.test(n) ? 'Running program' : /^form_/.test(n) ? 'Forms'
      : /^refs_/.test(n) ? 'References' : /^(com_|compat_)/.test(n) ? 'COM' : /^build_/.test(n) ? 'Builds' : /^sync_/.test(n) ? 'Sync'
      : /^recovery_/.test(n) ? 'Recovery' : /^(ws_|file_|res_)/.test(n) ? 'Workspace (disk)' : 'Other';
    const order = ['IDEs', 'Projects', 'Code', 'Run and debug', 'Running program', 'Forms', 'References', 'COM', 'Builds', 'Sync', 'Workspace (disk)', 'Recovery', 'Other'];
    const first = d => { const s = d.split(/(?<!\be\.g|\bi\.e)(?<=[a-z0-9)\]'"])\. /)[0].replace(/\|/g, '\\|'); return s.length > 150 ? s.slice(0, 147) + '…' : s.replace(/\.$/, ''); };
    const tools = m.result.tools.map(t => ({ name: t.name, area: area(t.name), what: first(t.description) }))
      .sort((a, b) => order.indexOf(a.area) - order.indexOf(b.area) || a.name.localeCompare(b.name));
    const rows = ['## Tool index', '', `All ${tools.length} tools, generated from the server's tool list (\`node scripts/tool-index.js\`). Details below, by topic.`, '',
      '| Area | Tool | What it does |', '|---|---|---|', ...tools.map(t => `| ${t.area} | \`${t.name}\` | ${t.what} |`), ''];
    fs.writeFileSync(outFile, rows.join('\r\n'));
    console.log(`${tools.length} tools`);
    p.stdin.end(); p.kill();
  }
});
send({ jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'tool-index', version: '1' } } });
