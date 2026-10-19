// Spike 42: a minimal MCP server whose one tool misbehaves on purpose, to see what Claude Code does.
// FAKE_MODE=silent    tools/call never answers (a dead or frozen server, as the client sees it)
// FAKE_MODE=progress  sends notifications/progress every 3 s for 21 s, then answers
// FAKE_MODE=late      answers after 21 s with no progress in between
const mode = process.env.FAKE_MODE || 'silent';
const log = m => process.stderr.write(`[fake ${mode}] ${new Date().toISOString().slice(11, 23)} ${m}\n`);
let buf = '';
const send = o => process.stdout.write(JSON.stringify(o) + '\n');
process.stdin.on('data', d => {
  buf += d;
  let i;
  while ((i = buf.indexOf('\n')) >= 0) {
    const line = buf.slice(0, i); buf = buf.slice(i + 1);
    if (!line.trim()) continue;
    const m = JSON.parse(line);
    if (m.method === 'initialize')
      send({ jsonrpc: '2.0', id: m.id, result: { protocolVersion: m.params.protocolVersion, capabilities: { tools: {} }, serverInfo: { name: 'fake', version: '1' } } });
    else if (m.method === 'tools/list')
      send({ jsonrpc: '2.0', id: m.id, result: { tools: [{ name: 'slow', description: 'A slow operation. Call it once with no arguments and report what it returns.', inputSchema: { type: 'object', properties: {} } }] } });
    else if (m.method === 'tools/call') {
      const token = m.params?._meta?.progressToken;
      log(`tools/call, progressToken=${token}`);
      const answer = () => { log('answering'); send({ jsonrpc: '2.0', id: m.id, result: { content: [{ type: 'text', text: `done (${mode})` }] } }); };
      if (mode === 'progress') {
        let n = 0;
        const t = setInterval(() => {
          n++;
          if (token !== undefined) send({ jsonrpc: '2.0', method: 'notifications/progress', params: { progressToken: token, progress: n, message: `still working (${n * 3} s)` } });
          log(`progress ${n}`);
          if (n === Math.round((+process.env.FAKE_SECONDS || 21) / 3)) { clearInterval(t); answer(); }
        }, 3000);
      } else if (mode === 'late') setTimeout(answer, (+process.env.FAKE_SECONDS || 21) * 1000);
      // silent: never answer
    } else if (m.id !== undefined) send({ jsonrpc: '2.0', id: m.id, error: { code: -32601, message: 'Method not found' } });
  }
});
