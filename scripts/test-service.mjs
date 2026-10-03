import fs from 'node:fs';
import assert from 'node:assert/strict';
import http from 'node:http';
import { randomBytes } from 'node:crypto';
import { performance } from 'node:perf_hooks';
const endpoint = JSON.parse(
  fs.readFileSync('.runtime/endpoint.json', 'utf8').replace(/^\uFEFF/, ''),
);
const headers = { Authorization: 'Bearer ' + endpoint.Token, 'Content-Type': 'application/json' };
const checks = [];
async function request(path, method = 'GET', body) {
  const response = await fetch(endpoint.BaseUrl + '/api' + path, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  return { status: response.status, body: await response.json() };
}
async function check(name, run) {
  const start = performance.now();
  await run();
  checks.push({ name, status: 'PASS', elapsedMs: Math.round(performance.now() - start) });
  console.log('PASS ' + name);
}
try {
  await check('unauthenticated request rejected', async () => {
    const r = await fetch(endpoint.BaseUrl + '/api/catalog');
    assert.equal(r.status, 401);
  });
  await check('foreign Origin rejected', async () => {
    const r = await fetch(endpoint.BaseUrl + '/api/catalog', {
      headers: { ...headers, Origin: 'https://malicious.example' },
    });
    assert.equal(r.status, 403);
  });
  await check('81 validated source facts served', async () => {
    const r = await request('/catalog');
    assert.equal(r.status, 200);
    assert.equal(r.body.facts.length, 81);
  });
  await check('invalid and unsupported requests rejected', async () => {
    assert.equal((await request('/facts/MSFT/EBITDA/FY2025')).status, 400);
    assert.equal((await request('/facts/UNKNOWN/Revenue/FY2025')).status, 404);
    assert.equal(
      (
        await request(
          '/facts/batch',
          'POST',
          Array(129).fill({ ticker: 'MSFT', metric: 'Revenue', period: 'FY2025' }),
        )
      ).status,
      400,
    );
  });
  await check('malformed, null-element and oversized payloads rejected', async () => {
    for (const body of ['{broken', 'null', '[null]']) {
      const r = await fetch(endpoint.BaseUrl + '/api/facts/batch', {
        method: 'POST',
        headers,
        body,
      });
      assert.equal(r.status, 400);
    }
    const r = await fetch(endpoint.BaseUrl + '/api/research', {
      method: 'POST',
      headers,
      body: JSON.stringify({ ticker: 'MSFT', question: 'x'.repeat(140000), useAi: false }),
    });
    assert.equal(r.status, 413);
  });
  await check('unexpected Host header rejected', async () => {
    const status = await new Promise((resolve, reject) => {
      http
        .get(endpoint.BaseUrl + '/health', { headers: { Host: 'attacker.example' } }, (r) => {
          r.resume();
          resolve(r.statusCode);
        })
        .on('error', reject);
    });
    assert.equal(status, 400);
  });
  await check('WebSocket authentication, events, disconnect and reconnect', async () => {
    const url = endpoint.BaseUrl.replace('http:', 'ws:') + '/api/events';
    await new Promise((resolve, reject) => {
      const ws = new WebSocket(url, ['ledgerlens.v1']);
      const timer = setTimeout(() => {
        ws.close();
        reject(new Error('Unauthorized WS did not reject'));
      }, 4000);
      ws.onerror = () => {
        clearTimeout(timer);
        resolve();
      };
      ws.onopen = () => {
        clearTimeout(timer);
        ws.close();
        reject(new Error('Unauthorized WS accepted'));
      };
    });
    const baseline = (await request('/diagnostics')).body.streamSubscribers;
    for (let cycle = 0; cycle < 3; cycle++) {
      const ws = new WebSocket(url, ['ledgerlens.v1', 'll-auth.' + endpoint.Token]);
      const events = [];
      await new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('WS connect timeout')), 4000);
        ws.onmessage = (e) => {
          events.push(JSON.parse(e.data));
          if (events[0].type === 'connected') {
            clearTimeout(timer);
            resolve();
          }
        };
        ws.onerror = () => reject(new Error('WS connect failed'));
      });
      await request('/events/test', 'POST', {});
      for (let i = 0; i < 40 && !events.some((e) => e.type === 'diagnostic'); i++)
        await new Promise((r) => setTimeout(r, 25));
      assert(events.some((e) => e.type === 'diagnostic'));
      await new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error('WS close timeout')), 4000);
        ws.onclose = () => {
          clearTimeout(timer);
          resolve();
        };
        ws.close(1000);
      });
    }
    for (let i = 0; i < 40; i++) {
      if ((await request('/diagnostics')).body.streamSubscribers === baseline) break;
      await new Promise((r) => setTimeout(r, 25));
    }
    assert.equal((await request('/diagnostics')).body.streamSubscribers, baseline);
  });
  await check('abrupt WebSocket transport loss releases every subscription', async () => {
    const baseline = (await request('/diagnostics')).body.streamSubscribers;
    await Promise.all(
      Array.from(
        { length: 20 },
        () =>
          new Promise((resolve, reject) => {
            const connection = http.request(endpoint.BaseUrl + '/api/events', {
              headers: {
                Connection: 'Upgrade',
                Upgrade: 'websocket',
                'Sec-WebSocket-Version': '13',
                'Sec-WebSocket-Key': randomBytes(16).toString('base64'),
                'Sec-WebSocket-Protocol': 'ledgerlens.v1, ll-auth.' + endpoint.Token,
              },
            });
            connection.setTimeout(5000, () =>
              connection.destroy(new Error('WebSocket upgrade timed out')),
            );
            connection.on('error', reject);
            connection.on('response', (response) => {
              response.resume();
              reject(new Error('Upgrade rejected: ' + response.statusCode));
            });
            connection.on('upgrade', (_response, socket) => {
              socket.destroy();
              resolve();
            });
            connection.end();
          }),
      ),
    );
    for (let attempt = 0; attempt < 80; attempt++) {
      if ((await request('/diagnostics')).body.streamSubscribers === baseline) break;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
    assert.equal((await request('/diagnostics')).body.streamSubscribers, baseline);
  });
  await check('snapshot reads never call a remote provider', async () => {
    await request('/connection', 'POST', { mode: 'online' });
    const before = (await request('/diagnostics')).body;
    const responses = await Promise.all(
      Array.from({ length: 100 }, () => request('/facts/MSFT/Revenue/FY2025')),
    );
    assert(responses.every((r) => r.status === 200 && r.body.fact.value === 281724));
    const after = (await request('/diagnostics')).body;
    assert.equal(after.snapshotReads - before.snapshotReads, 100);
    assert.equal(after.secRequests, before.secRequests);
    assert.equal(after.secActiveRequests, 0);
    assert.equal('providerCalls' in after, false);
  });
  await check('2000 mixed snapshot requests preserve financial values', async () => {
    const facts = (await request('/catalog')).body.facts;
    const before = (await request('/diagnostics')).body;
    const responses = await Promise.all(
      Array.from({ length: 2000 }, (_, i) => {
        const f = facts[i % facts.length];
        return request('/facts/' + f.ticker + '/' + f.metric + '/' + f.period);
      }),
    );
    responses.forEach((r, i) => {
      assert.equal(r.status, 200);
      assert.equal(r.body.fact.value, facts[i % facts.length].value);
    });
    const after = (await request('/diagnostics')).body;
    assert.equal(after.secRequests, before.secRequests);
    assert.equal(after.snapshotReads - before.snapshotReads, 2000);
  });
  await check('offline uses saved facts and blocks external requests', async () => {
    await request('/connection', 'POST', { mode: 'offline' });
    const before = (await request('/diagnostics')).body;
    const result = await request('/facts/AAPL/Revenue/FY2025');
    assert.equal(result.body.fact.value, 416161);
    assert.equal(result.body.freshness, 'cached');
    assert.equal((await request('/sync/MSFT', 'POST', {})).status, 400);
    const research = await request('/research', 'POST', {
      ticker: 'MSFT',
      question: 'Compare annual revenue',
      useAi: true,
    });
    assert.equal(research.body.isAiGenerated, false);
    const after = (await request('/diagnostics')).body;
    assert.equal(after.secRequests, before.secRequests);
    assert.equal(after.aiCalls, before.aiCalls);
  });
  await check('production API rejects artificial failure modes', async () => {
    for (const mode of ['slow', 'unavailable'])
      assert.equal((await request('/connection', 'POST', { mode })).status, 400);
  });
  await check('online mode restores the external request preference', async () => {
    await request('/connection', 'POST', { mode: 'online' });
    assert.equal((await request('/diagnostics')).body.mode, 'online');
    assert.equal((await request('/facts/NVDA/Revenue/FY2025')).body.freshness, 'snapshot');
  });
  await check('calculated research is fully cited', async () => {
    const result = await request('/research', 'POST', {
      ticker: 'MSFT',
      question: 'What happened to margins and cash?',
      useAi: false,
    });
    assert.equal(result.status, 200);
    assert.equal(result.body.isAiGenerated, false);
    const ids = new Set(result.body.sources.map((f) => f.sourceId));
    assert(
      result.body.claims.every((c) => c.sourceIds.length && c.sourceIds.every((id) => ids.has(id))),
    );
  });
  const durations = [];
  await check('warm request latency sampled 100 times', async () => {
    for (let i = 0; i < 100; i++) {
      const start = performance.now();
      assert.equal((await request('/facts/NVDA/Revenue/FY2025')).status, 200);
      durations.push(performance.now() - start);
    }
  });
  durations.sort((a, b) => a - b);
  fs.writeFileSync(
    'artifacts/test-results/service-results.json',
    JSON.stringify(
      {
        checkedAt: new Date().toISOString(),
        checks,
        performance: {
          samples: durations.length,
          warmP50Ms: durations[49],
          warmP95Ms: durations[94],
        },
      },
      null,
      2,
    ),
  );
} finally {
  await request('/connection', 'POST', { mode: 'online' });
}
