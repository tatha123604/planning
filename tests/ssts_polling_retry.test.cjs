const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const test = require('node:test');

const template = fs.readFileSync('templates/ssts_report.html', 'utf8');
const helpers = template.slice(template.indexOf('      const readJsonResponse'), template.indexOf('      const pollStatus'));

function client(responses) {
  let calls = 0;
  const context = {
    window: { setTimeout: fn => fn() },
    fetch: async () => {
      const result = responses[Math.min(calls++, responses.length - 1)];
      if (result instanceof Error) throw result;
      return {status: result, headers: { get: () => 'application/json' }, json: async () => ({status: 'completed'})};
    },
  };
  vm.createContext(context);
  vm.runInContext(helpers + '\nthis.request = fetchJsonWithRetry;', context);
  return {request: context.request, calls: () => calls};
}

test('status polling recovers from transient 502 and 504 responses', async () => {
  const api = client([502, 504, 200]);
  const result = await api.request('/status', {method: 'GET'}, 'Status failed');
  assert.equal(result.payload.status, 'completed');
  assert.equal(api.calls(), 3);
});

test('persistent upstream failure stops after bounded retries', async () => {
  const api = client([503]);
  await assert.rejects(api.request('/status', {method: 'GET'}, 'Status failed'), /HTTP 503/);
  assert.equal(api.calls(), 3);
});

test('start requests are not duplicated after network or gateway errors', async () => {
  for (const failure of [502, new Error('Connection lost')]) {
    const api = client([failure]);
    await assert.rejects(api.request('/start', {method: 'POST'}, 'Start failed'));
    assert.equal(api.calls(), 1);
  }
});

test('missing task responses are returned without retrying', async () => {
  const api = client([404]);
  const result = await api.request('/status', {method: 'GET'}, 'Status failed');
  assert.equal(result.response.status, 404);
  assert.equal(api.calls(), 1);
});
