const http = require('http');
const { spawn } = require('child_process');

const BASE = 'http://127.0.0.1:3456';
let token = '';
let serverProcess;

function request(path, options = {}) {
  return new Promise((resolve, reject) => {
    const url = new URL(path, BASE);
    const req = http.request(url, {
      method: options.method || 'GET',
      headers: { 'Content-Type': 'application/json', ...(options.headers || {}) },
    }, (res) => {
      let body = '';
      res.on('data', (chunk) => { body += chunk; });
      res.on('end', () => {
        try {
          resolve({ status: res.statusCode, data: JSON.parse(body) });
        } catch (_error) {
          resolve({ status: res.statusCode, raw: body });
        }
      });
    });

    req.on('error', reject);
    if (options.body) {
      req.write(JSON.stringify(options.body));
    }
    req.end();
  });
}

async function waitForServer(timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    try {
      const response = await request('/api/health');
      if (response.status === 200) {
        return;
      }
    } catch (_error) {}
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  throw new Error('Server did not start in time');
}

async function runTest(name, fn) {
  try {
    await fn();
    return { name, ok: true };
  } catch (error) {
    return { name, ok: false, error: error.message };
  }
}

async function main() {
  serverProcess = spawn(process.execPath, ['src/server.js'], { stdio: 'ignore' });
  try {
    await waitForServer();

    const results = [];
    results.push(await runTest('Health check', async () => {
      const { status, data } = await request('/api/health');
      if (status !== 200 || !data.success) throw new Error('Health check failed');
    }));

    results.push(await runTest('Providers list', async () => {
      const { status, data } = await request('/api/ai/providers');
      if (status !== 200 || !Array.isArray(data.data) || data.data.length === 0) throw new Error('Providers were not returned');
    }));

    results.push(await runTest('Register user', async () => {
      const username = `test_${Date.now()}`;
      const { status, data } = await request('/api/auth/register', {
        method: 'POST',
        body: { username, password: 'test123456', displayName: 'Test User' },
      });
      if (status !== 201 || !data.data?.token) throw new Error('Registration failed');
      token = data.data.token;
    }));

    results.push(await runTest('Login', async () => {
      const username = `login_${Date.now()}`;
      await request('/api/auth/register', { method: 'POST', body: { username, password: 'test123456' } });
      const { status, data } = await request('/api/auth/login', {
        method: 'POST',
        body: { username, password: 'test123456' },
      });
      if (status !== 200 || !data.data?.token) throw new Error('Login failed');
    }));

    results.push(await runTest('Get user info', async () => {
      const { status, data } = await request('/api/auth/me', {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (status !== 200 || !data.data?.username) throw new Error('User info failed');
    }));

    results.push(await runTest('Save settings', async () => {
      const { status } = await request('/api/settings', {
        method: 'PUT',
        headers: { Authorization: `Bearer ${token}` },
        body: { provider: 'gemini', style: 'business', temperature: 0.6 },
      });
      if (status !== 200) throw new Error('Settings save failed');
    }));

    results.push(await runTest('Get settings', async () => {
      const { status, data } = await request('/api/settings', {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (status !== 200 || data.data?.provider !== 'gemini') throw new Error('Settings get failed');
    }));

    results.push(await runTest('Create document', async () => {
      const { status, data } = await request('/api/documents', {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}` },
        body: { title: 'Test doc', type: 'text', content: 'Hello world', tags: ['demo'] },
      });
      if (status !== 201 || !data.data?.id) throw new Error('Document creation failed');
    }));

    results.push(await runTest('404 route', async () => {
      const { status, data } = await request('/api/nonexistent');
      if (status !== 404 || data.success !== false) throw new Error('Expected 404 response');
    }));

    const passed = results.filter((item) => item.ok).length;
    const failed = results.filter((item) => !item.ok);

    console.log(`Results: ${passed} passed, ${failed.length} failed`);
    for (const result of failed) {
      console.log(`FAIL: ${result.name} - ${result.error}`);
    }

    if (failed.length > 0) {
      process.exitCode = 1;
    }
  } finally {
    if (serverProcess) {
      serverProcess.kill('SIGTERM');
    }
  }
}

main().catch((error) => {
  console.error(error);
  if (serverProcess) {
    serverProcess.kill('SIGTERM');
  }
  process.exit(1);
});
