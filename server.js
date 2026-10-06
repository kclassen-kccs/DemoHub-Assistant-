// Outlook People directory agent.
// A server-side Chromium (Playwright) is streamed to the web UI so the user can log in
// interactively (MFA, SSO, etc.). Once signed in, the agent drives the Outlook People
// search box, captures the JSON the page itself fetches, and returns it for a chosen user.
const express = require('express');
const { chromium } = require('playwright-core');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const PEOPLE_URL = process.env.PEOPLE_URL || 'https://outlook.office.com/people/';
const STATE_FILE = path.join(__dirname, '.session', 'state.json');
const VIEWPORT = { width: 1280, height: 800 };

let browser, context, page;
const captured = []; // recent JSON responses {url, json, t}

function exePath() {
  if (process.env.CHROMIUM_PATH) return process.env.CHROMIUM_PATH;
  try { const p = chromium.executablePath(); if (fs.existsSync(p)) return p; } catch {}
  // fall back to any Chromium found under PLAYWRIGHT_BROWSERS_PATH (version may differ from playwright-core's)
  const root = process.env.PLAYWRIGHT_BROWSERS_PATH;
  if (root && fs.existsSync(root)) {
    for (const d of fs.readdirSync(root).filter((n) => /^chromium-\d+$/.test(n)).sort().reverse()) {
      for (const sub of ['chrome-linux/chrome', 'chrome-linux64/chrome']) {
        const f = path.join(root, d, sub);
        if (fs.existsSync(f)) return f;
      }
    }
  }
  return undefined;
}

async function ensureBrowser() {
  if (page && !page.isClosed()) return page;
  browser = browser || await chromium.launch({
    headless: process.env.HEADFUL ? false : true,
    executablePath: exePath(),
    args: ['--no-sandbox'],
  });
  const opts = { viewport: VIEWPORT };
  if (fs.existsSync(STATE_FILE)) opts.storageState = STATE_FILE;
  context = await browser.newContext(opts);
  page = await context.newPage();
  page.on('response', async (res) => {
    try {
      const ct = res.headers()['content-type'] || '';
      if (!ct.includes('json') || !/office|microsoft|live|sharepoint/.test(res.url())) return;
      const json = await res.json();
      captured.push({ url: res.url(), json, t: Date.now() });
      if (captured.length > 200) captured.shift();
    } catch {}
  });
  await page.goto(PEOPLE_URL, { waitUntil: 'domcontentloaded' }).catch(() => {});
  return page;
}

async function saveSession() {
  fs.mkdirSync(path.dirname(STATE_FILE), { recursive: true });
  await context.storageState({ path: STATE_FILE });
}

function isLoggedIn() {
  if (!page || page.isClosed()) return false;
  const u = new URL(page.url());
  return u.hostname === 'outlook.office.com' || u.hostname === 'outlook.cloud.microsoft';
}

// ---- JSON normalisation: find persona-like objects anywhere in a payload ----
const pick = (o, ...keys) => { for (const k of keys) if (o[k] != null && o[k] !== '') return o[k]; };
function emailOf(o) {
  const direct = pick(o, 'mail', 'Mail', 'EmailAddress', 'emailAddress', 'userPrincipalName', 'UserPrincipalName', 'PrimarySmtpAddress');
  if (typeof direct === 'string') return direct;
  if (direct && typeof direct === 'object') return direct.Address || direct.address || null;
  const list = pick(o, 'EmailAddresses', 'emailAddresses', 'scoredEmailAddresses');
  if (Array.isArray(list) && list.length) { const e = list[0]; return typeof e === 'string' ? e : (e.Address || e.address || null); }
  return null;
}
function normalise(o) {
  const name = pick(o, 'displayName', 'DisplayName');
  if (typeof name !== 'string') return null;
  const email = emailOf(o);
  if (!email && !pick(o, 'id', 'Id', 'PersonaId')) return null;
  return {
    id: pick(o, 'id', 'Id', 'objectId', 'ObjectId') || (o.PersonaId && o.PersonaId.Id) || null,
    displayName: name,
    email,
    jobTitle: pick(o, 'jobTitle', 'JobTitle', 'Title', 'title') || null,
    department: pick(o, 'department', 'Department') || null,
    company: pick(o, 'companyName', 'CompanyName', 'Company') || null,
    office: pick(o, 'officeLocation', 'OfficeLocation', 'Office') || null,
    phones: [].concat(pick(o, 'businessPhones', 'BusinessPhones') || [], pick(o, 'mobilePhone', 'MobilePhone') || []),
    manager: (o.manager && o.manager.displayName) || null,
    raw: o,
  };
}
function extract(entries) {
  const found = new Map();
  const walk = (v) => {
    if (Array.isArray(v)) return v.forEach(walk);
    if (!v || typeof v !== 'object') return;
    const n = normalise(v);
    if (n) {
      const key = (n.email || n.id || n.displayName).toLowerCase();
      const prev = found.get(key);
      // keep the richest record (most non-null fields) per person
      const score = (x) => Object.values(x).filter((y) => y != null && !(Array.isArray(y) && !y.length)).length + Object.keys(x.raw).length;
      if (!prev || score(n) > score(prev)) found.set(key, n);
    }
    Object.values(v).forEach(walk);
  };
  entries.forEach((e) => walk(e.json));
  return [...found.values()];
}

async function runSearch(q) {
  await ensureBrowser();
  if (!isLoggedIn()) throw Object.assign(new Error('Not signed in yet'), { status: 401 });
  if (!page.url().includes('/people')) await page.goto(PEOPLE_URL, { waitUntil: 'domcontentloaded' });
  const box = page.locator('input[type="search"], input[role="combobox"], input[aria-label*="Search" i], input[placeholder*="Search" i]').first();
  await box.waitFor({ state: 'visible', timeout: 30000 });
  const mark = Date.now();
  await box.click();
  await box.fill('');
  await box.pressSequentially(q, { delay: 40 });
  await box.press('Enter');
  await page.waitForLoadState('networkidle', { timeout: 15000 }).catch(() => {});
  await page.waitForTimeout(1500);
  const fresh = captured.filter((c) => c.t >= mark);
  return { results: extract(fresh), fresh };
}

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
const wrap = (fn) => (req, res) => fn(req, res).catch((e) => res.status(e.status || 500).json({ error: e.message }));

// --- remote browser view / input (for interactive login) ---
app.post('/api/start', wrap(async (req, res) => { await ensureBrowser(); res.json({ url: page.url() }); }));
app.get('/api/screen', wrap(async (req, res) => {
  await ensureBrowser();
  res.set('Content-Type', 'image/jpeg').set('Cache-Control', 'no-store')
    .send(await page.screenshot({ type: 'jpeg', quality: 70 }));
}));
app.get('/api/status', wrap(async (req, res) => {
  const loggedIn = isLoggedIn();
  if (loggedIn) await saveSession().catch(() => {});
  res.json({ started: !!page && !page.isClosed(), url: page && !page.isClosed() ? page.url() : null, loggedIn });
}));
app.post('/api/click', wrap(async (req, res) => { await ensureBrowser(); await page.mouse.click(+req.body.x, +req.body.y); res.json({ ok: true }); }));
app.post('/api/type', wrap(async (req, res) => { await ensureBrowser(); await page.keyboard.type(String(req.body.text), { delay: 15 }); res.json({ ok: true }); }));
app.post('/api/key', wrap(async (req, res) => { await ensureBrowser(); await page.keyboard.press(String(req.body.key)); res.json({ ok: true }); }));
app.post('/api/scroll', wrap(async (req, res) => { await ensureBrowser(); await page.mouse.wheel(0, +req.body.dy || 0); res.json({ ok: true }); }));
app.post('/api/logout', wrap(async (req, res) => {
  if (context) await context.close().catch(() => {});
  page = context = null;
  fs.rmSync(STATE_FILE, { force: true });
  res.json({ ok: true });
}));

// --- directory queries ---
app.get('/api/search', wrap(async (req, res) => {
  const q = String(req.query.q || '').trim();
  if (!q) return res.status(400).json({ error: 'q is required' });
  const { results } = await runSearch(q);
  res.json({ query: q, count: results.length, results: results.map(({ raw, ...r }) => r) });
}));

// JSON payload for one user: pass ?id=, ?email= or ?name=. Add &raw=1 for the full source objects.
app.get('/api/user', wrap(async (req, res) => {
  const { id, email, name, raw } = req.query;
  const term = email || name || id;
  if (!term) return res.status(400).json({ error: 'id, email or name is required' });
  const { results } = await runSearch(term);
  const want = String(term).toLowerCase();
  const user = results.find((r) =>
    (id && r.id === id) || (email && r.email && r.email.toLowerCase() === want) ||
    (name && r.displayName.toLowerCase() === want)) || (results.length === 1 ? results[0] : null);
  if (!user) return res.status(404).json({ error: 'No unique match', candidates: results.map(({ raw, ...r }) => r) });
  res.json(raw ? user : (({ raw: _r, ...rest }) => rest)(user));
}));

app.listen(PORT, () => console.log(`Outlook directory agent on http://localhost:${PORT}`));
process.on('SIGINT', async () => { await browser?.close().catch(() => {}); process.exit(0); });
