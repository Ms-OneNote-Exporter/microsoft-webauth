/**
 * @fileoverview The session probe behind `check`.
 *
 * The bug: `check` slept a fixed 2 s and asked whether the URL happened to be a
 * login host by then. Microsoft serves the app shell to a dead session and
 * redirects afterwards, so an auth file with no usable session came back
 * "authenticated" and the command exited 0 — the same class of failure as #24,
 * one layer deeper, and found while testing the exit-code fix.
 *
 * Pages are intercepted rather than served, so the hostnames are the real ones
 * (the login signal is a hostname match, not a path) and nothing here needs an
 * account or a network. The case that matters is the one that cannot be faked:
 * the silence of the unauthenticated shell.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const path = require('path');
const os = require('os');
const fs = require('fs-extra');
const { chromium, describeWithBrowser } = require('./helpers/browser-suite');

jest.mock('../src/utils/logger', () => ({
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    success: jest.fn(),
    debug: jest.fn(),
    step: jest.fn(),
    getDumpDir: async () => '/tmp/dumps',
    getDumpDisplayPath: () => 'logs/dumps/test'
}));

const { settleSessionProbe } = require('../src/auth');

const PAGE$ = (body) => `<!DOCTYPE html><html><head><title>t</title></head><body>${body}</body></html>`;

const ONENOTE_SHELL = 'https://onenote.cloud.microsoft/notebooks';
const APP_PATH = 'https://onenote.cloud.microsoft/copilotnotebooks';
const LOGIN = 'https://login.live.com/common/oauth2/v2.0/authorize';

// The authenticated OneNote app, post-rebrand. /notebooks is the entry point,
// /copilotnotebooks is where it actually lands.
const appPage = PAGE$(`<nav>All Notebooks</nav>`);

// The unauthenticated shell. It renders, offers a sign-up, and never navigates
// anywhere — which is the whole failure this probe has to call out.
const shellPage = PAGE$(`<h1>OneNote for the web</h1><a href="/en-us">Sign up for free</a>`);

const loginPage = PAGE$(`<h1>Sign in</h1><input name="loginfmt">`);

/** A page that sits where it is and navigates itself to `to` after `ms`. */
const afterMs = (to, ms) => PAGE$(`<script>
    setTimeout(() => { location.href = ${JSON.stringify(to)}; }, ${ms});
</script>`);

describeWithBrowser('settleSessionProbe — telling a live session from a dead one', () => {
    let browser;
    let context;

    beforeAll(async () => {
        browser = await chromium.launch({ headless: true });
        context = await browser.newContext();
    });

    afterAll(async () => {
        if (context) await context.close();
        if (browser) await browser.close();
    });

    /**
     * Loads `url`, then probes with `check`'s production target.
     *
     * `overrides` decides what each host serves, so a page that navigates itself
     * lands on the right one — which is the whole point of the late-arrival
     * cases below. Anything not named serves the unauthenticated shell, so a
     * test cannot pass by accident on a page that is simply missing.
     */
    const probe = async (url, overrides = {}, timeoutMs = 4000) => {
        const serves = host => (url => {
            const match = Object.keys(overrides).find(k => url.startsWith(k));
            return match ? overrides[match] : shellPage;
        })(host);

        const page = await context.newPage();
        try {
            await page.route('**/*', route => route.fulfill({
                status: 200,
                contentType: 'text/html',
                body: serves(route.request().url())
            }));
            await page.goto(url, { waitUntil: 'domcontentloaded' });
            return await settleSessionProbe(page, ONENOTE_SHELL, timeoutMs);
        } finally {
            await page.close();
        }
    };

    it('confirms a session already at the authenticated app', async () => {
        expect(await probe(APP_PATH, { [APP_PATH]: appPage })).toBe('app');
    });

    // The regression. This page sits still and never redirects, exactly like
    // the unauthenticated app shell, so a probe that waits for evidence rather
    // than a fixed delay has to eventually call it dead.
    it('calls a session that stays put unauthenticated', async () => {
        expect(await probe(ONENOTE_SHELL)).toBe('idle');
    });

    it('reports a session that gets sent to a login page', async () => {
        expect(await probe(LOGIN, { [LOGIN]: loginPage })).toBe('login');
    });

    // Both of these are about not deciding early. A probe that read page.url()
    // straight after goto would answer 'idle' here, and the old fixed 2 s sleep
    // had the same flaw in the other direction: it read 'app' for a session
    // that was never signed in.
    it('waits for the app to appear rather than deciding early', async () => {
        expect(await probe(ONENOTE_SHELL, { [APP_PATH]: appPage })).toBe('idle');
        expect(await probe(ONENOTE_SHELL, {
            [ONENOTE_SHELL]: afterMs(APP_PATH, 700),
            [APP_PATH]: appPage
        })).toBe('app');
    });

    it('waits for a late redirect rather than deciding early', async () => {
        expect(await probe(ONENOTE_SHELL, {
            [ONENOTE_SHELL]: afterMs(LOGIN, 700),
            [LOGIN]: loginPage
        })).toBe('login');
    });

    // And it must give up: a session that neither reaches the app nor the login
    // page has no session, and the caller needs an answer in bounded time.
    it('gives up rather than waiting forever', async () => {
        const started = Date.now();
        expect(await probe(ONENOTE_SHELL, {}, 1200)).toBe('idle');
        expect(Date.now() - started).toBeGreaterThanOrEqual(1000);
    });
});

/**
 * The probe above is only half of `check`, and the half that was wrong was not
 * missing — it was ignored. verifyAuth() awaited settleSessionProbe() and threw
 * the verdict away, then re-derived the answer from the URL on its own. Only a
 * redirect to a login host could count as a positive result, so every live
 * session was reported expired, on every run: a valid auth file that listed
 * notebooks perfectly well could still only make `check` exit 1.
 *
 * The probe's own tests pass while that is true, which is why this asserts on
 * the reported status instead. No real browser: the page resolves exactly the
 * signals each verdict describes and times out instantly on the rest, so all
 * three cases cost milliseconds rather than the production 30 s.
 */
describe('verifyAuth — the probe verdict is what gets reported', () => {
    let dir;
    let authFile;

    beforeEach(() => {
        jest.resetModules();
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'webauth-verdict-'));
        authFile = path.join(dir, 'auth-file.json');
        fs.writeJsonSync(authFile, { cookies: [{ name: 'ESAuth', value: 'x' }], origins: [] });
    });

    afterEach(() => {
        fs.rmSync(dir, { recursive: true, force: true });
        jest.dontMock('playwright');
        jest.resetModules();
    });

    /**
     * A Playwright stand-in parked at `url`, where `marker` names the one
     * signed-in selector that resolves and `login` decides whether the login-URL
     * wait resolves. Anything else rejects at once, standing in for the
     * timeout: the real waits would leave these tests taking a minute each.
     */
    const browserAt = (url, { marker, login } = {}) => {
        const arrived = yes => yes ? Promise.resolve() : Promise.reject(new Error('Timeout'));
        const page = {
            goto: async () => { },
            url: () => url,
            waitForSelector: selector => arrived(!!marker && selector.includes(marker)),
            waitForURL: () => arrived(!!login)
        };
        return {
            newContext: async () => ({ newPage: async () => page }),
            close: async () => { }
        };
    };

    const verifyWith = browser => {
        jest.doMock('playwright', () => ({ chromium: { launch: async () => browser } }));
        return require('../src/auth').verifyAuth(ONENOTE_SHELL, authFile);
    };

    // The reported bug, as a passing session. It reaches the signed-in
    // interface, so that is the answer — whatever the URL happens to say.
    it('reports a session that reaches the signed-in interface as authenticated', async () => {
        const status = await verifyWith(
            browserAt('https://onenote.cloud.microsoft/notebooks/', { marker: 'My notebooks' })
        );

        expect(status.authenticated).toBe(true);
        expect(status.reason).toBe('authenticated');
        expect(fs.existsSync(authFile)).toBe(true);
    });

    it('reports a session sent to a login page as expired, and clears it', async () => {
        const status = await verifyWith(
            browserAt('https://login.live.com/common/oauth2/v2.0/authorize', { login: true })
        );

        expect(status.authenticated).toBe(false);
        expect(status.reason).toBe('expired');
        // The one verdict that earns a deletion: Microsoft itself sent the
        // session to the login page.
        expect(fs.existsSync(authFile)).toBe(false);
    });

    // Silence is not proof of a dead session. It is also what a live session
    // gives when the app is slower than the wait — measured at ~10 s to render
    // the notebook UI, against a 10 s budget — so nothing may be deleted here.
    it('leaves the auth file alone when neither signal arrives', async () => {
        const status = await verifyWith(browserAt('https://onenote.cloud.microsoft/notebooks/'));

        expect(status.authenticated).toBe(false);
        // Not 'expired': nothing proved the session dead, only that this run
        // failed to see it, and the two deserve different advice.
        expect(status.reason).toBe('stayed_unauthenticated');
        expect(fs.existsSync(authFile)).toBe(true);
    });
});
