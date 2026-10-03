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