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
const logger = require('../src/utils/logger');
const { chromium, describeWithBrowser } = require('./helpers/browser-suite');

jest.mock('../src/utils/logger', () => {
    // Required inside the factory: jest.mock factories may not close over
    // out-of-scope variables, so the shared dump dir has to be built here. A
    // per-file temp dir rather than a fixed path, because the --dodump tests
    // below assert on real files and jest runs test files in parallel workers.
    const nodeFs = require('fs');
    const nodePath = require('path');
    const dir = nodeFs.mkdtempSync(nodePath.join(require('os').tmpdir(), 'check-dump-'));
    return {
        info: jest.fn(),
        warn: jest.fn(),
        error: jest.fn(),
        success: jest.fn(),
        debug: jest.fn(),
        step: jest.fn(),
        getDumpDir: jest.fn(async () => dir),
        getDumpDisplayPath: () => 'logs/dumps/test'
    };
});

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
/**
 * A Playwright stand-in parked at `url`, where `marker` names the one signed-in
 * selector that resolves and `login` decides whether the login-URL wait
 * resolves. Anything else rejects at once, standing in for the timeout: the real
 * waits would leave these tests taking a minute each.
 *
 * `evaluate` and `screenshot` are here for the --dodump suite below, which goes
 * through the real dumpPage(): evaluate stands in for the redacting read of the
 * document, screenshot for the PNG written beside it.
 */
const browserAt = (url, { marker, login } = {}) => {
    const arrived = yes => yes ? Promise.resolve() : Promise.reject(new Error('Timeout'));
    const page = {
        goto: async () => { },
        url: () => url,
        waitForSelector: selector => arrived(!!marker && selector.includes(marker)),
        waitForURL: () => arrived(!!login),
        evaluate: async () => `<!DOCTYPE html><html><body>stub page at ${url}</body></html>`,
        screenshot: async ({ path: pngPath }) => fs.writeFileSync(pngPath, Buffer.alloc(8))
    };
    return {
        newContext: async () => ({ newPage: async () => page }),
        close: async () => { }
    };
};

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

    const verifyWith = (browser, options = {}) => {
        jest.doMock('playwright', () => ({ chromium: { launch: async () => browser } }));
        return require('../src/auth').verifyAuth({
            targetUrl: ONENOTE_SHELL,
            authFilePath: authFile,
            ...options
        });
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

/**
 * `--dodump` on `check`, which exists for the case the verdict cannot explain.
 *
 * "check said expired" is a single line and any number of causes: a dead
 * session, a login page served for an unrelated reason, a captive portal, an
 * app shell that never finished rendering. The verdict names which of those it
 * thinks it saw; the dump is what lets a person confirm it. So the pages that
 * decide the answer have to be on disk, at every stage of the run.
 *
 * No real browser: the stub above answers exactly the signals each verdict
 * describes, and the real dumpPage() runs for real against a temp dir, so the
 * files asserted on here are the ones production writes.
 */
describe('check --dodump', () => {
    let dir;
    let authFile;
    let dumpDir;

    /**
     * The logger of the *current* module registry.
     *
     * jest.resetModules() in beforeEach re-runs the mock factory, so auth.js and
     * this test share an instance only if both ask for it after the reset —
     * a reference captured at file scope would be a different object, with a
     * different dump directory, from the one dumpPage() writes through.
     */
    const currentLogger = () => require('../src/utils/logger');

    const verifyWith = (browser, options = {}) => {
        jest.doMock('playwright', () => ({ chromium: { launch: async () => browser } }));
        return require('../src/auth').verifyAuth({
            targetUrl: ONENOTE_SHELL,
            authFilePath: authFile,
            ...options
        });
    };

    const APP_BROWSER = () => browserAt('https://onenote.cloud.microsoft/notebooks/', { marker: 'All Notebooks' });

    /**
     * One stub whose verdict follows `state`, so a single test can produce all
     * three — and therefore write all three into one dump directory, which is
     * the only way to show that they do not overwrite each other.
     *
     * A per-call stub would not do: src/auth.js captures chromium when it is
     * first required, so every later call in the same test would keep getting
     * whichever browser was set up first.
     */
    const browserFollowing = state => {
        const arrived = yes => yes ? Promise.resolve() : Promise.reject(new Error('Timeout'));
        const page = {
            goto: async () => { },
            url: () => state.url,
            waitForSelector: selector => arrived(!!state.marker && selector.includes(state.marker)),
            waitForURL: () => arrived(!!state.login),
            evaluate: async () => '<!DOCTYPE html><html><body>stub</body></html>',
            screenshot: async ({ path: pngPath }) => fs.writeFileSync(pngPath, Buffer.alloc(8))
        };
        return {
            newContext: async () => ({ newPage: async () => page }),
            close: async () => { }
        };
    };

    const written = name => fs.existsSync(path.join(dumpDir, name));

    beforeEach(async () => {
        jest.resetModules();
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'webauth-check-dump-'));
        authFile = path.join(dir, 'auth-file.json');
        fs.writeJsonSync(authFile, { cookies: [{ name: 'ESAuth', value: 'x' }], origins: [] });
        dumpDir = await currentLogger().getDumpDir();
    });

    afterEach(() => {
        fs.rmSync(dir, { recursive: true, force: true });
        jest.dontMock('playwright');
        jest.resetModules();
    });

    it('writes the page it navigated to, and the page that decided the verdict', async () => {
        const status = await verifyWith(APP_BROWSER(), { dodump: true });

        expect(status.authenticated).toBe(true);
        expect(written('debug_check_after_nav.html')).toBe(true);
        expect(written('debug_check_app.html')).toBe(true);
    });

    // The three are separate files on purpose. `check` is the command people
    // re-run, its dump directory is per-minute, and a single `debug_check.html`
    // would let a run that reported "expired" be overwritten minutes later by
    // one that reported "stayed_unauthenticated" — destroying exactly the
    // evidence the flag was asked for.
    it('names each verdict dump after the verdict that produced it', async () => {
        const state = { url: 'https://onenote.cloud.microsoft/notebooks/', marker: 'All Notebooks' };
        await verifyWith(browserFollowing(state), { dodump: true });
        Object.assign(state, { url: ONENOTE_SHELL, marker: undefined });
        await verifyWith(browserFollowing(state), { dodump: true });
        // Last, because this verdict deletes the auth file.
        Object.assign(state, { url: LOGIN, login: true });
        await verifyWith(browserFollowing(state), { dodump: true });

        expect(written('debug_check_app.html')).toBe(true);
        expect(written('debug_check_idle.html')).toBe(true);
        expect(written('debug_check_login.html')).toBe(true);
    });

    it('writes a PNG beside every dump when screenshots are asked for', async () => {
        await verifyWith(APP_BROWSER(), { dodump: true, screenshot: true });

        for (const name of ['debug_check_after_nav', 'debug_check_app']) {
            expect(written(`${name}.html`)).toBe(true);
            expect(written(`${name}.png`)).toBe(true);
        }
    });

    // The default, and the regression this whole flag could cause: the dumps
    // must not start appearing on a plain `check`, which runs unattended in CI.
    it('writes nothing at all unless dumps were asked for', async () => {
        const status = await verifyWith(APP_BROWSER());

        expect(status.authenticated).toBe(true);
        expect(fs.readdirSync(dumpDir).filter(name => name.startsWith('debug_check'))).toEqual([]);
    });

    // The error path is where a dump earns its keep: the exception carries a
    // message and nothing about the page that produced it, and a navigation
    // timeout, a proxy page and a TLS failure all look alike in the log.
    it('dumps the page that errored', async () => {
        const failing = browserAt(ONENOTE_SHELL);
        failing.newContext = async () => ({
            newPage: async () => ({
                goto: async () => { throw new Error('net::ERR_TIMED_OUT'); },
                url: () => ONENOTE_SHELL,
                evaluate: async () => '<!DOCTYPE html><html><body>error page</body></html>'
            })
        });

        const status = await verifyWith(failing, { dodump: true });

        expect(status.reason).toBe('unverifiable');
        expect(written('debug_check_error.html')).toBe(true);
    });

    // A dump is a debugging aid; it must not be able to change the answer it
    // exists to explain. dumpPage() throws on a write failure, and the call
    // sites are inside verifyAuth()'s try block, so without this an unwritable
    // dump directory would be caught as a verification error and a live session
    // reported as unverifiable — the check failing for a reason of its own.
    it('still reports the verdict when the dump cannot be written', async () => {
        currentLogger().getDumpDir.mockResolvedValue(
            path.join(os.tmpdir(), 'check-dump-no-such-dir', 'nested')
        );

        const status = await verifyWith(APP_BROWSER(), { dodump: true });

        expect(status.authenticated).toBe(true);
        expect(status.reason).toBe('authenticated');
        expect(currentLogger().warn).toHaveBeenCalledWith(
            expect.stringContaining('the check itself is unaffected')
        );
    });

    // The two verdicts decided before a browser exists. Silence here would be
    // indistinguishable from a broken dump directory, and those need opposite
    // investigations.
    it('says why there is no dump when no page was ever loaded', async () => {
        const status = await require('../src/auth').verifyAuth({
            targetUrl: ONENOTE_SHELL,
            authFilePath: path.join(dir, 'nope.json'),
            dodump: true
        });

        expect(status.reason).toBe('no_auth_file');
        expect(currentLogger().info).toHaveBeenCalledWith(
            expect.stringContaining('Nothing to dump')
        );
    });

    // verifyAuth took two positionals until this change. A caller upgrading
    // would have no other way to find out: the string would destructure to an
    // undefined targetUrl and the check would run against OneNote and the
    // default auth file, answering a question about a different session with
    // what looks like a real result.
    it('rejects the old positional arguments rather than quietly changing meaning', async () => {
        const { verifyAuth } = require('../src/auth');

        await expect(verifyAuth(ONENOTE_SHELL, authFile))
            .rejects.toThrow(/single options object/);
    });
});
