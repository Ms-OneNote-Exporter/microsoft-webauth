const { chromium, describeWithBrowser } = require('./helpers/browser-suite');

jest.mock('../src/utils/logger', () => {
    // A real temp dir, not a fixed path: this suite now dumps the late blocking
    // screens it catches, and "did a PNG land?" is not a question a nonexistent
    // directory can answer. Built inside the factory because jest.mock may not
    // close over out-of-scope variables.
    const nodeFs = require('fs');
    const nodePath = require('path');
    const dir = nodeFs.mkdtempSync(nodePath.join(require('os').tmpdir(), 'blocking-screens-'));
    return {
        info: jest.fn(),
        warn: jest.fn(),
        error: jest.fn(),
        success: jest.fn(),
        debug: jest.fn(),
        step: jest.fn(),
        getDumpDir: async () => dir,
        getDumpDisplayPath: () => 'logs/dumps/test'
    };
});

const fs = require('fs');
const path = require('path');

const logger = require('../src/utils/logger');
const { clearBlockingScreens, watchBlockingScreens } = require('../src/auth');


/**
 * These tests replay the two screens that actually block a real login, using the
 * markup captured from the live pages (see the debug_ HTML dumps under
 * logs/dumps). They are stubbed over page.route so nothing touches the network
 * and the suite is deterministic.
 */

const TOU_URL = 'https://account.live.com/tou/accrue?mkt=EN-US';
const FRESHNESS_URL = 'https://account.live.com/pf?mkt=EN-US';
const KMSI_URL = 'https://login.live.com/ppsecure/post.srf?username=x';
const FIDO_URL = 'https://login.microsoft.com/consumers/fido/create?mkt=en-US';
const ONENOTE_URL = 'https://onenote.cloud.microsoft/notebooks';
const WRONG_URL = 'https://onenote.cloud.microsoft/should-not-happen';

const page$ = (title, body) => `<!DOCTYPE html><html><head><title>${title}</title></head><body>${body}</body></html>`;

// "We're updating our terms" — Fluent UI, primary button + footer links that must
// never be touched.
const touHtml = page$('We\'re updating our terms', `
    <h1>We're updating our terms</h1>
    <a href="https://go.microsoft.com/fwlink/?LinkID=2092201">Learn more about these updates</a>
    <button type="submit" data-testid="primaryButton"
        onclick="__record('Next'); location.href='${FRESHNESS_URL}'">Next</button>
    <a href="https://www.microsoft.com/servicesagreement/en-us">Terms of use</a>
    <a href="https://privacy.microsoft.com/en-us">Privacy and cookies</a>
    <button type="button" onclick="__record('Help')">Help</button>
`);

// "Is your security info still accurate?" — three competing actions, only one of
// which is safe to press.
const freshnessHtml = page$('Is your security info still accurate?', `
    <h1>Is your security info still accurate?</h1>
    <a href="#" id="iMarkLost" onclick="__record('I don\\'t have any of these'); location.href='${WRONG_URL}'">I don't have any of these</a>
    <input type="button" role="button" id="iUpdateNow" value="Update now"
        onclick="__record('Update now'); location.href='${WRONG_URL}'">
    <input type="button" role="button" id="iLooksGood" class="btn-primary" value="Looks good!"
        onclick="__record('Looks good!'); location.href='${ONENOTE_URL}'">
`);

// A stuck screen: the button is there, clicking it changes nothing.
const stuckHtml = page$(TOU_URL, `
    <h1>We're updating our terms</h1>
    <button type="submit" data-testid="primaryButton" onclick="__record('Next')">Next</button>
`);

// "Stay signed in?" — the prompt Microsoft shows last, right before the app.
const kmsiHtml = page$('Sign in', `
    <h1>Stay signed in?</h1>
    <div id="KmsiDescription">Do you want to stay signed in?</div>
    <input type="checkbox" name="DontShowAgain" id="KmsiCheckboxField">
    <input type="submit" id="idSIButton9" value="Yes"
        onclick="__record('Yes, dontShowAgain=' + document.getElementById('KmsiCheckboxField').checked); location.href='${ONENOTE_URL}'">
    <input type="button" id="idSIButton8" value="No"
        onclick="__record('No'); location.href='${WRONG_URL}'">
`);

// The passkey prompt. The native WebAuthn dialog is suppressed by the
// addInitScript under test, so only the page-level Cancel is reachable.
const fidoHtml = page$('Sign in', `
    <h1>Sign in with a passkey</h1>
    <button type="button" class="cancel" onclick="__record('Cancel'); location.href='${ONENOTE_URL}'">Cancel</button>
`);

const oneNoteHtml = page$('OneNote', '<h1>My notebooks</h1>');

describeWithBrowser('clearBlockingScreens', () => {
    let browser;
    let context;
    let page;
    let clicks;

    /** Serves the captured markup for any of the stubbed URLs. */
    async function stubPages() {
        await page.route('**/*', route => {
            const url = route.request().url();
            const body = url.startsWith('https://account.live.com/tou/') ? (url.includes('stuck=1') ? stuckHtml : touHtml)
                : url.startsWith('https://account.live.com/pf') ? freshnessHtml
                    : url.startsWith('https://login.microsoft.com/consumers/fido/') ? fidoHtml
                        : url.startsWith('https://login.live.com/ppsecure/') ? kmsiHtml
                            : url.startsWith('https://onenote.cloud.microsoft/') ? oneNoteHtml
                                : page$('Unknown', '<h1>Unknown</h1>');
            return route.fulfill({ contentType: 'text/html', body });
        });
        await page.exposeFunction('__record', label => clicks.push(label));
    }

    // Fast timings: the real defaults (20 s change timeout) would dominate the suite.
    const options = progress => ({ progress, changeTimeout: 1500, stateTimeout: 3000 });

    beforeAll(async () => {
        browser = await chromium.launch({ headless: true });
    });

    afterAll(async () => {
        await browser.close();
    });

    beforeEach(async () => {
        clicks = [];
        context = await browser.newContext();
        page = await context.newPage();
        await stubPages();
    });

    afterEach(async () => {
        await context.close();
    });

    it('clears the terms screen and the security freshness screen that follows it', async () => {
        await page.goto(TOU_URL);

        const result = await clearBlockingScreens(page, options({ signatures: new Set(), lastClickAt: 0 }));

        expect(result.handled).toBe(2);
        expect(result.reason).toBe('no_blocking_screen');
        expect(clicks).toEqual(['Next', 'Looks good!']);
        expect(page.url()).toBe(ONENOTE_URL);
    }, 30000);

    it('never presses an account-altering action on the freshness screen', async () => {
        await page.goto(FRESHNESS_URL);

        const result = await clearBlockingScreens(page, options({ signatures: new Set(), lastClickAt: 0 }));

        expect(result.handled).toBe(1);
        // "Update now" and "I don't have any of these" would rewrite or wipe the
        // account's recovery methods — neither may ever be clicked.
        expect(clicks).toEqual(['Looks good!']);
        expect(page.url()).not.toBe(WRONG_URL);
    }, 30000);

    it('ignores links and buttons that are not the accepted action', async () => {
        await page.goto(TOU_URL);

        await clearBlockingScreens(page, options({ signatures: new Set(), lastClickAt: 0 }));

        expect(clicks).not.toContain('Help');
        expect(clicks.filter(c => c === 'Next')).toHaveLength(1);
    }, 30000);

    it('leaves a normal authenticated page completely alone', async () => {
        await page.goto(ONENOTE_URL);

        const result = await clearBlockingScreens(page, options({ signatures: new Set(), lastClickAt: 0 }));

        expect(result.handled).toBe(0);
        expect(result.reason).toBe('no_blocking_screen');
        expect(clicks).toEqual([]);
    }, 30000);

    it('does not re-click an unchanged screen within the cooldown window', async () => {
        await page.goto(`${TOU_URL}&stuck=1`);
        const progress = { signatures: new Set(), lastClickAt: 0 };

        const first = await clearBlockingScreens(page, options(progress));
        const second = await clearBlockingScreens(page, options(progress));

        expect(first.handled).toBe(1);
        expect(second.handled).toBe(0);
        expect(second.reason).toBe('unchanged');
        expect(clicks).toEqual(['Next']);
    }, 30000);

    it('answers "Stay signed in?" with Yes and ticks "do not show again"', async () => {
        await page.goto(KMSI_URL);

        const result = await clearBlockingScreens(page, options({ signatures: new Set(), lastClickAt: 0 }));

        expect(result.handled).toBe(1);
        // "Yes", not "No", and the box ticked so later logins skip the screen.
        expect(clicks).toEqual(['Yes, dontShowAgain=true']);
    }, 30000);

    it('cancels the passkey prompt', async () => {
        await page.goto(FIDO_URL);

        const result = await clearBlockingScreens(page, options({ signatures: new Set(), lastClickAt: 0 }));

        expect(result.handled).toBe(1);
        expect(clicks).toEqual(['Cancel']);
        expect(page.url()).toBe(ONENOTE_URL);
    }, 30000);

    it('stops immediately when asked, so the watcher cannot outlive the login', async () => {
        await page.goto(TOU_URL);
        let stop = true;

        const result = await clearBlockingScreens(page, {
            ...options({ signatures: new Set(), lastClickAt: 0 }),
            shouldStop: () => stop
        });

        expect(result.handled).toBe(0);
        expect(result.reason).toBe('stopped');
        expect(clicks).toEqual([]);
    }, 30000);
});

/**
 * The watcher is the pass that catches the interstitial screens Microsoft serves
 * *late* — 20-60 s in, typically behind a "Stay signed in?" prompt. It used to
 * clear them with no dumping at all, so a login killed by one of them left
 * nothing in the dump directory and the only evidence was a bare timeout.
 */
describeWithBrowser('watchBlockingScreens', () => {
    let browser;
    let context;
    let page;
    let clicks;
    let dumpDir;

    async function stubPages() {
        await page.route('**/*', route => {
            const url = route.request().url();
            const body = url.startsWith('https://account.live.com/tou/') ? touHtml
                : url.startsWith('https://account.live.com/pf') ? freshnessHtml
                    : url.startsWith('https://login.microsoft.com/consumers/fido/') ? fidoHtml
                        : url.startsWith('https://login.live.com/ppsecure/') ? kmsiHtml
                            : url.startsWith('https://onenote.cloud.microsoft/') ? oneNoteHtml
                                : page$('Unknown', '<h1>Unknown</h1>');
            return route.fulfill({ contentType: 'text/html', body });
        });
        await page.exposeFunction('__record', label => clicks.push(label));
    }

    beforeAll(async () => {
        browser = await chromium.launch({ headless: true });
    });

    afterAll(async () => {
        await browser.close();
    });

    beforeEach(async () => {
        jest.clearAllMocks();
        clicks = [];
        dumpDir = await logger.getDumpDir();
        // The dump directory is shared by the whole file, and several tests
        // assert that a file was *not* written. Without emptying it, the first
        // test to dump would make every later negative assertion fail on its
        // leftovers rather than on anything the code did.
        for (const name of fs.readdirSync(dumpDir)) {
            fs.rmSync(path.join(dumpDir, name), { force: true, recursive: true });
        }
        context = await browser.newContext();
        page = await context.newPage();
        await stubPages();
    });

    afterEach(async () => {
        await context.close();
    });

    const exists = name => fs.existsSync(path.join(dumpDir, name));

    /** Runs the watcher until the page leaves the blocking screen, then stops it. */
    const watch = async (overrides = {}) => {
        let stop = false;
        const progress = { signatures: new Set(), lastClickAt: 0 };
        const watching = watchBlockingScreens(page, {
            progress,
            pollInterval: 200,
            maxScreens: 2,
            // The production 20 s would make this suite take half a minute; the
            // screens under test all navigate away immediately.
            stateTimeout: 3000,
            changeTimeout: 1500,
            dodump: true,
            screenshot: true,
            shouldStop: () => stop,
            ...overrides
        });
        // Let it poll until the screen is cleared, then wind the watcher down.
        await page.waitForURL(url => !url.toString().includes('account.live.com'), { timeout: 10000 })
            .catch(() => {});
        stop = true;
        await watching;
        return progress;
    };

    it('dumps and screenshots a late blocking screen instead of clearing it blind', async () => {
        await page.goto(TOU_URL);

        await watch();

        // This is the whole point of the change: the screen is late, and before
        // it dumped, a login that died here had produced no evidence at all.
        expect(clicks).toEqual(['Next', 'Looks good!']);
        expect(exists('debug_late_blocking_screen_1.html')).toBe(true);
        expect(exists('debug_late_blocking_screen_1.png')).toBe(true);
    }, 30000);

    it('does not overwrite the early pass dump, which is a different screen', async () => {
        await page.goto(FRESHNESS_URL);

        // What step 2.5a writes when it catches a screen early.
        await clearBlockingScreens(page, {
            progress: { signatures: new Set(), lastClickAt: 0 },
            stateTimeout: 3000,
            changeTimeout: 1500,
            dodump: true,
            screenshot: true
        });
        const earlyHtml = fs.readFileSync(path.join(dumpDir, 'debug_blocking_screen_1.html'), 'utf8');

        await page.goto(TOU_URL);
        await watch();

        // The watcher numbers its dumps from 1 too, so a shared basename would
        // have had this late screen overwrite the early one — and the early one
        // is the screen that was actually in the way when the login stalled.
        const lateHtml = fs.readFileSync(path.join(dumpDir, 'debug_late_blocking_screen_1.html'), 'utf8');
        expect(earlyHtml).toContain('Is your security info still accurate?');
        expect(lateHtml).toContain("We're updating our terms");
        expect(earlyHtml).not.toBe(lateHtml);
    }, 30000);

    it('picks up a screen that appears only after the watcher has started', async () => {
        // Starts on a clean authenticated page: the watcher must not need the
        // screen to be there when it begins polling.
        await page.goto(ONENOTE_URL);
        const onNotebooks = page.url();

        let stop = false;
        const watching = watchBlockingScreens(page, {
            progress: { signatures: new Set(), lastClickAt: 0 },
            pollInterval: 200,
            maxScreens: 2,
            stateTimeout: 3000,
            changeTimeout: 1500,
            dodump: true,
            screenshot: true,
            shouldStop: () => stop
        });

        // The "Stay signed in?" prompt lands while the watcher is already running.
        await page.waitForTimeout(300);
        await page.goto(KMSI_URL);
        await page.waitForURL(url => url.toString().includes('onenote.cloud.microsoft'), { timeout: 10000 })
            .catch(() => {});
        stop = true;
        await watching;

        expect(clicks).toContain('Yes, dontShowAgain=true');
        expect(exists('debug_late_blocking_screen_1.html')).toBe(true);
        expect(onNotebooks).toBe(ONENOTE_URL);
    }, 30000);

    it('takes no screenshots when screenshots were not asked for', async () => {
        await page.goto(TOU_URL);

        await watch({ screenshot: false });

        expect(exists('debug_late_blocking_screen_1.html')).toBe(true);
        expect(exists('debug_late_blocking_screen_1.png')).toBe(false);
    }, 30000);

    it('writes nothing at all when dumping was not asked for', async () => {
        await page.goto(TOU_URL);

        await watch({ dodump: false });

        // The default of a login run without --dodump: the screens are still
        // cleared, and the dump directory stays untouched.
        expect(clicks).toEqual(['Next', 'Looks good!']);
        expect(exists('debug_late_blocking_screen_1.html')).toBe(false);
        expect(exists('debug_late_blocking_screen_1.png')).toBe(false);
    }, 30000);

    it('is wired to the real login, so a late screen is dumped in production too', () => {
        // Everything else in this suite calls watchBlockingScreens directly, which
        // means none of it can see whether login() actually hands the watcher the
        // dump flags. It did not, for the whole life of the watcher: the loop ran
        // with no dumping at all and the tests passed anyway. This asserts the
        // wiring instead, so dropping it is a failing test rather than a silent
        // return to the old behaviour.
        const source = fs.readFileSync(
            path.join(__dirname, '..', 'src', 'auth.js'), 'utf8');

        const calls = source.match(/watchBlockingScreens\(page, \{[\s\S]*?\}\);/g) || [];

        expect(calls.length).toBeGreaterThanOrEqual(1);
        const call = calls.find(c => c.includes('stopWatcher'));
        expect(call).toBeDefined();
        expect(call).toContain('dodump');
        expect(call).toContain('screenshot');
    });

    it('stops when asked and leaves the watcher no longer running', async () => {
        await page.goto(ONENOTE_URL);
        let stop = true;

        await watchBlockingScreens(page, {
            progress: { signatures: new Set(), lastClickAt: 0 },
            pollInterval: 200,
            shouldStop: () => stop
        });

        // A watcher that outlived its login would keep clicking pages that no
        // longer belong to it, so resolving on the stop flag is the guarantee.
        expect(clicks).toEqual([]);
    }, 30000);

    it('does not rewrite the dump on every poll while a screen sits unchanged', async () => {
        // A screen that never goes away: the retry cooldown must suppress the
        // re-dump, or the watcher would write a full-page PNG every second.
        await page.route('**/*', route => route.fulfill({
            contentType: 'text/html',
            body: stuckHtml
        }));
        await page.goto(`${TOU_URL}&stuck=1`);

        let stop = false;
        const watching = watchBlockingScreens(page, {
            progress: { signatures: new Set(), lastClickAt: 0 },
            pollInterval: 200,
            maxScreens: 2,
            stateTimeout: 3000,
            changeTimeout: 1500,
            dodump: true,
            screenshot: true,
            shouldStop: () => stop
        });
        // Long enough for several polls, and comfortably inside the 10 s retry
        // cooldown, so any repeat click would be the watcher's fault.
        await page.waitForTimeout(3000);
        stop = true;
        await watching;

        // One click despite ~7 polls: the cooldown inside clearBlockingScreens
        // is what stops the watcher hammering an unchanged screen.
        expect(clicks).toEqual(['Next']);
    }, 30000);
});
