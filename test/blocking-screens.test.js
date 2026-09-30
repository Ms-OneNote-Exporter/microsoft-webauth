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

const { clearBlockingScreens } = require('../src/auth');


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
