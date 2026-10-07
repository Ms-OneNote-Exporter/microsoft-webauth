/**
 * Does the number-match path actually work?
 *
 * This test exists because the answer was assumed twice and checked never.
 * `dismissFidoPage()` cancels the FIDO2 hardware-key screen — correct, and
 * unrelated — and that, plus an `approver_prompt` reason from the *pre-password*
 * navigation, led to the conclusion that the package cancels rather than waits.
 *
 * It does wait. The code was inline inside `login()` with no coverage at all, so
 * it was extracted to `src/phone-approval.js` and is driven here against a real
 * browser.
 *
 * What is asserted is what a user depends on:
 *   1. the number is read from the screen and shown
 *   2. it **waits** rather than clicking anything
 *   3. it gives up cleanly when the user never answers, rather than hanging
 */
const { chromium, describeWithBrowser } = require('./helpers/browser-suite');

jest.mock('../src/utils/logger', () => ({
    info: jest.fn(), warn: jest.fn(), error: jest.fn(), success: jest.fn(),
    debug: jest.fn(), step: jest.fn(),
    getDumpDir: async () => '/tmp/dumps',
    getDumpDisplayPath: () => 'logs/dumps/test',
}));

const { waitForPhoneApproval, NUMBER_MATCH_SELECTOR } = require('../src/phone-approval');

const APPROVAL_URL = 'https://login.microsoftonline.com/common/oauth2/authorize?x=1';
const AFTER_URL = 'https://login.microsoftonline.com/keep?y=1';

const approvalPage = (body) =>
    `<!DOCTYPE html><html><head><title>Approve sign in request</title></head><body>${body}</body></html>`;

// The Fluent shape: the number under data-testid, which is what current
// Microsoft pages serve.
const FLUENT = approvalPage(`
    <h1>Approve sign in request</h1>
    <div data-testid="displaySign">42</div>
    <p>Open Microsoft Authenticator and enter the number.</p>
`);

// The legacy server-rendered shape, still in the selector list.
const LEGACY = approvalPage(`
    <h1>Approve sign in request</h1>
    <div class="displaySign">7</div>
`);

// Records every click, so "did it click anything?" is answerable rather than
// assumed. This is the whole claim: a cancel-based implementation clicks one.
const WITH_BUTTONS = approvalPage(`
    <h1>Approve sign in request</h1>
    <div data-testid="displaySign">42</div>
    <button id="approve" onclick="window.__clicked=(window.__clicked||0)+1">Approve</button>
    <button id="cancel" onclick="window.__clicked=(window.__clicked||0)+1">Cancel</button>
`);

/** Passed as `{ logger }` — the function reads that key, not the object itself. */
const quiet = { logger: { warn() {}, step() {}, info() {}, success() {}, debug() {} } };

describeWithBrowser('the number-match wait', () => {
    let browser;
    let context;
    let page;

    beforeAll(async () => {
        browser = await chromium.launch({ headless: true });
    });

    afterAll(async () => {
        await browser.close();
    });

    beforeEach(async () => {
        context = await browser.newContext();
        page = await context.newPage();
    });

    afterEach(async () => {
        await context.close();
    });

    /**
     * Force the production 120 s wait down to something a test can afford.
     *
     * Only the `state: 'hidden'` wait on the number-match selector is shortened;
     * everything else passes through untouched, so the other two racers are
     * still exercised as written.
     */
    const shortWait = (ms) => {
        const real = page.waitForSelector.bind(page);
        jest.spyOn(page, 'waitForSelector').mockImplementation((sel, opts) =>
            typeof sel === 'string' && sel.includes('displaySign') && opts && opts.state === 'hidden'
                ? real(sel, { ...opts, timeout: ms })
                : real(sel, opts)
        );
    };

    const goto = async (html, url = APPROVAL_URL) => {
        await page.route('**/*', (route) =>
            route.fulfill({ status: 200, contentType: 'text/html', body: html })
        );
        await page.goto(url);
    };

    it('reads the number and shows it to the user', async () => {
        await goto(FLUENT);
        shortWait(300);
        const steps = [];
        const result = await waitForPhoneApproval(page, {
            logger: { ...quiet.logger, step: (s) => steps.push(String(s)) },
        });

        expect(result.shown).toBe('42');
        expect(steps.join('\n')).toMatch(/Enter the number:\s+42/);
    }, 20000);

    it('reads the legacy .displaySign shape too', async () => {
        await goto(LEGACY);
        shortWait(300);
        const result = await waitForPhoneApproval(page, quiet);
        expect(result.shown).toBe('7');
    }, 20000);

    it('WAITS — it does not click Approve or Cancel', async () => {
        // The claim under test. The screen offers both buttons; a cancel-based
        // implementation would click one and the sign-in would fail silently.
        await goto(WITH_BUTTONS);

        shortWait(2500);
        const wait = waitForPhoneApproval(page, quiet);

        await new Promise((r) => setTimeout(r, 500));
        const clicks = await page.evaluate(() => window.__clicked || 0);
        expect(clicks).toBe(0);

        // The screen clears when the user approves on their phone.
        await page.evaluate(() => {
            const el = document.querySelector('[data-testid="displaySign"]');
            if (el) el.remove();
        });

        const result = await wait;
        expect(result.waited).toBe(true);
    }, 20000);

    it('resolves without throwing when the user never answers', async () => {
        // A timeout is not an error. The caller carries on to the redirect wait,
        // which reports the real outcome rather than this function guessing.
        await goto(WITH_BUTTONS);

        shortWait(300);
        const result = await waitForPhoneApproval(page, quiet);

        expect(result.waited).toBe(false);
        expect(result.shown).toBe('42');
    }, 20000);

    it('exports the selector it matches, so the wait and login cannot drift', () => {
        // Both the legacy class and the current data-testid. If one is dropped,
        // this is where it should be noticed.
        expect(NUMBER_MATCH_SELECTOR).toContain('.displaySign');
        expect(NUMBER_MATCH_SELECTOR).toContain('[data-testid="displaySign"]');
    });
});