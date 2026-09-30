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

const { reachPasswordScreen, submitSignInForm } = require('../src/auth');


const CODE_URL = 'https://login.live.com/ppsecure/post.srf?id=100';
const METHODS_URL = 'https://login.live.com/ppsecure/post.srf?id=200';
const PASSWORD_URL = 'https://login.live.com/ppsecure/post.srf?id=300';
const APPROVAL_URL = 'https://login.live.com/ppsecure/post.srf?id=400';
const LEGACY_URL = 'https://login.live.com/ppsecure/post.srf?id=500';
const EMAIL_URL = 'https://login.live.com/ppsecure/post.srf?id=600';
const OTHER_WAYS_URL = 'https://login.live.com/ppsecure/post.srf?id=700';
const CODE_ONLY_URL = 'https://login.live.com/ppsecure/post.srf?id=800';
// Not in the stub map: whatever navigates here proves the submit actually fired.
const NOWHERE_URL = 'https://login.live.com/ppsecure/post.srf?id=900';

const page$ = (title, body) => `<!DOCTYPE html><html><head><title>${title}</title></head><body>${body}</body></html>`;

// The regression: the passwordless-first screen. "Use your password" is a
// <span role="button">, there is NO "Other ways to sign in" link anywhere, and
// "Send code" is the primary button.
const sendCodeHtml = page$('Get a code to sign in', `
    <h1 data-testid="title">Get a code to sign in</h1>
    <div data-testid="subtitle">We'll send a code to msout@phttp.com to sign you in.</div>
    <button type="submit" data-testid="primaryButton"
        onclick="__record('Send code')">Send code</button>
    <span data-testid="viewFooter"><div><span role="button" class="fui-Link" tabindex="0"
        onclick="__record('Use your password'); location.href='${PASSWORD_URL}'">Use your password</span></div></span>
    <button type="button" id="close-button" data-testid="dismissIcon"
        onclick="__record('Dismiss')">Close</button>
`);

// The older shape: an explicit "Other ways to sign in" step, then a method list.
const otherWaysHtml = page$('Verify your identity', `
    <h1>Verify your identity</h1>
    <input type="tel" name="otc">
    <button type="submit" onclick="__record('Next')">Next</button>
    <a href="#" id="signInAnotherWay" onclick="__record('Other ways'); location.href='${METHODS_URL}'">Other ways to sign in</a>
`);

const methodListHtml = page$('Sign in to your Microsoft account', `
    <h1>Select a sign-in method</h1>
    <div><button type="button" data-testid="primaryButton"
        onclick="__record('Password'); location.href='${PASSWORD_URL}'">Password</button></div>
    <div><button type="button" onclick="__record('Phone')">Use a phone number</button></div>
    <div><a href="#" onclick="__record('Email')">Use an email address</a></div>
`);

// The legacy server-rendered password form: label in `value`, not in text.
const legacyPasswordHtml = page$('Sign in to your Microsoft account', `
    <h1>Sign in to your account</h1>
    <input type="password" name="passwd" id="i0327">
    <input type="submit" id="idSIButton9" value="Sign in" onclick="__record('Sign in'); location.href='${NOWHERE_URL}'">
`);

// The Fluent password form: label in text.
const passwordHtml = page$('Sign in to your Microsoft account', `
    <h1 data-testid="title">Sign in</h1>
    <input type="password" name="passwd" id="passwordInput">
    <button type="submit" data-testid="primaryButton"
        onclick="__record('Sign in'); location.href='${NOWHERE_URL}'">Sign in</button>
`);

// An account that will only accept an authenticator approval: no password route
// exists, and the tool must say so instead of blaming the password.
const approvalHtml = page$('Sign in to your Microsoft account', `
    <h1>Approve sign in request</h1>
    <div data-testid="displaySign">42</div>
    <p>Open Microsoft Authenticator and enter the number.</p>
`);

// A screen with no route to a password at all.
const codeOnlyHtml = page$('Get a code to sign in', `
    <h1 data-testid="title">Get a code to sign in</h1>
    <button type="submit" data-testid="primaryButton" onclick="__record('Send code')">Send code</button>
`);

const emailHtml = page$('Sign in', `
    <h1>Sign in</h1>
    <input type="email" name="loginfmt" id="i0116">
    <input type="submit" id="idSIButton9" value="Next">
`);

describeWithBrowser('reachPasswordScreen', () => {
    let browser;
    let context;
    let page;
    let clicks;

    async function stubPages() {
        const byUrl = new Map([
            [EMAIL_URL, emailHtml],
            [CODE_URL, sendCodeHtml],
            [OTHER_WAYS_URL, otherWaysHtml],
            [METHODS_URL, methodListHtml],
            [PASSWORD_URL, passwordHtml],
            [LEGACY_URL, legacyPasswordHtml],
            [APPROVAL_URL, approvalHtml],
            [CODE_ONLY_URL, codeOnlyHtml]
        ]);

        await page.route('**/*', route => {
            const body = byUrl.get(route.request().url()) || page$('Unknown', '<h1>Unknown</h1>');
            return route.fulfill({ contentType: 'text/html', body });
        });
        await page.exposeFunction('__record', label => clicks.push(label));
    }

    // Fast timings: the real 15 s/10 s waits would dominate the suite.
    const options = extra => ({ stateTimeout: 3000, transitionTimeout: 3000, ...extra });

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

    it('reaches the password box from the passwordless "Get a code" screen', async () => {
        await page.goto(CODE_URL);

        const result = await reachPasswordScreen(page, options());

        expect(result.reached).toBe(true);
        expect(result.reason).toBe('password_field');
        // The whole point: "Use your password" is pressed, and "Send code" — which
        // would start a real MFA challenge — is never pressed.
        expect(clicks).toEqual(['Use your password']);
        expect(clicks).not.toContain('Send code');
    }, 30000);

    it('goes through the method list when only "Other ways to sign in" is offered', async () => {
        await page.goto(OTHER_WAYS_URL);

        const result = await reachPasswordScreen(page, options());

        expect(result.reached).toBe(true);
        // "Next" on the code field is not a password route, so only the two
        // password-sign-in actions may be pressed.
        expect(clicks).toEqual(['Other ways', 'Password']);
    }, 30000);

    it('reports a code-only screen instead of blaming the password', async () => {
        await page.goto(CODE_ONLY_URL);

        const result = await reachPasswordScreen(page, options());

        expect(result.reached).toBe(false);
        expect(result.reason).toBe('code_prompt');
        expect(result.state.heading).toBe('Get a code to sign in');
    }, 30000);

    it('reports an authenticator-approval screen rather than a missing password', async () => {
        await page.goto(APPROVAL_URL);

        const result = await reachPasswordScreen(page, options());

        expect(result.reached).toBe(false);
        expect(result.reason).toBe('approver_prompt');
    }, 30000);

    it('waits for the email step to hand over before deciding', async () => {
        await page.goto(EMAIL_URL);

        const result = await reachPasswordScreen(page, options({ stateTimeout: 1500, transitionTimeout: 1500 }));

        // Still on the email form: the next step never rendered, so nothing may be
        // clicked and the caller is told the screen was unreadable.
        expect(result.reached).toBe(false);
        expect(result.reason).toBe('unreadable');
        expect(clicks).toEqual([]);
    }, 30000);

    it('never clicks the dismiss/close control on the screen', async () => {
        await page.goto(CODE_URL);

        await reachPasswordScreen(page, options());

        expect(clicks).not.toContain('Dismiss');
    }, 30000);

    it('submits the legacy form whose button label lives in value, not text', async () => {
        await page.goto(LEGACY_URL);

        // filter({ hasText }) can never match an <input>; the label is in value.
        expect(await submitSignInForm(page)).toBe(true);
        expect(clicks).toEqual(['Sign in']);
    }, 30000);

    it('submits the Fluent form whose button label lives in text', async () => {
        await page.goto(PASSWORD_URL);

        expect(await submitSignInForm(page)).toBe(true);
        expect(clicks).toEqual(['Sign in']);
    }, 30000);
});
