const fs = require('fs');
const { chromium } = require('playwright');

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

const { reachPasswordScreen, waitForAuthSuccessProbe } = require('../src/auth');

/**
 * These tests replay a *work/school* account sign-in, using markup taken from the
 * real dumps of a successful login that the tool nevertheless reported as failed
 * (src/logs/dumps/2026-09-28_00h54). They are stubbed over page.route so nothing
 * touches the network and the suite is deterministic.
 *
 * The consumer account (login.live.com, Fluent UI) is covered in
 * sign-in-method.test.js; this suite covers the other half of the world.
 */
const describeWithBrowser = (() => {
    let executable = null;
    try {
        executable = chromium.executablePath();
    } catch (_) {
        // playwright could not resolve a path at all
    }
    if (!executable || !fs.existsSync(executable)) {
        console.warn('Skipping work-account tests: chromium is not installed. Run `npx playwright install chromium`.');
        return describe.skip;
    }
    return describe;
})();

const PAGE$ = (title, body) => `<!DOCTYPE html><html><head><title>${title}</title></head><body>${body}</body></html>`;

const PASSWORD_URL = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?stage=1';
const SUBMIT_URL = 'https://login.microsoftonline.com/common/login?stage=2';
const KMSI_URL = 'https://login.microsoftonline.com/common/login?stage=3';

// The work-account password page. The regression: the username field survives as
// an off-screen aria-hidden input, so it reads as "visible" to a geometry test and
// the page looks like the email step even though the password box is right there.
// Both the class and aria-hidden are copied from the real dump.
const workPasswordPage = PAGE$('Sign in to your account', `
    <h1>Enter password</h1>
    <div>
        <input type="text" name="loginfmt" data-bind="moveOffScreen, value: unsafe_displayName"
            class="moveOffScreen" tabindex="-1" aria-hidden="true">
        <input name="passwd" type="password" id="i0118" class="form-control input ext-input"
            aria-required="true">
    </div>
    <input type="submit" id="idSIButton9" value="Sign in"
        onclick="__record('Sign in'); location.href='${SUBMIT_URL}'">
`);

// "Stay signed in?" on a work account, then the authenticated app.
const kmsiPage = PAGE$('Sign in to your account', `
    <h1>Stay signed in?</h1>
    <div id="KmsiDescription">Do you want to stay signed in?</div>
    <input type="checkbox" name="DontShowAgain" id="KmsiCheckboxField">
    <input type="submit" id="idSIButton9" value="Yes"
        onclick="__record('Yes'); location.href='https://onenote.cloud.microsoft/copilotnotebooks'">
`);

// The authenticated OneNote app after the Microsoft 365 Copilot rebrand. The old
// check looked for "/notebooks", and "/copilotnotebooks" does not contain it.
const copilotNotebooks = PAGE$('OneNote', `
    <div>OneNote</div>
    <nav>Search Copilot Notebooks Recent</nav>
    <div>All Notebooks</div>
    <div>John PIGERET JP</div>
`);

// The unauthenticated marketing page the URL check must never accept.
const marketingPage = PAGE$('OneNote', `
    <h1>OneNote for the web</h1>
    <a href="/en-us">Sign up for free</a>
    <div>Get started with OneNote</div>
`);

// Some layouts keep the username field genuinely on screen next to the password
// box ("Use a different account" style). Here the email field is *not*
// aria-hidden, so this case exercises the check ordering rather than the
// aria-hidden detection, and fails if passwordField is tested after the
// "still on the email step" gate.
const visibleUsernamePage = PAGE$('Sign in to your account', `
    <h1>Enter password</h1>
    <div>
        <input type="text" name="loginfmt" value="john@mobilutils.eu" style="width:300px;height:30px">
        <input name="passwd" type="password" id="i0118" style="width:300px;height:30px">
    </div>
    <input type="submit" id="idSIButton9" value="Sign in" onclick="__record('Sign in')">
`);

// The rebrand URL with nothing rendered yet. OneNote is a JS single-page app, so
// the URL changes to the authenticated app *before* the DOM paints — the real
// timeout dump had heading "(none)" and 334 KB of shell. This is the only signal
// available in that window, which is why the URL match has to stand on its own
// rather than leaning on a text marker that may not have rendered.
const bareCopilotShell = PAGE$('OneNote', '<div id="root"></div>');

describeWithBrowser('work/school account sign-in', () => {
    let browser;
    let context;
    let page;
    let clicks;

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
        const byUrl = new Map([
            [PASSWORD_URL, workPasswordPage],
            [SUBMIT_URL, kmsiPage],
            [KMSI_URL, kmsiPage]
        ]);
        await page.route('**/*', route => {
            const url = route.request().url();
            const body = byUrl.get(url)
                || (url.startsWith('https://onenote.cloud.microsoft/copilotnotebooks') ? copilotNotebooks
                    : url.startsWith('https://onenote.cloud.microsoft') ? marketingPage
                        : PAGE$('Unknown', '<h1>Unknown</h1>'));
            return route.fulfill({ contentType: 'text/html', body });
        });
        await page.exposeFunction('__record', label => clicks.push(label));
    });

    afterEach(async () => {
        await context.close();
    });

    const options = extra => ({ stateTimeout: 3000, transitionTimeout: 3000, ...extra });

    it('recognises the password box even though the email field lingers', async () => {
        await page.goto(PASSWORD_URL);

        const result = await reachPasswordScreen(page, options());

        // The regression: this returned "unreadable" and burned the full
        // stateTimeout, because the off-screen loginfmt made the page look like
        // it had not left the email step.
        expect(result.reached).toBe(true);
        expect(result.reason).toBe('password_field');
        expect(result.state.passwordField).toBe(true);
        expect(result.state.emailField).toBe(false);
        expect(clicks).toEqual([]);
    }, 30000);

    it('reaches the password box when a real username field is on screen too', async () => {
        await page.route('**/*', route => route.fulfill({ contentType: 'text/html', body: visibleUsernamePage }));
        await page.goto(`${PASSWORD_URL}&visible=1`);

        const result = await reachPasswordScreen(page, options());

        // Here emailField really is true, so only the ordering of the two checks
        // can save this: the password box is proof the email step is done,
        // whatever else the page still shows.
        expect(result.state.emailField).toBe(true);
        expect(result.reached).toBe(true);
        expect(result.reason).toBe('password_field');
    }, 30000);

    it('accepts the rebrand path /copilotnotebooks as authenticated', async () => {
        await page.goto('https://onenote.cloud.microsoft/copilotnotebooks');

        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 5000))
            .resolves.toBe(true);
    }, 30000);

    it('accepts the rebrand URL on its own, before the SPA has rendered', async () => {
        await page.route('**/*', route => route.fulfill({ contentType: 'text/html', body: bareCopilotShell }));
        await page.goto('https://onenote.cloud.microsoft/copilotnotebooks');

        // No signed-in text is present, so this can only pass via the URL check.
        // The text marker must not be what carries the rebrand path.
        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 5000))
            .resolves.toBe(true);
    }, 30000);

    it('still accepts the old /notebooks path', async () => {
        await page.route('**/*', route => route.fulfill({ contentType: 'text/html', body: copilotNotebooks }));
        await page.goto('https://onenote.cloud.microsoft/notebooks');

        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 5000))
            .resolves.toBe(true);
    }, 30000);

    it('accepts the old /notebooks URL on its own, before the SPA has rendered', async () => {
        await page.route('**/*', route => route.fulfill({ contentType: 'text/html', body: bareCopilotShell }));
        await page.goto('https://onenote.cloud.microsoft/notebooks');

        // The twin of the rebrand case above, and the one that was missing. The
        // rendered-app test just above passes on the "All Notebooks" text
        // marker, so deleting the /notebooks branch from the URL check left the
        // whole suite green — dc04e9f found and closed that gap for the rebrand
        // path and not for this one. Confirmed by removing the branch: only
        // this test fails.
        //
        // No signed-in text here, so the URL is the only thing that can satisfy
        // the probe, and this is what makes both branches of ONENOTE_APP_PATH
        // load-bearing.
        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 5000))
            .resolves.toBe(true);
    }, 30000);

    it('never accepts the unauthenticated marketing page', async () => {
        await page.goto('https://onenote.cloud.microsoft/en-us');

        // The bare-hostname check this replaced matched this page, and caused
        // premature auth saving.
        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 2000))
            .resolves.toBe(false);
    }, 30000);

    it('does not accept a real login page as success', async () => {
        await page.goto('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');

        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 2000))
            .resolves.toBe(false);
    }, 30000);

    // The next two pin *where* in the URL the path has to appear. Matching the
    // serialised URL means a notebooks path anywhere satisfies the check — in a
    // query parameter, on any host. Both of these are unauthenticated pages that
    // the check must reject, and both carry the path in a place that is not a
    // path. The marketing page is the one that already caused premature auth
    // saving once, so it is the one worth being strict about.
    it('ignores a notebooks path carried in the query string', async () => {
        await page.goto('https://onenote.cloud.microsoft/en-us?next=/copilotnotebooks');

        // The body is the marketing page; only the query string names the app.
        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 2000))
            .resolves.toBe(false);
    }, 30000);

    it('ignores a notebooks path carried in a redirect parameter on a login host', async () => {
        await page.goto('https://login.microsoftonline.com/common/oauth2/v2.0/authorize?returnUrl=/notebooks');

        await expect(waitForAuthSuccessProbe(page, 'https://onenote.cloud.microsoft/notebooks', 2000))
            .resolves.toBe(false);
    }, 30000);
});
