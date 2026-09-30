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

const { reachPasswordScreen, waitForAuthSuccessProbe, waitForAuthSuccess } = require('../src/auth');


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

// The same page, but with the parked field marked aria-hidden="True". This case
// is taken from the ARIA spec rather than from a dump: Microsoft's markup uses
// lowercase, so nothing here has been observed in the wild. It is worth pinning
// because the rule that aria-hidden means "not visible" is now load-bearing for
// the whole work-account path, and a mixed-case token is the one way that rule
// can quietly stop holding.
const mixedCaseAriaHiddenPage = PAGE$('Sign in to your account', `
    <h1>Enter password</h1>
    <div>
        <input type="text" name="loginfmt" value="john@mobilutils.eu"
            class="moveOffScreen" aria-hidden="True">
        <input name="passwd" type="password" id="i0118" style="width:300px;height:30px">
    </div>
    <input type="submit" id="idSIButton9" value="Sign in" onclick="__record('Sign in')">
`);

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

    // Long enough that a suite which stalls is obvious, short enough that a
    // regression which burns the whole budget is still visible in the numbers.
    const STATE_TIMEOUT = 3000;

    const options = extra => ({ stateTimeout: STATE_TIMEOUT, transitionTimeout: 3000, ...extra });

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

    it('treats a mixed-case aria-hidden token as not visible', async () => {
        await page.route('**/*', route => route.fulfill({ contentType: 'text/html', body: mixedCaseAriaHiddenPage }));
        await page.goto(`${PASSWORD_URL}&mixedcase=1`);

        const result = await reachPasswordScreen(page, options());

        // "True" is the same token as "true" per the ARIA spec, so the parked
        // field is still not visible and the page is still the password step.
        // This passes on the *ordering* fix even if aria-hidden were ignored
        // entirely, so the assertion that matters is emailField — that is the
        // rule under test, and the reason the loop-level check cannot be the
        // only thing standing between this page and a 15 s stall.
        expect(result.state.emailField).toBe(false);
        expect(result.reached).toBe(true);
        expect(result.reason).toBe('password_field');
    }, 30000);

    it('reaches the password box when a real username field is on screen too', async () => {
        await page.route('**/*', route => route.fulfill({ contentType: 'text/html', body: visibleUsernamePage }));
        await page.goto(`${PASSWORD_URL}&visible=1`);

        const t0 = Date.now();
        const result = await reachPasswordScreen(page, options());
        const elapsed = Date.now() - t0;

        // Here emailField really is true, so only the ordering of the two checks
        // can save this: the password box is proof the email step is done,
        // whatever else the page still shows.
        expect(result.state.emailField).toBe(true);
        expect(result.reached).toBe(true);
        expect(result.reason).toBe('password_field');

        // The box is already on screen when we arrive, so this needs to wait for
        // nothing at all. It used to take the full stateTimeout: the initial read
        // accepted only isActionableSignInState, which is false whenever
        // emailField is set, so that wait could never be satisfied no matter how
        // long it ran — the password box was only noticed by the loop afterwards.
        // Half the budget is a wide margin: the stalled case lands above it at
        // ~3300 ms, the working case an order of magnitude below.
        expect(elapsed).toBeLessThan(STATE_TIMEOUT / 2);
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

    // The next two cover what a user is actually shown when none of the signals
    // arrives. Promise.any rejects with an AggregateError of bare TimeoutErrors,
    // which prints as "All promises were rejected" and names nothing — the
    // opposite of useful when the cause is a stale marker or a moved path, which
    // is what this file's two regressions were.
    it('names the signals it waited for when the authenticated app never arrives', async () => {
        await page.goto('https://onenote.cloud.microsoft/en-us');

        // The caller adds where the browser ended up; this half has to say what
        // was expected and what was being watched.
        const err = await waitForAuthSuccess(page, 'https://onenote.cloud.microsoft/notebooks', 2000)
            .then(() => null, e => e);

        expect(err).toBeInstanceOf(Error);
        expect(err).not.toBeInstanceOf(AggregateError);
        expect(err.message).toMatch(/Timed out after 2s waiting for the authenticated OneNote app/);
        // Every signal, by name, so a renamed marker is diagnosable from this.
        expect(err.message).toContain('OneNote app path');
        expect(err.message).toContain('text="My notebooks"');
        expect(err.message).toContain('text="All Notebooks"');
    }, 30000);

    it('reports the Outlook signals when an Outlook login does not land', async () => {
        await page.goto('https://onenote.cloud.microsoft/en-us');

        const err = await waitForAuthSuccess(page, 'https://outlook.cloud.microsoft/mail/', 2000)
            .then(() => null, e => e);

        expect(err).toBeInstanceOf(Error);
        expect(err.message).toMatch(/Timed out after 2s waiting for Outlook mail/);
        // The OneNote markers must not leak into the Outlook report.
        expect(err.message).toContain('Outlook marker');
        expect(err.message).not.toContain('My notebooks');
    }, 30000);
});
