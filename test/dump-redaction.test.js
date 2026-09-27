const fs = require('fs');
const os = require('os');
const path = require('path');

jest.mock('../src/utils/logger', () => {
    // Required inside the factory: jest.mock factories may not close over
    // out-of-scope variables, so a shared temp dir has to be built here.
    const nodeFs = require('fs');
    const nodePath = require('path');
    const dir = nodeFs.mkdtempSync(nodePath.join(require('os').tmpdir(), 'dump-redaction-'));
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

const logger = require('../src/utils/logger');

/**
 * These tests drive a real browser, so a missing chromium must skip rather than
 * fail the release build.
 */
const describeWithBrowser = (() => {
    let executable = null;
    try {
        executable = require('playwright').chromium.executablePath();
    } catch (_) {
        // playwright could not resolve a path at all
    }
    if (!executable || !fs.existsSync(executable)) {
        console.warn('Skipping dump-redaction tests: chromium is not installed. Run `npx playwright install chromium`.');
        return describe.skip;
    }
    return describe;
})();

const SECRET = 'AnamazingPass4Microsoft@!';
const PPFT = 'cHQeYzI1NS5TLkJqc09XTjFOR0ZaR1JEWUJB';
const EMAIL = 'msout@phttp.com';

const page$ = (title, body) => `<!DOCTYPE html><html><head><title>${title}</title></head><body>${body}</body></html>`;

// The live "Enter your password" page, with the values page.content() would
// otherwise serialise. Markup taken from the real debug_after_password.html dump.
const passwordPage = page$('Enter your password', `
    <h1 data-testid="title">Enter your password</h1>
    <form method="post" action="https://login.live.com/ppsecure/post.srf">
        <input type="email" name="loginfmt" value="${EMAIL}">
        <input type="hidden" name="PPFT" value="${PPFT}">
        <input type="hidden" name="i13" value="0">
        <input type="password" name="passwd" id="passwordEntry" autocomplete="current-password" value="${SECRET}">
        <input type="text" name="otc" value="123456">
        <textarea name="otcFallback">654321</textarea>
        <button type="submit" data-testid="primaryButton">Sign in</button>
    </form>
`);

const PASSWORD_URL = 'https://login.live.com/oauth20_authorize.srf?stage=1';

describeWithBrowser('dumpPage', () => {
    let chromium;
    let browser;
    let context;
    let page;
    let dumpDir;

    beforeAll(async () => {
        chromium = require('playwright').chromium;
        browser = await chromium.launch({ headless: true });
    });

    afterAll(async () => {
        await browser.close();
    });

    beforeEach(async () => {
        dumpDir = await logger.getDumpDir();
        context = await browser.newContext();
        page = await context.newPage();
        await page.route('**/*', route =>
            route.fulfill({ contentType: 'text/html', body: passwordPage }));
        await page.goto(PASSWORD_URL);
    });

    afterEach(async () => {
        await context.close();
    });

    const readDump = async name => fs.readFileSync(path.join(dumpDir, name), 'utf8');

    it('never writes the password to the dump', async () => {
        const { dumpPage } = require('../src/auth');

        const displayPath = await dumpPage(page, 'redaction_password.html');
        const dump = await readDump('redaction_password.html');

        expect(displayPath).toBe('logs/dumps/test');
        expect(dump).not.toContain(SECRET);
        expect(dump).toContain('name="passwd"');
        expect(dump).toContain('[redacted]');
        // Guards the placeholder actually arriving in the page: a redactor that
        // forgets to pass its argument to page.evaluate writes the literal
        // string "undefined" over the secret and looks like it worked.
        expect(dump).not.toContain('value="undefined"');
        expect(dump).not.toContain('>undefined<');
    }, 30000);

    it('redacts the flow token and the one-time-code fields', async () => {
        const { dumpPage } = require('../src/auth');

        await dumpPage(page, 'redaction_tokens.html');
        const dump = await readDump('redaction_tokens.html');

        // PPFT is the pre-auth flow token the form is about to POST.
        expect(dump).not.toContain(PPFT);
        expect(dump).not.toContain('123456');
        expect(dump).not.toContain('654321');
    }, 30000);

    it('keeps the non-secret UI state that makes a dump useful', async () => {
        const { dumpPage } = require('../src/auth');

        await dumpPage(page, 'redaction_kept.html');
        const dump = await readDump('redaction_kept.html');

        // Over-redaction makes a debug dump worthless, so the parts that identify
        // *which* screen and *which* account failed must survive.
        expect(dump).toContain('Enter your password');
        expect(dump).toContain(EMAIL);
        expect(dump).toContain('PPFT');
        expect(dump).toContain('value="0"');   // the i13 state flag
        expect(dump).toContain('Sign in');     // the submit button
    }, 30000);

    it('leaves the live page untouched, so the login can still be submitted', async () => {
        const { dumpPage } = require('../src/auth');

        // Redacting in the live DOM would blank the very values the form posts
        // and break the login being debugged. This is the regression guard.
        await dumpPage(page, 'redaction_live.html');

        const live = await page.evaluate(() => ({
            passwd: document.querySelector('input[name="passwd"]').value,
            ppft: document.querySelector('input[name="PPFT"]').value,
            otc: document.querySelector('input[name="otc"]').value
        }));

        expect(live.passwd).toBe(SECRET);
        expect(live.ppft).toBe(PPFT);
        expect(live.otc).toBe('123456');
    }, 30000);

    it('does not crash on a page with nothing to redact', async () => {
        const { dumpPage } = require('../src/auth');
        await page.route('**/*', route => route.fulfill({
            contentType: 'text/html',
            body: page$('My notebooks', '<h1>My notebooks</h1>')
        }));
        await page.goto(`${PASSWORD_URL}&clean=1`);

        await dumpPage(page, 'redaction_clean.html');

        expect(await readDump('redaction_clean.html')).toContain('My notebooks');
    }, 30000);
});
