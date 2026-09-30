/**
 * @fileoverview Screenshots written alongside the --dodump HTML dumps.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const fs = require('fs');
const path = require('path');

const { chromium, describeWithBrowser } = require('./helpers/browser-suite');

jest.mock('../src/utils/logger', () => {
    // Required inside the factory: jest.mock factories may not close over
    // out-of-scope variables, so a shared temp dir has to be built here.
    const nodeFs = require('fs');
    const nodePath = require('path');
    const dir = nodeFs.mkdtempSync(nodePath.join(require('os').tmpdir(), 'dump-screenshot-'));
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
const { clearBlockingScreens, dumpPage, reachPasswordScreen } = require('../src/auth');

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

/** Width and height out of the PNG IHDR chunk, which starts at byte 8. */
const pngSize = buffer => ({
    width: buffer.readUInt32BE(16),
    height: buffer.readUInt32BE(20)
});

/** Taller than the viewport below, so a full-page capture is distinguishable. */
const TALL_PAGE = `<!DOCTYPE html><html><head><title>Enter your password</title></head>
<body style="margin:0">
    <h1 data-testid="title">Enter your password</h1>
    <input type="email" name="loginfmt" value="msout@phttp.com">
    <input type="password" name="passwd" id="passwordEntry" value="not-in-the-html">
    <div style="height:3000px;background:#eee">below the fold</div>
</body></html>`;

const VIEWPORT = { width: 1280, height: 720 };

/** A screen that hijacks the navigation mid-login, and what it leads to. */
const ONENOTE_STUB_URL = 'https://onenote.cloud.microsoft/notebooks';
const TOU_URL = 'https://account.live.com/tou/accrue?mkt=EN-US';
const SEND_CODE_URL = 'https://login.live.com/ppsecure/post.srf?id=100';
const PASSWORD_URL = 'https://login.live.com/ppsecure/post.srf?id=300';
const touPage = `<!DOCTYPE html><html><body><h1>We're updating our terms</h1>
    <button type="submit" data-testid="primaryButton"
        onclick="location.href='${ONENOTE_STUB_URL}'">Next</button></body></html>`;

/** The passwordless-first screen: the password is one footer link away. */
const sendCodePage = `<!DOCTYPE html><html><body><h1 data-testid="title">Get a code to sign in</h1>
    <button type="submit" data-testid="primaryButton">Send code</button>
    <span role="button" class="fui-Link" tabindex="0"
        onclick="location.href='${PASSWORD_URL}'">Use your password</span></body></html>`;
const passwordPage = `<!DOCTYPE html><html><body><h1 data-testid="title">Enter your password</h1>
    <input type="password" name="passwd"></body></html>`;

const byUrl = new Map([
    [TOU_URL, touPage],
    [SEND_CODE_URL, sendCodePage],
    [PASSWORD_URL, passwordPage]
]);

describeWithBrowser('dumpPage --screenshot', () => {
    let browser;
    let context;
    let page;
    let dumpDir;

    beforeAll(async () => {
        browser = await chromium.launch({ headless: true });
    });

    afterAll(async () => {
        await browser.close();
    });

    beforeEach(async () => {
        jest.clearAllMocks();
        dumpDir = await logger.getDumpDir();
        context = await browser.newContext({ viewport: VIEWPORT });
        page = await context.newPage();
        await page.route('**/*', route => route.fulfill({
            contentType: 'text/html',
            body: byUrl.get(route.request().url()) || TALL_PAGE
        }));
        await page.goto('https://login.live.com/oauth20_authorize.srf?stage=1');
    });

    afterEach(async () => {
        await context.close();
    });

    const exists = name => fs.existsSync(path.join(dumpDir, name));

    /** Drops a dump and its screenshot, so a test can assert on a fresh pair. */
    const removeDump = name => {
        fs.rmSync(path.join(dumpDir, `${name}.html`), { force: true });
        fs.rmSync(path.join(dumpDir, `${name}.png`), { force: true });
    };

    it('writes a PNG beside the HTML dump, under the same basename', async () => {
        await dumpPage(page, 'shot_pair.html', { screenshot: true });

        expect(exists('shot_pair.html')).toBe(true);
        expect(exists('shot_pair.png')).toBe(true);
    }, 30000);

    it('writes a real PNG, not a truncated or empty file', async () => {
        await dumpPage(page, 'shot_real.html', { screenshot: true });

        const png = fs.readFileSync(path.join(dumpDir, 'shot_real.png'));
        expect(png.length).toBeGreaterThan(1000);
        expect(png.subarray(0, 8)).toEqual(PNG_SIGNATURE);
    }, 30000);

    it('captures the whole page, not just the visible viewport', async () => {
        await dumpPage(page, 'shot_fullpage.html', { screenshot: true });

        // A viewport-only capture of this page is exactly VIEWPORT tall, so a
        // taller image proves fullPage for real rather than by assertion on the
        // options object.
        const { height } = pngSize(fs.readFileSync(path.join(dumpDir, 'shot_fullpage.png')));
        expect(height).toBeGreaterThan(VIEWPORT.height);
    }, 30000);

    it('takes no screenshot unless one was asked for', async () => {
        await dumpPage(page, 'shot_off.html');

        expect(exists('shot_off.html')).toBe(true);
        expect(exists('shot_off.png')).toBe(false);
        expect(logger.warn).not.toHaveBeenCalled();
    }, 30000);

    it('writes the HTML dump even when the page is gone before the capture', async () => {
        // The timing this guards is real, not hypothetical: a dump taken on a
        // screen that hijacks the navigation races that navigation, and the
        // dump is most wanted precisely when the login is about to fail.
        await context.close();

        await expect(dumpPage(page, 'shot_gone.html', { screenshot: true }))
            .resolves.toBe('logs/dumps/test');

        expect(exists('shot_gone.html')).toBe(true);
        expect(logger.warn).toHaveBeenCalledWith(
            expect.stringContaining('shot_gone.png')
        );
    }, 30000);

    it('screenshots the blocking screens it dumps', async () => {
        await page.goto(TOU_URL);

        const result = await clearBlockingScreens(page, {
            progress: { signatures: new Set(), lastClickAt: 0 },
            stateTimeout: 3000,
            changeTimeout: 3000,
            dodump: true,
            screenshot: true
        });

        // The dump is taken by the blocking-screen handler, not by login() itself,
        // so this is what catches the flag being dropped on the way down.
        expect(result.handled).toBe(1);
        expect(exists('debug_blocking_screen_1.html')).toBe(true);
        expect(exists('debug_blocking_screen_1.png')).toBe(true);
    }, 30000);

    it('screenshots the intermediate sign-in screen it dumps', async () => {
        await page.goto(SEND_CODE_URL);

        const nav = await reachPasswordScreen(page, {
            stateTimeout: 3000,
            transitionTimeout: 3000,
            dodump: true,
            screenshot: true
        });

        expect(nav.reached).toBe(true);
        expect(exists('debug_intermediate_screen.html')).toBe(true);
        expect(exists('debug_intermediate_screen.png')).toBe(true);
    }, 30000);

    it('leaves no screenshot behind when only --dodump was asked for', async () => {
        await page.goto(TOU_URL);
        // The dump directory is shared by the whole suite and the blocking-screen
        // dump name is fixed in the implementation, so the screenshot the test
        // above wrote has to go before "no screenshot" can mean anything.
        removeDump('debug_blocking_screen_1');

        await clearBlockingScreens(page, {
            progress: { signatures: new Set(), lastClickAt: 0 },
            stateTimeout: 3000,
            changeTimeout: 3000,
            dodump: true
        });

        expect(exists('debug_blocking_screen_1.html')).toBe(true);
        expect(exists('debug_blocking_screen_1.png')).toBe(false);
    }, 30000);
});

describe('dumpPage --screenshot: capture options and failure handling', () => {
    let dumpDir;

    /** A page stub: the HTML dump needs evaluate(), the screenshot needs screenshot(). */
    const stubPage = (screenshot) => ({
        evaluate: jest.fn().mockResolvedValue('<!DOCTYPE html><html><body>stub</body></html>'),
        screenshot
    });

    beforeEach(async () => {
        jest.clearAllMocks();
        dumpDir = await logger.getDumpDir();
    });

    it('asks for a full-page capture with a bounded timeout', async () => {
        const page = stubPage(jest.fn().mockResolvedValue(Buffer.alloc(8)));

        await dumpPage(page, 'stub_options.html', { screenshot: true });

        expect(page.screenshot).toHaveBeenCalledTimes(1);
        const options = page.screenshot.mock.calls[0][0];
        expect(options.fullPage).toBe(true);
        expect(options.path).toBe(path.join(dumpDir, 'stub_options.png'));
        // A screenshot that can outlive the waits it sits inside would make a
        // dump the slowest part of a login that is already late.
        expect(options.timeout).toBeGreaterThan(0);
        expect(options.timeout).toBeLessThanOrEqual(30000);
    });

    it('does not fail the dump when the capture itself fails', async () => {
        const page = stubPage(jest.fn().mockRejectedValue(new Error('Target closed')));

        await expect(dumpPage(page, 'stub_failed.html', { screenshot: true }))
            .resolves.toBe('logs/dumps/test');

        // The HTML is the point of the dump; a missing PNG must not cost it.
        expect(fs.readFileSync(path.join(dumpDir, 'stub_failed.html'), 'utf8')).toContain('stub');
        expect(logger.warn).toHaveBeenCalledWith(expect.stringContaining('stub_failed.png'));
    });

    it('names the PNG after the dump even without an .html extension', async () => {
        const page = stubPage(jest.fn().mockResolvedValue(Buffer.alloc(8)));

        await dumpPage(page, 'stub_noext', { screenshot: true });

        expect(page.screenshot.mock.calls[0][0].path).toBe(path.join(dumpDir, 'stub_noext.png'));
    });
});

describe('every --dodump write site takes a screenshot', () => {
    /**
     * Screenshots are taken inside dumpPage(), but only for the callers that ask
     * for them, and the eight call sites in login(), clearBlockingScreens() and
     * reachPasswordScreen() are the ones that know a dump is happening. One of
     * them forgetting to pass the flag on would silently drop a screenshot from
     * exactly one dump — the sort of gap that is only noticed while debugging
     * the dump that is missing it.
     */
    it('passes { screenshot } at every dumpPage() call in src/auth.js', () => {
        const source = fs.readFileSync(
            path.join(__dirname, '..', 'src', 'auth.js'), 'utf8');

        const calls = source.match(/await dumpPage\([\s\S]*?\);/g) || [];

        // A regex that quietly matched nothing would pass this test for ever.
        expect(calls.length).toBeGreaterThanOrEqual(8);
        expect(calls.filter(call => !call.includes('screenshot'))).toEqual([]);
    });
});