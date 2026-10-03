/**
 * @fileoverview Regression tests for #24 — the CLI exited 0 whatever happened.
 *
 * The bug lived in the gap between the auth module and the process, so most of
 * these assert on `process.exitCode` rather than on a function's return value:
 * asserting that login() returns false would have passed while the CLI still
 * exited 0, which is exactly the bug.
 *
 * Nothing here touches the network or launches a browser. A test that logs in
 * for real to find out it failed is slow, rate-limited, and fails for a dozen
 * reasons that have nothing to do with this change.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const path = require('path');
const os = require('os');
const fs = require('fs-extra');
const { execFile } = require('child_process');

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

const CLI = path.join(__dirname, '..', 'src', 'index.js');

/** Lets the commander's async action handlers settle before asserting. */
const flush = () => new Promise(resolve => setImmediate(resolve));

describe('verifyAuthStateFile — what "the auth file was created" means', () => {
    const { verifyAuthStateFile } = require('../src/auth');
    let dir;
    let file;

    beforeEach(() => {
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'webauth-exitcode-'));
        file = path.join(dir, 'auth-file.json');
    });

    afterEach(() => {
        fs.rmSync(dir, { recursive: true, force: true });
    });

    // The issue's rule is "login succeeded AND the json auth file is
    // successfully created". A call that returned without throwing is not
    // evidence a file exists, which is the half that was never checked.
    it('rejects a file that was never written', async () => {
        const result = await verifyAuthStateFile(file);
        expect(result.ok).toBe(false);
        expect(result.reason).toBe('missing');
    });

    it('accepts a real Playwright storage state', async () => {
        await fs.writeJson(file, {
            cookies: [{ name: 'ESAuth', value: 'x', domain: '.live.com' }],
            origins: []
        });
        const result = await verifyAuthStateFile(file);
        expect(result.ok).toBe(true);
    });

    // Truncation is the realistic version of this: a crash or a full disk
    // mid-write leaves a file that exists and does not parse.
    it('rejects a truncated file rather than trusting the write call', async () => {
        await fs.writeFile(file, '{"cookies":[{"name":"ESAuth"');
        const result = await verifyAuthStateFile(file);
        expect(result.ok).toBe(false);
        expect(result.reason).toBe('unreadable');
    });

    // Not tidiness: getAuthenticatedContext and verifyAuth both hand this path
    // straight to newContext({ storageState }). A file that parses but is not a
    // storage state would be a "successful" login that then cannot build a
    // usable context.
    it('rejects valid JSON that is not a storage state', async () => {
        await fs.writeJson(file, { hello: 'world' });
        const result = await verifyAuthStateFile(file);
        expect(result.ok).toBe(false);
        expect(result.reason).toBe('malformed');
    });
});

describe('login() reports failure instead of swallowing it', () => {
    let auth;
    let dir;
    let authFile;

    beforeEach(() => {
        jest.resetModules();
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'webauth-login-'));
        authFile = path.join(dir, 'auth-file.json');

        // A missing Chromium is the most common setup failure, and it used to
        // reject straight out of login() — past its own catch — because the
        // launch sat above the try block.
        jest.doMock('playwright', () => ({
            chromium: {
                launch: async () => { throw new Error('Chromium distribution not found'); }
            }
        }));
        jest.doMock('../src/config', () => ({
            ...jest.requireActual('../src/config'),
            DEFAULT_AUTH_FILE: path.join(dir, 'auth-file.json')
        }));

        auth = require('../src/auth');
    });

    afterEach(() => {
        fs.rmSync(dir, { recursive: true, force: true });
        jest.dontMock('playwright');
        jest.dontMock('../src/config');
    });

    // Throwing would have been the obvious fix for a CLI and the wrong one here:
    // this is also the package main, and ms-onenote-exporter awaits it.
    it('resolves false rather than rejecting when the browser cannot launch', async () => {
        await expect(auth.login({ email: 'a@b.example', password: 'x' }))
            .resolves.toBe(false);
    });

    // Otherwise the next `check` would report success off a login that failed.
    it('leaves no usable auth file behind when the login fails', async () => {
        await auth.login({ email: 'a@b.example', password: 'x' });
        expect(fs.existsSync(authFile)).toBe(false);
    });
});

describe('CLI exit code wiring', () => {
    let savedArgv;
    let authMock;

    beforeEach(() => {
        savedArgv = process.argv;
        authMock = {
            login: jest.fn(async () => false),
            verifyAuth: jest.fn(async () => ({
                authenticated: false, reason: 'no_auth_file', detail: 'no file'
            })),
            getAuthMeta: jest.fn(async () => null),
            logout: jest.fn(async () => { }),
            checkAuth: jest.fn(async () => false)
        };
    });

    afterEach(() => {
        process.argv = savedArgv;
        process.exitCode = 0;
        jest.dontMock('../src/auth');
        jest.resetModules();
    });

    /**
     * Requires src/index.js with process.argv pointed at `argv` and the auth
     * module mocked, then waits for the command's action handler to settle.
     *
     * index.js calls program.parse() at load time and its handlers are async,
     * so the exit code is only set some microtasks later — reading it straight
     * after the require would read the previous run's value.
     */
    async function runCli(argv, authOverrides = {}) {
        jest.resetModules();
        jest.doMock('../src/auth', () => ({ ...authMock, ...authOverrides }));
        process.argv = ['node', 'microsoft-webauth', ...argv];
        delete process.exitCode;
        require('../src/index.js');
        await flush();
        await flush();
        return process.exitCode;
    }

    it('login exits 0 when login() succeeds', async () => {
        const code = await runCli(['login', '--auth-file', '/tmp/x.json'], {
            login: jest.fn(async () => true)
        });
        expect(code).toBe(0);
    });

    it('login exits 1 when login() fails', async () => {
        const code = await runCli(['login', '--auth-file', '/tmp/x.json']);
        expect(code).toBe(1);
    });

    // The backstop: even a thrown error has to reach the exit code, or a crash
    // is one more way to report success.
    it('login exits 1 when the command throws outright', async () => {
        const code = await runCli(['login', '--auth-file', '/tmp/x.json'], {
            login: jest.fn(async () => { throw new Error('boom'); })
        });
        expect(code).toBe(1);
    });

    it('check exits 0 when the session is verified', async () => {
        const code = await runCli(['check', '--auth-file', '/tmp/x.json'], {
            verifyAuth: jest.fn(async () => ({
                authenticated: true, reason: 'authenticated', detail: 'session verified'
            })),
            getAuthMeta: jest.fn(async () => ({ email: 'a@b.example', loginTime: new Date().toISOString() }))
        });
        expect(code).toBe(0);
    });

    it('check exits 1 when there is no auth file', async () => {
        const code = await runCli(['check', '--auth-file', '/tmp/x.json']);
        expect(code).toBe(1);
    });

    it('check exits 1 when the session has expired', async () => {
        const code = await runCli(['check', '--auth-file', '/tmp/x.json'], {
            verifyAuth: jest.fn(async () => ({
                authenticated: false, reason: 'expired', detail: 'redirected to login'
            }))
        });
        expect(code).toBe(1);
    });

    // The decision that needed verifyAuth() rather than checkAuth(): checkAuth
    // still returns true here on purpose, so that a network blip does not
    // delete a valid session. An exit code cannot be hedged the same way.
    it('check exits 1 when the session cannot be verified at all', async () => {
        const code = await runCli(['check', '--auth-file', '/tmp/x.json'], {
            verifyAuth: jest.fn(async () => ({
                authenticated: false, reason: 'unverifiable', detail: 'network error'
            }))
        });
        expect(code).toBe(1);
    });

    it('logout exits 0', async () => {
        const code = await runCli(['logout', '--auth-file', '/tmp/x.json']);
        expect(code).toBe(0);
    });

    // The dump flags are only useful if they survive the trip through the
    // command line: commander hands them over as booleans, and a `--dodump`
    // that reached verifyAuth as undefined would be a flag that silently does
    // nothing, which is the exact failure the pair is documented to prevent.
    it('check passes --dodump and --screenshot down to verifyAuth', async () => {
        await runCli(['check', '--dodump', '--screenshot', '--auth-file', '/tmp/x.json']);

        expect(authMock.verifyAuth).toHaveBeenCalledWith(expect.objectContaining({
            authFilePath: '/tmp/x.json',
            dodump: true,
            screenshot: true
        }));
    });

    // Same rule as login's: a screenshot of a page that was never dumped is
    // not a screenshot, so the one flag turns on the other.
    it('check turns on --dodump when only --screenshot is given', async () => {
        await runCli(['check', '--screenshot', '--auth-file', '/tmp/x.json']);

        expect(authMock.verifyAuth).toHaveBeenCalledWith(
            expect.objectContaining({ dodump: true, screenshot: true })
        );
        // runCli() resets the module registry, so the logger index.js warned
        // through has to be asked for again — a reference captured at file scope
        // would be a different mock object and would record nothing.
        expect(require('../src/utils/logger').warn).toHaveBeenCalledWith(
            expect.stringContaining('--dodump')
        );
    });

    // And a plain `check` — the one that runs unattended in CI — must not start
    // writing dumps.
    it('check passes the dump flags as false when neither is given', async () => {
        await runCli(['check', '--auth-file', '/tmp/x.json']);

        expect(authMock.verifyAuth).toHaveBeenCalledWith(
            expect.objectContaining({ dodump: false, screenshot: false })
        );
    });
});

describe('the shipped binary, end to end', () => {
    let dir;

    beforeEach(() => {
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'webauth-cli-'));
    });

    afterEach(() => {
        fs.rmSync(dir, { recursive: true, force: true });
    });

    function runCli(args, { timeout = 30000 } = {}) {
        return new Promise(resolve => {
            execFile(
                process.execPath,
                [CLI, ...args],
                // The log dir is redirected so a real run cannot write into the
                // checkout, and --auth-file always points into a temp dir so
                // nothing here can read or destroy a developer's real session.
                { timeout, env: { ...process.env, ONENOTE_EXPORT_LOG_DIR: dir } },
                (error, stdout, stderr) => {
                    resolve({
                        // `error` is null on a clean exit, and a string code
                        // means the process failed to spawn at all rather than
                        // exiting non-zero. Both are -1 here so a harness
                        // problem cannot masquerade as a passing exit code.
                        code: error === null ? 0 : (typeof error.code === 'number' ? error.code : -1),
                        stdout,
                        stderr
                    });
                }
            );
        });
    }

    // The reported bug, on the real binary: "check" logged an error and exited
    // 0, so `microsoft-webauth check && <next step>` ran the next step.
    it('check exits 1 when there is no auth file', async () => {
        const { code, stderr } = await runCli(['check', '--auth-file', path.join(dir, 'missing.json')]);
        expect(stderr).toContain('Not authenticated');
        expect(code).toBe(1);
    });

    it('logout exits 0', async () => {
        const { code } = await runCli(['logout', '--auth-file', path.join(dir, 'auth-file.json')]);
        expect(code).toBe(0);
    });
});