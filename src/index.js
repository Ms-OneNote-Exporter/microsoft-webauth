#!/usr/bin/env node
/**
 * @fileoverview Main CLI interface to handle authentication commands.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const { program } = require('commander');
const logger = require('./utils/logger');
const { login, getAuthMeta, verifyAuth, logout } = require('./auth');
const { DEFAULT_AUTH_FILE, ONENOTE_URL, OUTLOOK_URL } = require('./config');
const { version: PKG_VERSION } = require('../package.json');

/**
 * Process exit codes.
 *
 * 0 and 1, and nothing else. A login either produced usable state or it did
 * not, and a script driving this CLI needs to be able to branch on that with
 * `if ! microsoft-webauth login ...`. Enumerating distinct failures (bad
 * credentials, timeout, network) would be more informative, but it would make
 * the *success* case the only value a caller must not hardcode, and the failure
 * case is the one people end up hardcoding. The reason is already in the log
 * and in the message on stderr.
 */
const EXIT_SUCCESS = 0;
const EXIT_FAILURE = 1;

/**
 * Runs a command's work and records the exit code.
 *
 * `work` returns a boolean and does its own reporting, so the reporting stays in
 * one place with the logic that produced it.
 *
 * `process.exitCode` is assigned rather than `process.exit()` called, because
 * the latter cuts the process off mid-flight: stdout is a pipe here often
 * enough — that is the entire use case — and a truncated last line is a mangled
 * error message. Assigning it lets Node finish writing and exit on its own,
 * while still producing exactly the code a caller checks.
 */
async function run(work) {
    let ok;
    try {
        ok = await work();
    } catch (e) {
        // Every command reports its own failure through its return value. This
        // is the backstop for anything that escapes — an unwritable auth file,
        // a full disk — so a crash is a non-zero exit rather than yet another
        // silent success.
        logger.error(e && e.message ? e.message : String(e), e instanceof Error ? e : null);
        ok = false;
    }
    process.exitCode = ok ? EXIT_SUCCESS : EXIT_FAILURE;
    return ok;
}

/** One line saying why a command failed, on stderr. */
function reportFailure(command, hint) {
    logger.error(`${command} failed (exit ${EXIT_FAILURE}). ${hint}`);
}

/**
 * Turns --screenshot on its own into a dump as well.
 *
 * Screenshots are only ever taken of a dumped page, so on their own they would
 * be an option that silently does nothing — the kind of flag a user debugs for
 * an hour. Both `login` and `check` accept the pair and mean the same thing by
 * it, so the rule lives here rather than being restated per command and left to
 * drift.
 *
 * @param {object} options  a commander's parsed options object, mutated in place
 */
function applyDumpFlags(options) {
    if (options.screenshot && !options.dodump) {
        logger.warn('--screenshot only applies to the pages written by --dodump; enabling --dodump as well.');
        options.dodump = true;
    }
}

program
    .name('webauth')
    .description('Microsoft web authentication via Playwright — extracted from MSOneNote Exporter')
    .version(PKG_VERSION);

program
    .command('login')
    .description('Authenticate with Microsoft Account')
    .option('--email <email>', 'Microsoft account email')
    .option('--password <password>', 'Microsoft account password')
    .option('--notheadless', 'Run in visible browser mode for debugging')
    .option('--dodump', 'Dump HTML content to files for debugging')
    .option('--screenshot', 'With --dodump, also save a PNG screenshot of each dumped page')
    .option('--against <target>', 'Target service: onenote (default) or outlook', 'onenote')
    .option('--auth-file <path>', 'Path to auth file (default: ~/.microsoft-webauth/auth-file.json)', DEFAULT_AUTH_FILE)
    .action(async (options) => {
        applyDumpFlags(options);
        const targetUrl = options.against === 'outlook' ? OUTLOOK_URL : ONENOTE_URL;

        // login() resolves a boolean rather than throwing: it covers both halves
        // of "logged in" — the session authenticated, and the auth file was
        // written and read back — and it is also this package's main export, so
        // throwing would break existing library callers.
        await run(async () => {
            const ok = await login({ ...options, targetUrl });
            if (!ok) {
                reportFailure('login', 'No usable auth state was saved. See the errors above.');
            }
            return ok;
        });
    });

program
    .command('check')
    .description('Check if authenticated')
    .option('--dodump', 'Dump HTML content to files for debugging')
    .option('--screenshot', 'With --dodump, also save a PNG screenshot of each dumped page')
    .option('--against <target>', 'Target service: onenote (default) or outlook', 'onenote')
    .option('--auth-file <path>', 'Path to auth file (default: ~/.microsoft-webauth/auth-file.json)', DEFAULT_AUTH_FILE)
    .action(async (options) => {
        applyDumpFlags(options);
        const targetUrl = options.against === 'outlook' ? OUTLOOK_URL : ONENOTE_URL;
        await run(async () => {
            // verifyAuth() rather than checkAuth(): checkAuth answers "is there a
            // session worth keeping", where a failed check counts as yes, and
            // that conservative default is exactly what must not leak into an
            // exit code. An unverifiable session exits 1 — the honest answer to
            // "am I logged in?" is no — while the auth file is left untouched,
            // since a check that failed is not evidence the session is bad.
            const status = await verifyAuth({
                targetUrl,
                authFilePath: options.authFile,
                dodump: !!options.dodump,
                screenshot: !!options.screenshot
            });
            logger.debug(`Check result: ${status.reason} — ${status.detail}`);

            if (!status.authenticated) {
                logger.error(`Not authenticated (${status.reason}). ${status.detail}`);
                // Every reason here is fixed by logging in again, including the
                // one that only failed to see a live session: the honest advice
                // for "check could not confirm it" is still to log in.
                logger.error('Run "login" first.');
                reportFailure('check', 'No authenticated session could be confirmed. See the reason above.');
                return false;
            }

            logger.success('Authentication file found. You are authenticated.');
            const meta = await getAuthMeta(options.authFile);
            if (meta && meta.email) {
                const loginTime = new Date(meta.loginTime).toLocaleString();
                logger.info(`Logged in as: ${meta.email}`);
                logger.debug(`Session started at: ${loginTime}`);
            }
            return true;
        });
    });

program
    .command('logout')
    .description('Clear authentication state')
    .option('--auth-file <path>', 'Path to auth file (default: ~/.microsoft-webauth/auth-file.json)', DEFAULT_AUTH_FILE)
    .action(async (options) => {
        // Not routed through run(): there is nothing here that can report a
        // failure. logout() removes both files unconditionally and never
        // throws, so this command has exactly one outcome and it is the
        // successful one — exiting 0 for "there was nothing to delete" included,
        // which is the answer a caller wants.
        await logout(options.authFile);
        logger.success('Logged out successfully. Authentication state cleared.');
        process.exitCode = EXIT_SUCCESS;
    });

program.parse();
