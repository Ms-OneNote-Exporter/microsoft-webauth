/**
 * @fileoverview This file handles user authentication logic.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const { chromium } = require('playwright');
const fs = require('fs-extra');
const logger = require('./utils/logger');
const { waitForPhoneApproval } = require('./phone-approval');
const { DEFAULT_AUTH_FILE, getAuthMetaFilePath, ensureAuthDir, ONENOTE_URL } = require('./config');
const { version: PKG_VERSION } = require('../package.json');
const path = require('path');
const readline = require('readline');

/** Returns the auth file path, defaulting to DEFAULT_AUTH_FILE */
function getAuthFilePath(authFilePath) {
    return authFilePath || DEFAULT_AUTH_FILE;
}

/** Returns { email, loginTime } from auth-meta.json, or null if not found. */
async function getAuthMeta(authFilePath) {
    const filePath = getAuthFilePath(authFilePath);
    const metaPath = getAuthMetaFilePath(filePath);
    try {
        if (await fs.pathExists(metaPath)) {
            return await fs.readJson(metaPath);
        }
    } catch (e) { }
    return null;
}

/** Generates backup path by appending .old to file */
function getBackupPath(filePath) {
    return filePath + '.old';
}

/**
 * Checks if files exist and handles backup logic
 * Returns object with { authFileExists, metaFileExists, willOverwriteOld }
 */
async function checkAndPrepareFiles(authFilePath) {
    const metaFilePath = getAuthMetaFilePath(authFilePath);

    const authFileExists = await fs.pathExists(authFilePath);
    const metaFileExists = await fs.pathExists(metaFilePath);

    // Check if .old versions exist
    const authOldPath = getBackupPath(authFilePath);
    const metaOldPath = getBackupPath(metaFilePath);
    const authOldExists = await fs.pathExists(authOldPath);
    const metaOldExists = await fs.pathExists(metaOldPath);

    let willOverwriteOld = false;

    // If current files exist, backup them
    if (authFileExists || metaFileExists) {
        // Warn if .old files already exist (they will be erased)
        if (authOldExists || metaOldExists) {
            logger.warn(`Warning: Backup files (.old) already exist and will be erased:`);
            if (authOldExists) logger.warn(`  ${authOldPath}`);
            if (metaOldExists) logger.warn(`  ${metaOldPath}`);
            willOverwriteOld = true;
        }

        // Create backup directory if it doesn't exist
        const dir = path.dirname(authFilePath);
        await ensureAuthDir(authFilePath);

        // Backup existing files
        if (authFileExists) {
            await fs.move(authFilePath, authOldPath, { overwrite: true });
            logger.info(`Backed up ${authFilePath} to ${authOldPath}`);
        }
        if (metaFileExists) {
            await fs.move(metaFilePath, metaOldPath, { overwrite: true });
            logger.info(`Backed up ${metaFilePath} to ${metaOldPath}`);
        }
    } else {
        // Ensure directory exists for new files
        await ensureAuthDir(authFilePath);
    }

    return {
        authFileExists,
        metaFileExists,
        willOverwriteOld,
        authOldPath,
        metaOldPath
    };
}

/** Deletes auth.json and auth-meta.json (full logout). */
async function logout(authFilePath) {
    const filePath = getAuthFilePath(authFilePath);
    const metaPath = getAuthMetaFilePath(filePath);
    
    await fs.remove(filePath);
    await fs.remove(metaPath);
}

/**
 * Prompts the user for input in the terminal.
 */
function promptUser(query) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });
    return new Promise(resolve => rl.question(query, ans => {
        rl.close();
        resolve(ans);
    }));
}

/**
 * The authenticated OneNote web app's path. It is not stable: the Microsoft 365
 * Copilot rebrand moved it from /notebooks to /copilotnotebooks, and
 * "/copilotnotebooks".includes("/notebooks") is false — so a substring check on
 * "/notebooks" alone timed out on a login that had in fact fully succeeded, with
 * the notebooks UI rendered and the account name on screen, and never saved the
 * auth state.
 *
 * Both spellings are listed, and the "notebooks" component is still required so
 * the unauthenticated marketing page (onenote.cloud.microsoft/en-us) can never
 * satisfy this. A bare hostname check would, and did: it caused premature
 * auth saving.
 *
 * This is matched against `url.pathname` and never against the serialised URL.
 * A substring match on the whole string accepts a notebooks path that is not a
 * path at all — a `?next=/copilotnotebooks` on the marketing page, or a
 * `?returnUrl=/notebooks` on a login screen — which is the same class of false
 * positive as the bare-hostname check, one level removed. It decides when live
 * auth state gets written to disk, so the check is worth making an actual
 * invariant about the path.
 *
 * Deliberately not constrained to a host allowlist: OneNote notebooks are
 * SharePoint-backed, and a tenant served from another host would be a real
 * successful login reported as a failure. Requiring the *path* is enough to
 * separate the app from the marketing page, and it cannot go stale the way a
 * host list would.
 */
const ONENOTE_APP_PATH = /\/copilotnotebooks\b|\/notebooks\b/;

/** UI markers that only render once the session is actually signed in. */
const ONENOTE_SIGNED_IN_MARKERS = [
    'text="My notebooks"',
    'text="Create new notebook"',
    'text="All Notebooks"',
    'text="Welcome, "'
];

/** The same idea for Outlook, which lands in a mailbox rather than a notebook list. */
const OUTLOOK_SIGNED_IN_MARKERS = [
    // Outlook: wait for email message list (table with emails)
    '[aria-label*="message list"], [role="grid"][aria-label*="mail"], .messageList',
    // Fallback: wait for folder navigation (Inbox, Sent, etc.)
    'text=/Inbox|Sent Mail|Drafts/i',
    // Fallback: wait for any email-like content
    'div[role="row"]'
];

/** How long success detection waits before giving up. */
const AUTH_SUCCESS_TIMEOUT = 60000;

/**
 * How long `check` waits for the navigation, and then for the page to settle.
 *
 * Split in two because they answer different questions. CHECK_NAV_TIMEOUT is
 * the old `page.goto` timeout and is about the network. CHECK_SETTLE_TIMEOUT
 * replaces the hardcoded 2 s sleep that used to stand in for "wait and see": a
 * session being redirected can take several seconds to get there, and cutting
 * that short is what made a dead session read as live.
 *
 * Sized against what `check` is for. It is a gate that decides whether to
 * export a notebook, so a few seconds of certainty is cheap next to a wrong
 * answer; and it only ever waits this long when the session is already broken.
 * A live session resolves on the first signal, usually in well under a second.
 *
 * Was 10 s, which was a coin flip rather than a budget: measured against a
 * working auth file, the signed-in UI ("My notebooks", "All Notebooks") rendered
 * 9.7 s after `domcontentloaded` — a 0.3 s margin on a warm run, and nothing at
 * all on a cold one. `check` therefore reported a perfectly good session as
 * expired on any day the network or the SPA's cache was slower than usual.
 * 30 s clears the measured case with room to spare, and it is only ever paid in
 * full by a session that is already broken.
 */
const CHECK_NAV_TIMEOUT = 15000;
const CHECK_SETTLE_TIMEOUT = 30000;

/** Outlook and OneNote are recognised differently, so the target picks the set. */
const isOutlookTarget = targetUrl => !!targetUrl && targetUrl.includes('outlook.cloud.microsoft');

/**
 * Every signal that the session reached the authenticated app, each paired with
 * a label naming it.
 *
 * One definition, used by both the production wait and the test probe, so a
 * selector changed here cannot pass the suite while production still waits on
 * the old one. They previously kept separate copies of this list; the OneNote
 * half was already shared, the Outlook half was a verbatim copy-paste.
 *
 * The labels earn their keep on the failure path. This repo has been bitten
 * twice by a success signal going stale without anyone noticing — the /notebooks
 * path moving to /copilotnotebooks, and a marker no longer rendered — and in
 * both cases the only symptom was a bare timeout. When every signal has missed,
 * the error can now say which ones were being watched, so the next one is
 * diagnosable from the report alone.
 *
 * Note these are *waits*, not a race with a timer: the first to resolve wins
 * and the rest are left to settle on their own.
 *
 * @returns {{label: string, wait: Promise<unknown>}[]}
 */
function authSuccessMarkerAttempts(page, targetUrl, timeoutMs) {
    const outlook = isOutlookTarget(targetUrl);
    const markers = outlook ? OUTLOOK_SIGNED_IN_MARKERS : ONENOTE_SIGNED_IN_MARKERS;
    const kind = outlook ? 'Outlook marker' : 'OneNote marker';
    return markers.map(marker => ({
        label: `${kind} "${marker}"`,
        wait: page.waitForSelector(marker, { state: 'visible', timeout: timeoutMs })
    }));
}

function authSuccessAttempts(page, targetUrl, timeoutMs) {
    if (isOutlookTarget(targetUrl)) return authSuccessMarkerAttempts(page, targetUrl, timeoutMs);

    return [
        {
            label: `OneNote app path ${ONENOTE_APP_PATH}`,
            // Primary: the URL must be the authenticated app, not the marketing page
            wait: page.waitForURL(url => ONENOTE_APP_PATH.test(url.pathname), { timeout: timeoutMs })
        },
        // Fallback UI elements that only appear when actually signed in
        ...authSuccessMarkerAttempts(page, targetUrl, timeoutMs)
    ];
}

/**
 * Waits for the authenticated app, bounded by `timeoutMs` instead of the full
 * production timeout. Returns whether it arrived, so it can be asserted on.
 * @returns {Promise<boolean>}
 */
async function waitForAuthSuccessProbe(page, targetUrl, timeoutMs) {
    const won = await Promise.any(
        authSuccessAttempts(page, targetUrl, timeoutMs).map(a => a.wait.then(() => true))
    ).catch(() => false);
    return won;
}

/**
 * Waits for successful authentication based on target URL.
 *
 * Rejects with a message naming what was being waited for and which signals were
 * being watched. `Promise.any` would otherwise reject with an AggregateError of
 * bare TimeoutErrors, which surfaces to the user as "All promises were
 * rejected" and names nothing at all.
 *
 * @param {import('playwright').Page} page - Playwright page object
 * @param {string} targetUrl - The target URL (ONENOTE_URL or OUTLOOK_URL)
 * @param {number} [timeoutMs] override for the wait; tests use a short one
 */
async function waitForAuthSuccess(page, targetUrl, timeoutMs = AUTH_SUCCESS_TIMEOUT) {
    const isOutlook = isOutlookTarget(targetUrl);

    logger.info(isOutlook
        ? 'Waiting for redirection to Outlook mail...'
        : 'Waiting for redirection to authenticated notebooks interface...');

    const attempts = authSuccessAttempts(page, targetUrl, timeoutMs);

    try {
        await Promise.any(attempts.map(a => a.wait));
    } catch (e) {
        // The caller catches this and adds where the browser actually ended up,
        // so this half only has to say what was being waited for and which
        // signals never arrived.
        throw new Error(
            `Timed out after ${timeoutMs / 1000}s waiting for ${isOutlook ? 'Outlook mail' : 'the authenticated OneNote app'}. `
            + `None of the ${attempts.length} success signals appeared: `
            + attempts.map(a => a.label).join('; ')
        );
    }

    logger.success(isOutlook
        ? 'Outlook mail interface detected.'
        : 'Authenticated notebooks interface detected.');
}

/**
 * Clicks the page-level "Cancel" button on the FIDO/security-key page and
 * waits for the browser to navigate away. Works because the addInitScript
 * override makes navigator.credentials.create() reject immediately, so the
 * native OS-level WebAuthn dialog never appears and the DOM is fully accessible.
 *
 * @param {import('playwright').Page} page
 * @param {object} logger
 * @returns {Promise<string|null>} 'Cancel' once dismissed, null if nothing was clickable
 */
async function dismissFidoPage(page, logger) {
    // Try multiple selectors for the page-level Cancel button.
    // On the FIDO page there are two Cancel-labelled things:
    //   1. The native OS WebAuthn dialog (blocked by addInitScript — never appears)
    //   2. The page-level gray "Cancel" button at the bottom of the form
    const cancelSelectors = [
        // Exact role button with text "Cancel" (the bottom-of-form button)
        'button:has-text("Cancel")',
        '[value="Cancel"]',
        'input[type="button"][value="Cancel"]',
    ];

    let clicked = false;
    for (const sel of cancelSelectors) {
        try {
            const btn = page.locator(sel).first();
            await btn.waitFor({ state: 'visible', timeout: 4000 });
            logger.info(`FIDO: clicking page-level Cancel via selector "${sel}"...`);
            await btn.click({ force: true });
            clicked = true;
            break;
        } catch (_) {
            // try next selector
        }
    }

    if (!clicked) {
        // Last resort: JS click on any visible Cancel button
        logger.warn('FIDO: DOM selectors failed, trying JS click fallback...');
        clicked = await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"]'));
            const cancel = btns.find(b => /^cancel$/i.test((b.textContent || b.value || '').trim()));
            if (cancel) {
                cancel.click();
                return true;
            }
            return false;
        });
    }

    if (!clicked) return null;

    // Wait for navigation away from the FIDO page (up to 8 s)
    try {
        await page.waitForURL(url => !url.toString().includes('/fido/'), { timeout: 8000 });
        logger.debug('FIDO: navigated away from FIDO page successfully.');
    } catch (_) {
        logger.warn('FIDO: still on FIDO URL after Cancel — continuing anyway.');
    }

    return 'Cancel';
}

/**
 * Answers the "Stay signed in?" prompt with "Yes" and ticks "Don't show this
 * again" so later logins skip the screen entirely.
 * @param {import('playwright').Page} page
 * @returns {Promise<string|null>} 'Yes' once clicked, null if the prompt is absent
 */
async function dismissStaySignedIn(page) {
    const staySignedIn = page.getByText(/Stay signed in?/i)
        .or(page.locator('#KmsiDescription'))
        .first();

    if (!(await staySignedIn.isVisible().catch(() => false))) return null;

    logger.info('Detected "Stay signed in?" prompt.');

    const dontShowAgain = page.locator('input[name="DontShowAgain"], #KmsiCheckboxField').first();
    if (await dontShowAgain.isVisible().catch(() => false)) {
        logger.debug('Checking "Don\'t show this again" checkbox...');
        await dontShowAgain.check().catch(() => { });
    }

    const yesButton = page.getByRole('button', { name: /^Yes$/i })
        .or(page.locator('button[data-testid="primaryButton"]'))
        .or(page.locator('#idSIButton9'))
        .first();

    logger.info('Clicking "Yes" to stay signed in...');
    await yesButton.click({ timeout: 10000 });
    return 'Yes';
}

/**
 * Screens Microsoft injects in the middle of an otherwise successful login.
 * They are full-page forms that hijack the navigation, so nothing after them
 * (MFA checks, "Stay signed in?", the redirect to OneNote/Outlook) can be
 * reached until they are dismissed. All of them are full-page forms that hijack
 * the navigation, and all of them arrive *late* — typically 20-60 s after the
 * password is accepted, once per account:
 *
 *   1. account.live.com/interrupt/credentialaction or /proofs/remind
 *      "Is your security info still accurate?" -> Looks good! (proof freshness)
 *   2. account.live.com/tou/accrue
 *      "We're updating our terms" -> Next (Services Agreement update)
 *   3. account.live.com/interrupt/passkey -> login.microsoft.com/consumers/fido/create
 *      the passkey prompt -> Cancel
 *   4. login.live.com/... "Stay signed in?" -> Yes
 *
 * Each entry carries the only action labels that may be pressed on it, so a
 * broad label like "Yes" can never be pressed on a screen that does not offer it.
 * Screens are matched on the URL *and/or* the heading, because Microsoft moves
 * them between paths (and serves the same screen from several) over time.
 */
const BLOCKING_SCREENS = [
    {
        name: 'FIDO / passkey prompt',
        match: state => /consumers\/fido\//i.test(state.url)
            || /passkey/i.test(state.url)
            || /passkey|security key/i.test(state.heading),
        // Dedicated WebAuthn dismisser rather than a label match: the page-level
        // "Cancel" is the only safe action and it also waits out the navigation.
        handle: page => dismissFidoPage(page, logger)
    },
    {
        name: 'Terms of Use / Services Agreement update',
        match: state => /account\.live\.com\/tou\//i.test(state.url),
        actions: /^(next|accept|i accept|i agree|agree|continue|finish|done)$/i
    },
    {
        name: 'Microsoft consent prompt',
        match: state => /consent\./i.test(state.url),
        actions: /^(accept|i accept|i agree|agree|continue|next)$/i
    },
    {
        name: 'Security info freshness check',
        match: state => (/account\.live\.com\/(pf|proofs|interrupt)/i.test(state.url) && !/passkey/i.test(state.url))
            || /is your security info still accurate/i.test(state.heading)
            || /help protect your account/i.test(state.heading),
        // "Update now" and "I don't have any of these" are deliberately absent:
        // either would rewrite or wipe the account's recovery methods.
        actions: /^(looks good!?|skip for now|skip|continue|next)$/i
    },
    {
        name: '"Stay signed in?" prompt',
        match: state => /stay signed in/i.test(state.heading) || /kmsi/i.test(state.url),
        handle: page => dismissStaySignedIn(page)
    },
];

/**
 * Buttons that may carry an action label, most specific first. account.live.com
 * renders its pages with Fluent UI, where the action is always the primary
 * button; the plain selectors are the fallback for the older server-rendered
 * account pages.
 */
const BLOCKING_SCREEN_BUTTONS = [
    'button[data-testid="primaryButton"]',
    'input[data-testid="primaryButton"]',
    'button',
    'input[type="submit"]',
    'input[type="button"]',
    '[role="button"]',
    'a',
];

/** Do not re-click the same unchanged screen more often than this. */
const BLOCKING_SCREEN_RETRY_MS = 10000;

/**
 * Dump basenames for the two passes that clear blocking screens.
 *
 * They must differ. The first pass runs once, just after the password is
 * accepted; the watcher then polls for the rest of the login and is where the
 * screens that arrive *late* turn up — the terms update queued behind a "Stay
 * signed in?" prompt, typically 20-60 s in. Both passes call
 * clearBlockingScreens(), which numbers its dumps from 1 within a single call,
 * and the watcher starts a new call every second. On a shared basename the
 * late screen would overwrite the earlier pass's dump of a different screen,
 * and the one that was actually in the way when the login stalled would be the
 * one lost.
 */
const BLOCKING_SCREEN_DUMP = 'debug_blocking_screen';
const LATE_BLOCKING_SCREEN_DUMP = 'debug_late_blocking_screen';

/** Returns the blocking-screen descriptor matching { url, heading }, or null. */
function matchBlockingScreen(state) {
    if (!state) return null;
    for (const screen of BLOCKING_SCREENS) {
        try {
            if (screen.match(state)) return screen;
        } catch (_) {
            // A malformed heading/url must never abort the whole login.
        }
    }
    return null;
}

/**
 * Reads the current URL + heading of the page.
 * Returns null while the document is being swapped (mid-navigation), so callers
 * must retry rather than treat null as "no blocking screen".
 * @param {import('playwright').Page} page
 */
async function readScreenState(page) {
    try {
        return await page.evaluate(() => {
            const heading = document.querySelector('h1, [role="heading"], [data-testid="title"]');
            return {
                url: location.href,
                heading: (heading ? heading.textContent : '').replace(/\s+/g, ' ').trim()
            };
        });
    } catch (_) {
        // Execution context destroyed while navigating — caller should retry.
        return null;
    }
}

/** Shortens a URL for logging: keeps host + path, drops the query string. */
function shortUrl(url) {
    try {
        const parsed = new URL(url);
        return `${parsed.host}${parsed.pathname}`;
    } catch (_) {
        return url;
    }
}

/**
 * Stable identity of a screen: same URL + same heading means the same step.
 * Used to tell a genuine step change from a re-render of the same page.
 */
function screenSignature(state) {
    return state ? `${state.url}::${state.heading}` : null;
}

/** Reads the visible label of a button-like element. */
async function readActionLabel(handle) {
    const text = await handle.textContent().catch(() => '') || '';
    const value = await handle.getAttribute('value').catch(() => '') || '';
    return `${text} ${value}`.replace(/\s+/g, ' ').trim();
}

/**
 * Clicks the single accept/continue action on a blocking screen.
 * Only labels accepted by *this* screen's `actions` regex are eligible, so a
 * stray "Skip" in a footer or a "Yes" meant for a different screen is never
 * pressed by mistake.
 * @param {import('playwright').Page} page
 * @param {RegExp} actions
 * @returns {Promise<string|null>} the label that was clicked, or null
 */
async function clickBlockingScreenAction(page, actions) {
    for (const selector of BLOCKING_SCREEN_BUTTONS) {
        const buttons = page.locator(selector);
        const count = await buttons.count().catch(() => 0);

        for (let i = 0; i < Math.min(count, 30); i++) {
            const button = buttons.nth(i);
            if (!(await button.isVisible().catch(() => false))) continue;

            const label = await readActionLabel(button);
            if (!label || !actions.test(label)) continue;

            await button.click({ timeout: 10000 });
            return label;
        }
    }
    return null;
}

/**
 * Waits until the blocking screen actually moves on. The Terms of Use flow is a
 * single-page app, so the URL stays put across steps and only the heading (or the
 * presence of the button) changes — a navigation wait alone would always time out.
 * @returns {Promise<string|null>} the new signature, or null on timeout
 */
async function waitForBlockingScreenChange(page, previousSignature, timeout) {
    const deadline = Date.now() + timeout;

    while (Date.now() < deadline) {
        await page.waitForTimeout(400).catch(() => {});

        const state = await readScreenState(page);
        if (!state) continue;                       // mid-navigation, keep polling
        if (screenSignature(state) !== previousSignature) return screenSignature(state);
    }
    return null;
}

/**
 * Clears every blocking screen currently in the way of the login.
 *
 * A single acceptance usually leads to one or two more (e.g. the Services
 * Agreement summary, then a "Finish" confirmation), so this loops until the
 * page is no longer a blocking screen. `progress` is shared with the caller so
 * that repeated invocations from the watcher do not hammer an unchanged screen.
 *
 * @param {import('playwright').Page} page
 * @param {object} [options]
 * @param {{ signatures: Set<string>, lastClickAt: number }} [options.progress]
 * @param {boolean} [options.dodump]
 * @param {boolean} [options.screenshot]  also screenshot each dump
 * @param {string} [options.dumpFile]      basename for the dumps this call writes
 * @param {() => boolean} [options.shouldStop]
 * @returns {Promise<{ handled: number, reason: string }>}
 */
async function clearBlockingScreens(page, options = {}) {
    const {
        progress = { signatures: new Set(), lastClickAt: 0 },
        maxScreens = 5,
        stateTimeout = 10000,
        changeTimeout = 20000,
        dodump = false,
        screenshot = false,
        dumpFile = BLOCKING_SCREEN_DUMP,
        shouldStop = null
    } = options;

    let handled = 0;

    for (let i = 0; i < maxScreens; i++) {
        if (shouldStop && shouldStop()) return { handled, reason: 'stopped' };

        // The screen may still be loading; give it a bounded number of chances.
        let state = null;
        const stateDeadline = Date.now() + stateTimeout;
        do {
            state = await readScreenState(page);
            if (!state) await page.waitForTimeout(500).catch(() => {});
        } while (!state && Date.now() < stateDeadline && !(shouldStop && shouldStop()));

        if (!state) return { handled, reason: 'unreadable' };
        const signature = screenSignature(state);

        const screen = matchBlockingScreen(state);
        if (!screen) return { handled, reason: 'no_blocking_screen' };

        // Never click the exact same screen twice in quick succession: the click
        // either worked (signature changes) or the page is stuck, and a tight
        // retry loop would only spam requests at Microsoft.
        if (progress.signatures.has(signature)) {
            if (Date.now() - progress.lastClickAt < BLOCKING_SCREEN_RETRY_MS) {
                logger.debug(`Blocking screen unchanged since last attempt — not re-clicking.`);
                return { handled, reason: 'unchanged' };
            }
        } else {
            progress.signatures.add(signature);
        }

        logger.info(`Blocking screen detected: ${screen.name} (${shortUrl(state.url)}). Accepting it...`);

        if (dodump) {
            const fileName = `${dumpFile}_${i + 1}.html`;
            const displayPath = await dumpPage(page, fileName, { screenshot });
            logger.debug(`[dodump] Blocking screen state dumped to ${displayPath}/${fileName}`);
        }

        let label = null;
        try {
            label = screen.handle
                ? await screen.handle(page)
                : await clickBlockingScreenAction(page, screen.actions);
        } catch (e) {
            logger.debug(`Blocking screen click failed: ${e.message}`);
        }

        if (!label) {
            logger.warn(`No acceptable action button found on "${screen.name}". Stopping.`);
            return { handled, reason: 'no_action' };
        }

        handled++;
        progress.lastClickAt = Date.now();
        logger.debug(`Clicked "${label}" on ${screen.name}.`);

        await waitForBlockingScreenChange(page, signature, changeTimeout);
    }

    return { handled, reason: 'max_screens' };
}

/**
 * Polls for blocking screens until told to stop, clearing any that appear.
 *
 * This is the last line of defence for the interstitial screens that arrive
 * *late*. The pass in step 2.5a runs once, right after the password is
 * accepted; Microsoft often serves these screens much later — 20-60 s in, once
 * per account, behind a "Stay signed in?" prompt or behind a terms update that
 * was itself queued behind something else. The watcher is what catches those,
 * and until now it caught them blind: it ran with no dumping at all, so a login
 * killed by a late consent page produced nothing in the dump directory and the
 * only evidence was the bare "still on <url>" timeout.
 *
 * It inherits `dodump`/`screenshot` from the caller so a late screen is captured
 * the same way an early one is, but under LATE_BLOCKING_SCREEN_DUMP: see the
 * note on that constant for why the two passes cannot share a basename.
 *
 * `progress` is shared with the earlier pass, which is what keeps the watcher
 * from re-capturing a screen that pass already handled, and the retry cooldown
 * inside clearBlockingScreens is what keeps it from writing a fresh dump every
 * second while a screen sits there unchanged.
 *
 * @param {import('playwright').Page} page
 * @param {object} [options]
 * @param {{ signatures: Set<string>, lastClickAt: number }} [options.progress]
 * @param {number} [options.pollInterval]  ms between polls
 * @param {number} [options.maxScreens]    screens to clear per poll
 * @param {number} [options.stateTimeout]  forwarded to clearBlockingScreens
 * @param {number} [options.changeTimeout] forwarded to clearBlockingScreens
 * @param {boolean} [options.dodump]
 * @param {boolean} [options.screenshot]
 * @param {() => boolean} options.shouldStop  required; the watcher never ends on its own
 * @returns {Promise<void>} resolves once shouldStop() has been true for one poll
 */
async function watchBlockingScreens(page, options = {}) {
    const {
        progress = { signatures: new Set(), lastClickAt: 0 },
        pollInterval = 1000,
        maxScreens = 2,
        stateTimeout = 10000,
        changeTimeout = 20000,
        dodump = false,
        screenshot = false,
        shouldStop = () => false
    } = options;

    while (!shouldStop()) {
        try {
            await clearBlockingScreens(page, {
                progress,
                maxScreens,
                stateTimeout,
                changeTimeout,
                dodump,
                screenshot,
                dumpFile: LATE_BLOCKING_SCREEN_DUMP,
                shouldStop
            });
        } catch (e) {
            // The watcher runs for the whole wait, so anything thrown here would
            // otherwise be logged once and the loop would carry on regardless.
            logger.debug(`Blocking screen watcher error: ${e.message}`);
        }
        await page.waitForTimeout(pollInterval).catch(() => {});
    }
}

/* ------------------------------------------------------------------------- *
 * "How do you want to sign in?" — the screens between email and password
 *
 * For a passwordless-enabled account, Microsoft serves this after the email
 * step (captured in logs/dumps, PageID i5030):
 *
 *     <h1 data-testid="title">Get a code to sign in</h1>
 *     <button type="submit" data-testid="primaryButton">Send code</button>
 *     <span role="button" class="fui-Link" tabindex="0">Use your password</span>
 *
 * Note what is *not* there: an "Other ways to sign in" link. The password is
 * only reachable via the "Use your password" link in the footer of that very
 * screen, and it is a <span role="button">, not an <a>, so it only responds to
 * real pointer events. The previous implementation assumed "Other ways to sign
 * in" always came first: it waited 15 s for a link that does not exist, threw
 * STUCK, swallowed it, and then let the password selector time out 30 s later.
 * A ~45 s stall ending in a misleading "incorrect credentials" message.
 *
 * So instead of racing text selectors and betting on which one wins, the state
 * is read from the DOM — what is actually on offer — and "Use your password" is
 * pressed wherever it appears.
 * ------------------------------------------------------------------------- */

/**
 * Label patterns for the sign-in-choice screens, as *source strings*: they are
 * handed into page.evaluate(), where RegExp objects do not survive
 * serialization and must be rebuilt on the far side.
 */
const SIGN_IN_LABELS = {
    usePassword: 'use your password|use a password instead|use password instead|sign in with a password',
    otherWays: 'other ways to sign in|sign in another way|look for another way|try another way',
    // The list of methods, where "Password" is one entry among several.
    passwordEntry: '^password$|use your password|use a password',
    methodList: 'select a (?:sign-in |verification )?method|choose (?:a |another )?way to sign in|how do you want to sign in',
    sendCode: 'send code|get a code to sign in|text me a code|email me a code',
    approveApp: 'approve a request on my microsoft authenticator app|approve sign in request',
    otcPrompt: 'enter (?:the )?code|type (?:the )?code|verification code',
};

/** Compiles a SIGN_IN_LABELS entry into an anchored, case-insensitive RegExp. */
function signInLabel(key) {
    return new RegExp(SIGN_IN_LABELS[key], 'i');
}

/** Selectors for anything on the page that can be clicked, on either UI generation. */
const CLICKABLE_SELECTOR = 'a[href], button, input[type="submit"], input[type="button"], [role="button"], [role="link"]';

/** The password box, on both the legacy and the Fluent sign-in pages. */
const PASSWORD_FIELD_SELECTOR = 'input[name="passwd"], input[type="password"]';

/**
 * Reads what the current sign-in screen actually offers, in a single round trip.
 *
 * Everything is answered from one evaluate() so the reading is a consistent
 * snapshot: seven separate locators would each sample the page at a slightly
 * different moment, which is how a screen mid-navigation gets misclassified.
 *
 * @param {import('playwright').Page} page
 * @returns {Promise<object|null>} null while the document is being swapped
 */
async function readSignInState(page) {
    try {
        return await page.evaluate(labels => {
            const visible = el => {
                if (!el) return false;
                // aria-hidden means the element is not exposed to the user, so it
                // is not something the user can interact with. login.microsoftonline.com
                // parks its leftover fields in the DOM exactly this way:
                // <input name="loginfmt" class="moveOffScreen" aria-hidden="true">.
                // Those are off-screen rather than display:none, so they still have
                // a non-zero box and pass a pure geometry test — which is how a
                // work account's password page was mistaken for the email step.
                // Compared case-insensitively because ARIA token values are
                // defined as ASCII case-insensitive, so "True" means the same
                // thing and must not read as visible.
                if ((el.getAttribute('aria-hidden') || '').trim().toLowerCase() === 'true') return false;
                const rect = el.getBoundingClientRect();
                if (rect.width <= 0 || rect.height <= 0) return false;
                const style = window.getComputedStyle(el);
                return style.visibility !== 'hidden' && style.display !== 'none';
            };

            const labelOf = el => `${el.value || ''} ${el.textContent || ''}`.replace(/\s+/g, ' ').trim();

            // True when a *control* carrying this label is on screen. Deliberately
            // not a body-text search: prose mentioning "another way" must never
            // be mistaken for the button that acts on it.
            const offersAction = key => {
                const rx = new RegExp(labels[key], 'i');
                return Array.from(document.querySelectorAll(
                    'a[href], button, input[type="submit"], input[type="button"], [role="button"], [role="link"]'
                )).some(el => visible(el) && rx.test(labelOf(el)));
            };

            const headingEl = document.querySelector('h1, [role="heading"], [data-testid="title"]');
            const body = document.body ? (document.body.innerText || '').replace(/\s+/g, ' ') : '';

            return {
                url: location.href,
                heading: (headingEl ? headingEl.textContent : '').replace(/\s+/g, ' ').trim(),
                // Still on the email form: the step after it has not rendered yet.
                emailField: visible(document.querySelector('input[name="loginfmt"]')),
                passwordField: visible(document.querySelector('input[name="passwd"], input[type="password"]')),
                otcField: visible(document.querySelector('input[name="otc"], input[type="tel"]')),
                usePassword: offersAction('usePassword'),
                otherWays: offersAction('otherWays'),
                sendCode: offersAction('sendCode') || new RegExp(labels.sendCode, 'i').test(body),
                // Prose-only screens: read from the page text, nothing to click.
                methodList: new RegExp(labels.methodList, 'i').test(body),
                approveApp: new RegExp(labels.approveApp, 'i').test(body),
                otcPrompt: new RegExp(labels.otcPrompt, 'i').test(body)
            };
        }, SIGN_IN_LABELS);
    } catch (_) {
        // Execution context destroyed mid-navigation — caller should retry.
        return null;
    }
}

/** True when the screen is one this handler knows how to act on. */
function isActionableSignInState(state) {
    if (!state || state.emailField) return false;
    return !!(state.passwordField || state.usePassword || state.otherWays
        || state.methodList || state.sendCode || state.approveApp || state.otcPrompt);
}

/** Identity of a sign-in screen, to tell a real step change from a re-render. */
function signInStateSignature(state) {
    if (!state) return null;
    return [state.url, state.heading, state.passwordField, state.usePassword,
        state.otherWays, state.methodList, state.sendCode, state.approveApp].join('::');
}

/**
 * Polls until `accept` is satisfied or the timeout runs out.
 * @returns {Promise<object|null>} the accepted state, else the last readable one
 */
async function waitForSignInState(page, timeout, accept = isActionableSignInState) {
    const deadline = Date.now() + timeout;
    let state = null;
    do {
        const read = await readSignInState(page);
        if (read) {
            state = read;
            if (accept(read)) return read;
        }
        await page.waitForTimeout(400).catch(() => {});
    } while (Date.now() < deadline);
    // The last *readable* state, not whatever a mid-navigation read happened to
    // return: null would discard the only evidence of why the step is stuck.
    return state;
}

/**
 * Clicks the control carrying a given label.
 *
 * The Fluent pages render links as <span role="button">, which only react to a
 * real pointer event: a .click() on an ancestor wrapper fires nothing. So the
 * lookup escalates from the most semantic to the most forceful, and the last
 * step targets the *innermost* match rather than the first in document order.
 *
 * @param {import('playwright').Page} page
 * @param {RegExp} rx
 * @param {{ timeout?: number }} [options]
 * @returns {Promise<string|null>} the label clicked, or null if nothing matched
 */
async function clickByLabel(page, rx, { timeout = 8000 } = {}) {
    // Precise strategies first: they match on the accessible name, so they hit
    // the control the user sees rather than a wrapper.
    const precise = [
        page.getByRole('button', { name: rx }),
        page.getByRole('link', { name: rx })
    ];

    // Raced, not sequential: a control is a button on one screen generation and a
    // link on the other, so trying them in turn means paying the full timeout on
    // whichever role does not apply.
    const hit = await Promise.any(
        precise.map((locator, i) => locator.first().waitFor({ state: 'visible', timeout }).then(() => i))
    ).catch(() => -1);

    // getByText is the fuzzier fallback, and the JS click the forceful one.
    if (hit < 0) {
        try {
            await page.getByText(rx).first().waitFor({ state: 'visible', timeout: Math.min(timeout, 2000) });
            precise.push(page.getByText(rx));
            hit = precise.length - 1;
        } catch (_) { /* fall through to the JS click */ }
    }

    if (hit >= 0) {
        const first = precise[hit].first();
        try {
            // Read the label *before* clicking. These controls navigate on click,
            // and textContent() against a locator whose element the navigation just
            // removed blocks for the full 30 s default timeout before rejecting —
            // which is exactly the stall this function is meant to avoid.
            const label = ((await first.textContent({ timeout: 2000 }).catch(() => '')) || '')
                .replace(/\s+/g, ' ').trim();
            await first.click({ timeout: Math.min(timeout, 5000) });
            return label || '(clicked)';
        } catch (_) {
            // Found but not clickable — let the JS click have a go.
        }
    }

    return await page.evaluate(({ selector, source }) => {
        const rx = new RegExp(source, 'i');
        const labelOf = el => `${el.value || ''} ${el.textContent || ''}`.replace(/\s+/g, ' ').trim();
        const hits = Array.from(document.querySelectorAll(selector)).filter(el => rx.test(labelOf(el)));
        // Innermost hit: an ancestor's textContent includes the descendant's, so
        // the first match in document order is usually a wrapper whose click
        // never reaches the handler the user can see.
        const target = hits.find(el => !hits.some(other => other !== el && el.contains(other)));
        if (!target) return null;
        target.click();
        return labelOf(target);
    }, { selector: CLICKABLE_SELECTOR, source: rx.source }).catch(() => null);
}

/**
 * Presses the sign-in submit control.
 *
 * The legacy pages use <input type="submit" value="Sign in">, where the label
 * lives in `value` and `filter({ hasText })` can never match it; the Fluent pages
 * use <button type="submit" data-testid="primaryButton">Sign in</button>. Matching
 * on the accessible role name covers both.
 *
 * @param {import('playwright').Page} page
 * @returns {Promise<boolean>} true if a submit control was clicked
 */
async function submitSignInForm(page) {
    const strategies = [
        page.getByRole('button', { name: /^(sign in|next|finish|continue)$/i }),
        page.locator('input[type="submit"]'),
        page.locator('button[type="submit"]')
    ];

    for (const locator of strategies) {
        const button = locator.first();
        if (!(await button.isVisible().catch(() => false))) continue;
        // click() waits for the element to become enabled, which covers the
        // short window after fill() while the page validates the password.
        await button.click({ timeout: 10000 });
        return true;
    }
    return false;
}

/**
 * Walks the "how do you want to sign in?" screens and lands on the password box.
 *
 * Prefers "Use your password" wherever it is offered, because on the
 * passwordless screen that link is the only route to the password — there is no
 * "Other ways to sign in" step to go through first. Only falls back to that
 * step when the screen really does present it, then picks "Password" out of the
 * resulting method list.
 *
 * @param {import('playwright').Page} page
 * @param {object} [options]
 * @param {number} [options.maxSteps]
 * @param {number} [options.stateTimeout]   how long to wait for the first screen
 * @param {number} [options.transitionTimeout] how long to wait after each click
 * @param {boolean} [options.dodump]
 * @param {boolean} [options.screenshot]   also screenshot the dump
 * @param {string} [options.dumpFile]        basename written under the dump dir
 * @returns {Promise<{ reached: boolean, reason: string, steps: number, state: object|null }>}
 */
async function reachPasswordScreen(page, options = {}) {
    const {
        maxSteps = 4,
        stateTimeout = 15000,
        transitionTimeout = 10000,
        dodump = false,
        screenshot = false,
        dumpFile = 'debug_intermediate_screen'
    } = options;

    // The password box has to be accepted by this *initial* wait, not only by
    // the loop below. isActionableSignInState is false whenever emailField is
    // set, so on a layout that keeps a real username field on screen next to the
    // password box — "Use a different account" and friends — this wait could
    // never be satisfied and always ran to its full stateTimeout, no matter how
    // long it was given. The password box then got noticed by the loop
    // afterwards, so the login still worked: it just took 15 s to start.
    //
    // Same shape as the post-click wait at the bottom of the loop. On the
    // ordinary email step both terms are false, so the behaviour there is
    // unchanged: still waits out the timeout and reports "unreadable".
    let state = await waitForSignInState(page, stateTimeout,
        s => s.passwordField || isActionableSignInState(s));
    let dumped = false;

    for (let steps = 0; steps < maxSteps; steps++) {
        // The password box is checked FIRST, ahead of the "still on the email
        // step" gate below. A work/school account signs in on
        // login.microsoftonline.com, whose password page keeps the username
        // field in the DOM as <input name="loginfmt" class="moveOffScreen"> —
        // rendered off-screen, and therefore "visible" by any geometry test.
        // Gating on emailField first made every such login report "unreadable"
        // and burn the whole stateTimeout while the password box sat right
        // there, ready to be filled.
        if (state && state.passwordField) {
            logger.debug(`Password field reached after ${steps} step(s).`);
            return { reached: true, reason: 'password_field', steps, state };
        }

        if (!state || !isActionableSignInState(state)) {
            return { reached: false, reason: 'unreadable', steps, state };
        }

        if (dodump && !dumped) {
            dumped = true;
            const displayPath = await dumpPage(page, `${dumpFile}.html`, { screenshot });
            logger.debug(`[dodump] Intermediate screen state dumped to ${displayPath}/${dumpFile}.html`);
        }

        const before = signInStateSignature(state);
        let clicked = null;

        if (state.usePassword) {
            logger.info('Password is offered on this screen — clicking "Use your password"...');
            clicked = await clickByLabel(page, signInLabel('usePassword'));
        } else if (state.otherWays) {
            logger.info('Opening "Other ways to sign in"...');
            clicked = await clickByLabel(page, signInLabel('otherWays'));
        } else if (state.methodList) {
            logger.info('Choosing "Password" from the sign-in method list...');
            clicked = await clickByLabel(page, signInLabel('passwordEntry'));
        }

        if (!clicked) {
            // No route to a password from here. Say which kind of screen it is,
            // so a mandatory code/approval is not reported as a bad password.
            const reason = state.approveApp ? 'approver_prompt'
                : state.otcPrompt || state.sendCode || state.otcField ? 'code_prompt'
                    : 'no_password_route';
            logger.warn(`No way to reach the password screen from "${state.heading || shortUrl(state.url)}" (${reason}).`);
            return { reached: false, reason, steps, state };
        }

        logger.debug(`Clicked "${clicked}". Waiting for the next screen...`);

        // Wait for a *different* screen, not just the password box: the method
        // list is a legitimate hop in the middle of this walk, and waiting only
        // for the password field would burn the whole timeout on it. The
        // signature guard keeps the not-yet-navigated pre-click state from
        // satisfying the wait immediately.
        state = await waitForSignInState(page, transitionTimeout,
            s => s.passwordField || (isActionableSignInState(s) && signInStateSignature(s) !== before));
        if (state && signInStateSignature(state) === before) {
            logger.warn(`Screen did not change after clicking "${clicked}".`);
            return { reached: false, reason: 'unchanged', steps, state };
        }
    }

    return { reached: false, reason: 'max_steps', steps: maxSteps, state };
}

/**
 * Writes the current page to a debug dump, with credentials removed.
 *
 * `--dodump` calls page.content(), and that serialises the *live value* of every
 * form control. On the page where the password is typed that means the user's
 * actual Microsoft password lands on disk in cleartext:
 *
 *     <input type="password" name="passwd" value="the-real-password">
 *
 * The dumps are gitignored and never packed, but they land in the working
 * tree, where backups, sync clients and file-sharing will pick them up, and
 * they get pasted into issues and chat. Every dump goes through here so the
 * redaction cannot be forgotten at the next call site.
 *
 * The scrubbing happens on a *clone* of the document rather than by rewriting
 * the HTML string: PPFT and the other flow tokens are the values the login
 * form is about to POST, so blanking them in the live DOM would break the very
 * login being debugged. A clone cannot affect the page.
 *
 * With `screenshot`, a PNG of the rendered page is written beside it under the
 * same basename, because the HTML serialisation says which screen this was but
 * nothing about how it *looked*: which element covered the button, whether a
 * banner or an overlay was in the way, how the screen was laid out on the
 * viewport. Several of the failures this tool reports ("the password box never
 * appeared", "stuck on <url>") are resolved by a glance at the picture.
 *
 * The PNG is not scrubbed, and cannot be: it is a bitmap. That costs nothing in
 * credential terms — a password field renders as dots, and nothing else on a
 * Microsoft login screen is secret in a way the redacted HTML does not already
 * disclose — but a screenshot *does* show the number-match MFA code, which the
 * HTML dump and the terminal log both show too. So the screenshots are as
 * sensitive as the dumps and are written to the same gitignored directory.
 *
 * @param {import('playwright').Page} page
 * @param {string} fileName  basename, e.g. debug_after_password.html
 * @param {object} [options]
 * @param {boolean} [options.screenshot] also write <basename>.png beside the HTML
 * @returns {Promise<string>} the display path the dump was written to
 */
async function dumpPage(page, fileName, options = {}) {
    const { screenshot = false } = options;
    const dumpDir = await logger.getDumpDir();
    const displayPath = logger.getDumpDisplayPath();
    await fs.writeFile(path.join(dumpDir, fileName), await redactedPageContent(page));
    if (screenshot) await screenshotPage(page, dumpDir, displayPath, fileName);
    return displayPath;
}

/**
 * The screenshot that accompanies a dump: debug_after_email.html ->
 * debug_after_email.png. Derived from the HTML basename so the pair is found
 * together, and an unexpected extension cannot produce "x.html.png".
 *
 * @param {string} fileName
 * @returns {string}
 */
function screenshotNameFor(fileName) {
    return fileName.replace(/\.html?$/i, '') + '.png';
}

/**
 * How long a screenshot may take before it is given up on. Playwright's own
 * default is 30 s, which is longer than most of the waits it would sit inside:
 * a dump taken while the login is already behind schedule must not become the
 * slowest thing in it.
 */
const SCREENSHOT_TIMEOUT_MS = 15000;

/**
 * Screenshots the page into the dump directory, never throwing.
 *
 * A capture is the one part of a dump that depends on the page being alive and
 * still rendering, and it is by far the least important part: the HTML dump is
 * already on disk by the time this runs. So every failure — the page navigated
 * away mid-capture, the tab was closed, a headless browser that cannot rasterise
 * — is a warning, never a failed login.
 *
 * @param {import('playwright').Page} page
 * @param {string} dumpDir
 * @param {string} displayPath
 * @param {string} fileName  basename of the HTML dump this accompanies
 */
async function screenshotPage(page, dumpDir, displayPath, fileName) {
    const pngName = screenshotNameFor(fileName);
    try {
        // fullPage: the point is to see the screen as it is presented, and a
        // Microsoft login page that scrolls below the fold — the method list, the
        // footer links — is exactly the case where a viewport-only capture
        // answers "nothing is here" when something is.
        await page.screenshot({
            path: path.join(dumpDir, pngName),
            fullPage: true,
            timeout: SCREENSHOT_TIMEOUT_MS
        });
        logger.debug(`[screenshot] Saved ${displayPath}/${pngName}`);
    } catch (e) {
        logger.warn(`[screenshot] Could not capture ${pngName}: ${e.message} (the HTML dump was still written)`);
    }
}

/** Placeholder written over any redacted value, so a scrubbed dump is obvious. */
const REDACTED = '[redacted]';

/**
 * Serialises the page with credential-bearing control values blanked.
 *
 * Redacted: every `input[type=password]` whatever it is named, plus any control
 * whose name or id reads as a credential or a bearer token — PPFT and the other
 * pre-auth flow tokens, id/access/refresh tokens, secrets, and OTP / one-time
 * code fields.
 *
 * Deliberately *not* redacted: the account identifier (`loginfmt`, `login`).
 * It is already written in cleartext by the "Attempting automated login for
 * <email>" line and is on the command line, so hiding it in the HTML would
 * protect nothing while removing the one field worth having when a login fails
 * on the wrong account.
 *
 * @param {import('playwright').Page} page
 * @returns {Promise<string>}
 */
async function redactedPageContent(page) {
    return await page.evaluate(REDACTED => {
        // Any control whose name or id holds a credential or a bearer token
        // rather than UI state. Substring matching on purpose: these names are
        // compound in the wild (otc, otcFallback, verificationCode, srfSFT) and
        // an anchored pattern misses all but the exact spelling. Over-redacting
        // one extra field costs a debug dump nothing; under-redacting leaks a
        // credential. PPFT is Microsoft's pre-auth flow token and matches no
        // generic word, so it is named outright.
        const SENSITIVE = /ppft|token|canary|secret|passw|credential|otp|otc|code$|pin$/i;

        const isSensitive = el => {
            if (String(el.type || '').toLowerCase() === 'password') return true;
            return SENSITIVE.test(`${el.name || ''}`.trim()) || SENSITIVE.test(`${el.id || ''}`.trim());
        };

        const clone = document.documentElement.cloneNode(true);
        for (const el of clone.querySelectorAll('input, textarea')) {
            // Only a value that is actually there needs hiding.
            if (!el.value || !isSensitive(el)) continue;

            if (el.tagName === 'TEXTAREA') {
                el.textContent = REDACTED;
            } else {
                // setAttribute, never `.value =`. Assigning the IDL property on a
                // *visible* input puts the element into "dirty value mode": the IDL
                // value changes but the content attribute is left alone, and
                // outerHTML serialises the content attribute — so the secret comes
                // out unchanged. Only type="hidden" inputs are saved by that
                // accident, and a password field is a visible input, which is
                // exactly the case that would have leaked.
                el.setAttribute('value', REDACTED);
            }
        }
        return `<!DOCTYPE html>\n${clone.outerHTML}`;
    }, REDACTED).catch(e => `<!-- Error redacting or reading page: ${e.message} -->`);
}

/**
 * Confirms the auth state file really landed on disk as parseable JSON.
 *
 * A login is only "successful" if there is something to be successful *with*.
 * `context.storageState({ path })` and `fs.writeJson()` both report failure by
 * throwing, but neither of them guarantees success: a path that resolves
 * somewhere unexpected, a filesystem that swallows the write, or a file left
 * truncated by a crash mid-write all end the login with a cheerful
 * "Authentication successful!" and no usable state behind it. This re-reads
 * what was written, so "successful" is a statement about the disk rather than
 * about two calls that returned.
 *
 * The shape check is the part that matters. Playwright's storageState is
 * `{ cookies: [...], origins: [...] }`, and `getAuthenticatedContext` /
 * `checkAuth` both feed the path straight to `browser.newContext({ storageState
 * })`. A file that exists and parses but has no `cookies` array would be
 * rejected there — a login that reports success and then produces a context
 * that cannot authenticate anything, which is precisely the silent failure this
 * whole change exists to eliminate.
 *
 * @param {string} authFilePath  the path login() wrote to
 * @returns {Promise<{ ok: boolean, reason: string, detail: string }>}
 *   `ok` is true only when the file exists, parses as JSON and has a `cookies`
 *   array. `reason` is a stable short code; `detail` is for the log.
 */
async function verifyAuthStateFile(authFilePath) {
    if (!(await fs.pathExists(authFilePath))) {
        return {
            ok: false,
            reason: 'missing',
            detail: `${authFilePath} does not exist`
        };
    }

    let state;
    try {
        state = await fs.readJson(authFilePath);
    } catch (e) {
        // A truncated or half-written file lands here, and is the most likely
        // real-world shape of this failure.
        return {
            ok: false,
            reason: 'unreadable',
            detail: `${authFilePath} is not readable JSON: ${e.message}`
        };
    }

    if (!state || typeof state !== 'object' || Array.isArray(state)) {
        return {
            ok: false,
            reason: 'malformed',
            detail: `${authFilePath} does not contain a Playwright storage-state object`
        };
    }

    if (!Array.isArray(state.cookies)) {
        return {
            ok: false,
            reason: 'malformed',
            detail: `${authFilePath} has no "cookies" array, so it is not a usable auth state`
        };
    }

    return { ok: true, reason: 'written', detail: `${authFilePath} written with ${state.cookies.length} cookie(s)` };
}

/**
 * Runs the login flow.
 *
 * Resolves `true` only when *both* halves of "logged in" held: the session
 * reached the authenticated app, and the auth state file was then written and
 * verified on disk. Resolves `false` for every failure, and never throws.
 *
 * The never-throws part is the fix for #24. This function used to catch every
 * error, log it, and resolve `undefined`, so a caller had no way to tell a
 * completed login from one that died at the password prompt — and `src/index.js`
 * had nothing to translate into a process exit code, so a failed login exited 0
 * and every shell script and CI step built on it reported success. Throwing
 * instead would fix the CLI but break the library contract: this is also
 * exported as the package main, and ms-onenote-exporter awaits it without a
 * try/catch. A boolean is the one shape that is useful to both.
 *
 * @param {object} [credentials]
 * @param {string} [credentials.email]
 * @param {string} [credentials.password]
 * @param {string} [credentials.targetUrl]
 * @param {string} [credentials.authFile]
 * @param {boolean} [credentials.notheadless]
 * @param {boolean} [credentials.dodump]
 * @param {boolean} [credentials.screenshot]
 * @returns {Promise<boolean>} true when the login succeeded and the auth file is on disk
 */
async function login(credentials = {}) {
    const { email, password, targetUrl, authFile, screenshot } = credentials;
    const isAutomated = !!(email && password);
    const headless = !credentials.notheadless && isAutomated;
    // Use targetUrl if provided, otherwise default to ONENOTE_URL for backward compatibility
    const finalTargetUrl = targetUrl || ONENOTE_URL;
    // Shared across every clearBlockingScreens() call in this login so a screen that
    // never changes is clicked once, not once per polling round.
    const blockerProgress = { signatures: new Set(), lastClickAt: 0 };

    // Get the auth file path (use provided or default)
    const filePath = getAuthFilePath(authFile);
    const metaPath = getAuthMetaFilePath(filePath);

    // Added to verify version on user's machine. Read from package.json so it
    // cannot drift away from the published version.
    logger.debug(`Authentication Module: v${PKG_VERSION} starting...`);

    logger.debug(`Using auth file path: ${filePath}`);
    logger.debug(`Using meta file path: ${metaPath}`);

    if (isAutomated) {
        logger.info(`Attempting automated login for ${email}...`);
    } else {
        logger.info('Launching browser for manual authentication...');
        logger.warn('Please log in to your Microsoft account in the browser window.');
        const serviceName = finalTargetUrl.includes('outlook') ? 'Outlook' : 'OneNote';
        logger.warn(`The script will wait until you successfully reach the ${serviceName} interface.`);
    }
    
    // Prepare files (backup existing if needed, create directory)
    await checkAndPrepareFiles(filePath);

    // Declared out here so the finally can close it, but *launched* inside the
    // try. It used to sit above it, which meant a missing Chromium — the single
    // most common setup failure — rejected out of login() before any handler
    // here ran, leaving the caller with an unhandled rejection instead of the
    // false it now gets.
    let browser = null;

    try {
        browser = await chromium.launch({ headless: !!headless });
        const context = await browser.newContext({
            // Disable WebAuthn/FIDO hardware key prompts
            ignoreHTTPSErrors: false,
        });
        const page = await context.newPage();

        // Dismiss any native browser dialogs (alert/confirm/prompt) automatically
        page.on('dialog', async dialog => {
            logger.debug(`[dialog] Auto-dismissing native dialog: type=${dialog.type()}, message="${dialog.message()}"`);
            await dialog.dismiss();
        });

        // Inject WebAuthn override BEFORE any navigation.
        // navigator.credentials.create() triggers a native OS-level dialog that
        // Playwright cannot dismiss via DOM clicks. Rejecting it programmatically
        // prevents the dialog from ever appearing, making the page-level Cancel
        // button accessible instead.
        await page.addInitScript(() => {
            if (typeof navigator !== 'undefined' && navigator.credentials) {
                navigator.credentials.create = () =>
                    Promise.reject(new DOMException('Cancelled by automation', 'NotAllowedError'));
                navigator.credentials.get = (options) => {
                    if (options && options.publicKey) {
                        return Promise.reject(new DOMException('Cancelled by automation', 'NotAllowedError'));
                    }
                    // Allow password-manager / federated credential requests through
                    return Promise.reject(new DOMException('Cancelled by automation', 'NotAllowedError'));
                };
            }
        });

        await page.goto(finalTargetUrl);

        if (isAutomated) {
            logger.step('Automating login steps...');

            // 0. Handle landing page if it appears (redirection to onenote.cloud.microsoft)
            try {
                // Look for "Sign in" button.
                const signInButton = page.getByRole('button', { name: 'Sign in' }).first();

                await signInButton.waitFor({ state: 'visible', timeout: 10000 });

                logger.info('Landing page detected. Clicking "Sign in"...');

                await signInButton.click({ noWaitAfter: true });

                logger.debug('Clicked "Sign in", waiting for login form...');
            } catch (e) {
                logger.debug('Landing page not detected or "Sign in" button not found within timeout.');
            }

            // 1. Enter Email
            try {
                await page.waitForSelector('input[name="loginfmt"]', { state: 'visible', timeout: 30000 });
                await page.fill('input[name="loginfmt"]', email);

                logger.info('Email entered. Clicking "Next"...');
                await page.click('input[type="submit"]');

                // Wait for the email form to actually be replaced, but not for the
                // loginfmt input to vanish: on login.microsoftonline.com it is
                // parked as <input name="loginfmt" class="moveOffScreen"> and stays
                // in the DOM for the rest of the login, so "hidden" never happens
                // and this used to burn the full timeout on every work account.
                // The next step appearing is the real signal that we have moved on.
                await Promise.race([
                    page.waitForSelector(PASSWORD_FIELD_SELECTOR, { state: 'visible', timeout: 15000 }).then(() => 'password'),
                    page.waitForSelector('input[name="loginfmt"]', { state: 'hidden', timeout: 15000 }).then(() => 'advanced'),
                ]).catch(() => {
                    logger.debug('Email form did not visibly change yet; the sign-in method step will wait for the next screen.');
                });

                logger.info('Will wait 1 seconds to give the UI a moment to settle into the next screen (MFA/Password)');
                await page.waitForTimeout(1000);

                const usernameError = page.locator('#usernameError');
                if (await usernameError.isVisible({ timeout: 2000 })) {
                    const errorMsg = await usernameError.textContent();
                    throw new Error(`Login Error (Username): ${errorMsg?.trim()}`);
                }
            } catch (e) {
                logger.error(`Failed to enter email: ${e.message}`);
                if (credentials.dodump) {
                    const displayPath = await dumpPage(page, 'debug_login_error_email.html', { screenshot });
                    logger.error(`Email submission failed. HTML dumped to ${displayPath}/debug_login_error_email.html`);
                }
                throw e;
            }

            // Proactive dump after email step (before MFA detection)
            if (credentials.dodump) {
                const displayPath = await dumpPage(page, 'debug_after_email.html', { screenshot });
                logger.debug(`[dodump] Post-email state dumped to ${displayPath}/debug_after_email.html`);
            }

            // 1.5. Get from the email step to the password box. This screen has
            // no "Other ways to sign in" step on it — the "Use your password" link
            // in its footer is the only route to the password — so the state is
            // read from the DOM rather than guessed from a race between text
            // selectors. See reachPasswordScreen() above.
            try {
                const nav = await reachPasswordScreen(page, { dodump: credentials.dodump, screenshot });

                logger.debug(`Sign-in method step: reached=${nav.reached} (${nav.reason}) after ${nav.steps} step(s)`);
                if (nav.state) {
                    logger.debug(`Current screen: ${shortUrl(nav.state.url)} — heading: "${nav.state.heading || '(none)'}"`);
                }

                if (!nav.reached && nav.reason === 'code_prompt') {
                    logger.warn('Microsoft is asking for a verification code instead of a password. This account cannot finish a password-only login.');
                }
            } catch (e) {
                // Never fatal: step 2 still waits for the password box and reports
                // precisely which screen is in the way if it is not there.
                logger.debug(`Sign-in method step skipped: ${e.message}`);
            }

            // 2. Enter Password
            try {
                const passwordField = page.locator(PASSWORD_FIELD_SELECTOR).first();
                try {
                    await passwordField.waitFor({ state: 'visible', timeout: 30000 });
                } catch (e) {
                    // "page.waitForSelector: Timeout 30000ms exceeded" is the least
                    // actionable error this tool can emit, and it is what a
                    // passwordless screen used to produce. Name the screen instead.
                    const stuck = await readSignInState(page) || await readScreenState(page);
                    if (stuck) {
                        const needsCode = stuck.sendCode || stuck.approveApp || stuck.otcPrompt;
                        throw new Error(
                            `Password field never appeared. Still on ${shortUrl(stuck.url)} — heading: "${stuck.heading || '(none)'}".` +
                            (needsCode
                                ? ' Microsoft is offering a code/phone approval here, not a password.'
                                : ' This screen does not offer a password sign-in.')
                        );
                    }
                    throw e;
                }

                await passwordField.fill(password);

                logger.debug('Submitting the sign-in form...');
                if (!(await submitSignInForm(page))) {
                    throw new Error('Password filled but no sign-in submit control was found.');
                }

                const passwordError = page.locator('#passwordError');
                if (await passwordError.isVisible({ timeout: 2000 })) {
                    const errorMsg = await passwordError.textContent();
                    throw new Error(`Login Error (Password): ${errorMsg?.trim()}`);
                }
            } catch (e) {
                if (credentials.dodump) {
                    const displayPath = await dumpPage(page, 'debug_login_error_password.html', { screenshot });
                    logger.error(`Password entry failed. HTML dumped to ${displayPath}/debug_login_error_password.html`);
                }
                throw e;
            }

            // Proactive dump after password submission (before post-password MFA check)
            if (credentials.dodump) {
                const displayPath = await dumpPage(page, 'debug_after_password.html', { screenshot });
                logger.debug(`[dodump] Post-password state dumped to ${displayPath}/debug_after_password.html`);
            }

            // 2.5a. Clear blocking screens (consent, proof freshness, FIDO, "Stay signed
            // in?"). All of them hijack the navigation after the password is accepted.
            // This first pass catches the ones that appear immediately; the watcher in
            // step 4 covers the rest, which is where they usually turn up.
            try {
                const cleared = await clearBlockingScreens(page, {
                    progress: blockerProgress,
                    dodump: credentials.dodump,
                    screenshot
                });
                logger.debug(`Blocking screen pass: handled=${cleared.handled} (${cleared.reason})`);
            } catch (e) {
                logger.debug(`Blocking screen pass skipped: ${e.message}`);
            }

            // 2.5b. Handle post-password MFA/Verification if needed
            try {
                // ".displaySign" is the legacy number-match element; the Fluent
                // pages put the same number under a data-testid instead.
                const NUMBER_MATCH = '.displaySign, [data-testid="displaySign"]';
                const verificationScreen = await Promise.race([
                    page.waitForSelector('text="Verify your identity"', { timeout: 10000 }).then(() => 'verify'),
                    page.waitForSelector('text="Enter code"', { timeout: 10000 }).then(() => 'enter_code'),
                    page.waitForSelector('input[name="otc"]', { timeout: 10000 }).then(() => 'otc_input'),
                    page.waitForSelector('text=/Approve sign in request/i', { timeout: 10000 }).then(() => 'number_match'),
                    page.waitForSelector(NUMBER_MATCH, { timeout: 10000 }).then(() => 'number_match'),
                ]).catch(() => null);

                if (credentials.dodump) {
                    const displayPath = await dumpPage(page, 'debug_post_password_mfa.html', { screenshot });
                    logger.debug(`[dodump] Post-password MFA screen state dumped to ${displayPath}/debug_post_password_mfa.html`);
                }

                if (verificationScreen === 'number_match') {
                    // Extracted so it can be tested. It was inline and anonymous,
                    // which is why nobody knew it waited rather than cancelled --
                    // and why this path was described as unsupported, and told to
                    // users, without anyone having run it.
                    await waitForPhoneApproval(page, { logger });
                } else if (verificationScreen) {
                    logger.warn('MFA/Verification screen detected.');
                    logger.step('A verification code is required. Please check your email or authenticator app.');

                    const code = await promptUser('Enter the verification code: ');

                    if (await page.locator('input[name="otc"]').isVisible()) {
                        await page.fill('input[name="otc"]', code);
                    } else if (await page.locator('input[type="tel"]').isVisible()) {
                        await page.fill('input[type="tel"]', code);
                    } else {
                        await page.locator('input[type="text"]:visible, input[type="tel"]:visible').first().fill(code);
                    }

                    if (!(await submitSignInForm(page))) {
                        logger.debug('No submit control found on the verification screen.');
                    }
                }
            } catch (e) {
                logger.debug(`Post-password verification handling skipped or failed: ${e.message}`);
            }

            // 2.7/3. "Help protect your account", "Stay signed in?" and the FIDO page are
            // all entries in the BLOCKING_SCREENS table: answered here, and again by the
            // watcher in step 4 for the copies that arrive after this point.

            // 4. Wait for redirection to target interface (notebooks or mail)
            // A consent screen can still show up after any of the steps above (e.g. a
            // terms update queued behind "Stay signed in?"), so keep clearing them
            // while we wait instead of only checking once up front.
            //
            // The watcher dumps and screenshots too: these are the screens that
            // arrive late, and they are exactly the ones a login dies on without
            // leaving anything behind to look at.
            let stopWatcher = false;
            const blockerWatcher = watchBlockingScreens(page, {
                progress: blockerProgress,
                maxScreens: 2,
                dodump: credentials.dodump,
                screenshot,
                shouldStop: () => stopWatcher
            });

            try {
                await Promise.race([waitForAuthSuccess(page, finalTargetUrl), blockerWatcher]);
            } catch (e) {
                // Name the screen we are stuck on: a plain timeout is the single most
                // common report for this tool and "still on X" is what makes it fixable.
                const stuck = await readScreenState(page);
                if (stuck) {
                    logger.error(`Timed out waiting for the authenticated interface. Still on ${shortUrl(stuck.url)} — heading: "${stuck.heading || '(none)'}"`);
                    if (!matchBlockingScreen(stuck)) {
                        logger.warn('That screen is not one this tool knows how to dismiss automatically.');
                    }
                }
                if (credentials.dodump) {
                    const displayPath = await dumpPage(page, 'debug_login_error_success.html', { screenshot });
                    logger.error(`Success detection failed. HTML dumped to ${displayPath}/debug_login_error_success.html`);
                }
                throw e;
            } finally {
                stopWatcher = true;
                await blockerWatcher;
            }
        } else {
            logger.warn('Login flow requires manual interaction.');
            const serviceName = finalTargetUrl.includes('outlook') ? 'Outlook' : 'OneNote';
            logger.step(`>>> Once you see your ${serviceName} interface in the browser, return here and press ENTER to continue. <<<`);

            const rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });

            await new Promise(resolve => {
                rl.question('', () => {
                    rl.close();
                    resolve();
                });
            });
        }

        logger.info('Saving authentication state...');
        await context.storageState({ path: filePath });

        await fs.writeJson(metaPath, {
            email: email || 'manual login',
            loginTime: new Date().toISOString()
        });

        // Reaching the app and writing the file are two separate claims, and
        // only the second one tells the caller there is something to log in
        // with. Reading it back is what separates "the login worked" from "the
        // login worked and the file is actually usable" — the distinction the
        // issue's exit-code rule turns on.
        const written = await verifyAuthStateFile(filePath);
        if (!written.ok) {
            logger.error(`Login reached the authenticated interface but the auth file is not usable (${written.reason}): ${written.detail}`);
            logger.error('Treating this as a failed login: there is no usable state to save.');
            return false;
        }

        logger.success(`Authentication successful! State saved to ${filePath}`);
        return true;
    } catch (error) {
        logger.error('Authentication failed or cancelled:', error);
        if (isAutomated) {
            logger.debug('Possible cause: incorrect credentials, MFA requirement, or selector change.');
        }
        return false;
    } finally {
        if (browser) {
            await browser.close().catch(() => { });
        }
    }
}

async function getAuthenticatedContext(browser, authFilePath) {
    const filePath = getAuthFilePath(authFilePath);
    if (await fs.pathExists(filePath)) {
        return browser.newContext({ storageState: filePath });
    } else {
        throw new Error('No authentication state found. Please run "login" command first.');
    }
}

/**
 * Writes one of `check`'s page states, when dumps were asked for.
 *
 * A dump is a debugging aid and may not become a verdict. Two things could make
 * it one, and both are handled here rather than at the three call sites:
 *
 *   - The write can fail on its own terms — an unwritable dump directory, a full
 *     disk. dumpPage() throws there, and the caller is inside verifyAuth()'s
 *     try block, so an unwritable dump directory would be caught as a
 *     verification error and turn a live session into `unverifiable`. A debug
 *     flag must not be able to change the answer it exists to explain.
 *   - There may be no page at all: the browser can fail to launch, and the
 *     error-path dump has to survive that.
 *
 * @param {import('playwright').Page|null} page
 * @param {string} fileName  basename, e.g. debug_check_login.html
 * @param {object} options
 * @param {boolean} options.dodump
 * @param {boolean} options.screenshot
 */
async function dumpCheckPage(page, fileName, { dodump, screenshot }) {
    if (!dodump) return;
    if (!page) {
        logger.debug(`[dodump] No page to dump for ${fileName} — the browser never got one.`);
        return;
    }
    try {
        const displayPath = await dumpPage(page, fileName, { screenshot });
        logger.debug(`[dodump] Check state dumped to ${displayPath}/${fileName}`);
    } catch (e) {
        logger.warn(`[dodump] Could not write ${fileName}: ${e.message} (the check itself is unaffected)`);
    }
}

/**
 * Says why `--dodump` produced nothing, on the paths that never load a page.
 *
 * Two of the six verdicts are decided before a browser exists, so a user who
 * asked for dumps and got none would otherwise have no way to tell that from a
 * broken dump directory — the two need completely different investigations.
 *
 * @param {boolean} dodump
 * @param {string} why
 */
function noteNothingToDump(dodump, why) {
    if (dodump) logger.info(`[dodump] Nothing to dump: ${why}.`);
}

/**
 * Checks the saved session and says which of the three possible answers it got.
 *
 * A boolean cannot carry this. "Not logged in" and "could not tell" are very
 * different to a caller, and this function has been forced to return `true` for
 * both for a good reason: on a network error the conservative move is to leave
 * the auth file alone rather than delete a possibly-valid session on the
 * strength of a DNS blip. That decision is right for the *data* and wrong for
 * the *exit code*, which is why both answers used to collapse into one `true`
 * and the `check` command could not fail.
 *
 * So the decision is split rather than reversed. verifyAuth() keeps the
 * distinction; checkAuth() below keeps the old conservative boolean, unchanged,
 * for the library callers that depend on it.
 *
 * `unverifiable` deliberately does not touch the auth file: the state on disk
 * may be perfectly good, and the only thing that failed was the check.
 *
 * The arguments are one object rather than two positionals so that the debug
 * flags ride along with the rest instead of needing a signature that grows a
 * parameter per flag — the same shape login() takes its credentials in.
 *
 * @param {object} [options]
 * @param {string} [options.targetUrl] ONENOTE_URL or OUTLOOK_URL
 * @param {string} [options.authFilePath]
 * @param {boolean} [options.dodump] write the pages this check looked at, so a
 *   surprising verdict can be read rather than guessed at
 * @param {boolean} [options.screenshot] also write a PNG beside each dump
 * @returns {Promise<{ authenticated: boolean, reason: string, detail: string }>}
 *   `reason` is one of `authenticated`, `no_auth_file`, `unusable_auth_file`,
 *   `expired`, `stayed_unauthenticated`, `unverifiable`. Only `authenticated`
 *   is true. Only `expired` deletes the auth file, because only it is
 *   Microsoft's own word that the session is gone.
 * @throws {TypeError} if called with the positional arguments it used to take
 */
async function verifyAuth(options = {}, legacyAuthFilePath) {
    // The positional form used to be verifyAuth(targetUrl, authFilePath), and a
    // caller upgrading this package would have no other way to find out: the
    // string would destructure to an undefined targetUrl and the check would
    // quietly run against OneNote and the default auth file, reporting a real
    // answer to a question about a different session. A loud TypeError is worth
    // more here than a plausible wrong result.
    if (typeof options === 'string' || legacyAuthFilePath !== undefined) {
        throw new TypeError(
            'verifyAuth() takes a single options object: verifyAuth({ targetUrl, authFilePath, dodump, screenshot }). ' +
            'The positional verifyAuth(targetUrl, authFilePath) form is no longer supported; ' +
            'checkAuth(targetUrl, authFilePath) still takes positionals.'
        );
    }

    const {
        targetUrl = ONENOTE_URL,
        authFilePath,
        dodump = false,
        screenshot = false
    } = options;
    const filePath = getAuthFilePath(authFilePath);

    if (!(await fs.pathExists(filePath))) {
        noteNothingToDump(dodump, 'there is no auth file to check');
        return { authenticated: false, reason: 'no_auth_file', detail: `${filePath} does not exist` };
    }

    // An auth file that exists but is unusable is a definite "not authenticated",
    // not an unknown: handing it to newContext() would throw, so there is
    // nothing to be uncertain about.
    const usable = await verifyAuthStateFile(filePath);
    if (!usable.ok) {
        noteNothingToDump(dodump, `the auth file is unusable (${usable.reason}), so no page was ever loaded`);
        return { authenticated: false, reason: 'unusable_auth_file', detail: usable.detail };
    }

    let browser = null;
    let page = null;
    try {
        logger.debug('Verifying authentication session...');
        browser = await chromium.launch({ headless: true });
        const context = await browser.newContext({ storageState: filePath });
        page = await context.newPage();

        await page.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: CHECK_NAV_TIMEOUT });

        // Before the probe gets to change anything. On an expired session this is
        // the last thing that is still the app shell, which is the half of the
        // story that the verdict dump below cannot tell: by the time the probe
        // answers, the shell has already been replaced by the login page.
        await dumpCheckPage(page, 'debug_check_after_nav.html', { dodump, screenshot });

        // The old check waited a fixed 2 s and asked whether the URL happened to
        // be a login host by then. That is a race with a timer, and it lost:
        // Microsoft serves the app shell first and redirects a dead session to
        // login.live.com several seconds later, so a completely empty auth file
        // was read as "authenticated" and `check` exited 0. Measured on an auth
        // file with zero cookies: the page was still on
        // onenote.cloud.microsoft/ after the 2 s, and reported as signed in.
        //
        // The redirect is not even the signal — it is only how a dead session
        // usually announces itself, and waiting for it is how the old code
        // mistook "has not happened yet" for "will not happen". So both
        // outcomes are waited on, and the first to arrive decides.
        logger.info('Waiting for Microsoft to either open the app or send this session to a login page...');
        const verdict = await settleSessionProbe(page, targetUrl, CHECK_SETTLE_TIMEOUT);
        const where = shortUrl(page.url());

        // The page that decided the answer, named for the verdict so the file
        // says which of the three it is without having to be opened — and so
        // that three different failures cannot overwrite one another's
        // evidence in the same minute, the way a single `debug_check.html`
        // would.
        await dumpCheckPage(page, `debug_check_${verdict}.html`, { dodump, screenshot });

        // The verdict decides, and nothing else does.
        //
        // It used not to: settleSessionProbe's answer was awaited and dropped on
        // the floor, after which the code went on to re-derive the answer from
        // the URL alone. Since only a redirect to a login host could then count
        // as a positive result, every live session was reported as expired —
        // deterministically, not intermittently — and `check` could only ever
        // exit 1. The probe was built to be the answer; asking it and then
        // ignoring it was the bug.
        if (verdict === 'app') {
            return {
                authenticated: true,
                reason: 'authenticated',
                detail: `the signed-in interface rendered at ${where}`
            };
        }

        // Only a login redirect is allowed to delete anything. It is Microsoft
        // stating that the session is dead. The silent case below is the absence
        // of evidence, and a live session produces it too whenever the app is
        // slower than the timeout — deleting there would throw away a working
        // login over a timeout, which is the destructive half of the bug above.
        if (verdict === 'login') {
            logger.warn('Authentication session has expired. Deleting stale auth state.');
            await logout(authFilePath);
            return {
                authenticated: false,
                reason: 'expired',
                detail: `the session was redirected to ${where}; stale auth state deleted`
            };
        }

        // Neither side arrived. The session was never actually authenticated —
        // it sat on the unauthenticated shell the whole time — and saying so is
        // the whole point of the probe above. Deliberately not `expired`: no
        // redirect means nothing proved the session dead, only that this run
        // failed to see it, and the auth file is left in place for the retry.
        return {
            authenticated: false,
            reason: 'stayed_unauthenticated',
            detail: `the signed-in interface never rendered at ${where} and the session was never sent to a login page; the auth file was left in place`
        };
    } catch (e) {
        logger.debug(`Session verification encountered an error (timeout/network): ${e.message}`);
        // The one case where a dump is the only evidence there is: the exception
        // carries a message and nothing about the page that produced it. A
        // navigation timeout, a proxy interception page or a TLS failure all land
        // here, and all look identical in the log. `page` may be null if the
        // browser never launched, which is why this goes through the same
        // null-tolerant helper rather than dumpPage() directly.
        await dumpCheckPage(page, 'debug_check_error.html', { dodump, screenshot });
        return {
            authenticated: false,
            reason: 'unverifiable',
            detail: `could not verify the session (${e.message}); the auth file was left in place`
        };
    } finally {
        // Was `logger.debug('Looks like user is logged in.')`, printed from a
        // finally on every single run — including for an expired session that
        // had just been deleted, and for a network failure. It asserted the
        // opposite of the return value on two of the three paths.
        if (browser) {
            await browser.close().catch(() => { });
        }
    }
}

/**
 * Waits for a loaded Microsoft page to reveal whether the session is real.
 *
 * Three things count as an answer, and whichever comes first wins:
 *
 *   - a login page: the session is dead, which is the case this whole function
 *     exists to catch;
 *   - the authenticated app's UI markers: the session is live;
 *   - neither, before the timeout: also dead. The app shell renders for an
 *     unauthenticated session and never navigates anywhere, so silence is
 *     itself a verdict rather than a reason to keep waiting.
 *
 * The authenticated-app signals are the UI markers login() waits for, built by
 * the same authSuccessMarkerAttempts(). A change to what counts as "signed in"
 * therefore cannot pass login's tests while leaving `check` quietly disagreeing
 * with it.
 *
 * Deliberately *not* waited for: a URL match, in either form.
 *
 *   - Polling `page.url()` is the race the old 2 s sleep lost — it mistook
 *     "has not redirected yet" for "will not redirect".
 *   - login()'s own app-path signal cannot be reused here at all: `check`
 *     navigates to ONENOTE_URL, which is /notebooks, and ONENOTE_APP_PATH
 *     matches /notebooks. Used here it would answer 'app' the instant the page
 *     loaded, never look at anything else, and reintroduce the same false
 *     positive one layer down.
 *
 * The markers cannot be fooled either way: they only render once a session is
 * real.
 *
 * Errors are swallowed by design: a timeout or a navigation that interrupts a
 * wait is not a failure here, it just means the other signal gets its turn.
 *
 * @param {import('playwright').Page} page
 * @param {string} targetUrl
 * @param {number} [timeoutMs]
 * @returns {Promise<'login'|'app'|'idle'>} which one decided it
 */
async function settleSessionProbe(page, targetUrl, timeoutMs = CHECK_SETTLE_TIMEOUT) {
    const isLoginUrl = url => /login\.(live|microsoftonline)\.com/i.test(url.hostname);
    const attempts = authSuccessMarkerAttempts(page, targetUrl, timeoutMs)
        .map(a => a.wait.then(() => 'app'));
    attempts.push(
        page.waitForURL(isLoginUrl, { timeout: timeoutMs }).then(() => 'login')
    );

    // Promise.any keeps waiting after the first rejection, so this only gives
    // up once every signal has either answered or timed out.
    return await Promise.any(attempts).catch(() => 'idle');
}

/**
 * True when there is a saved session that might work.
 *
 * Unchanged contract, including the conservative `true` on a verification
 * error: an existing caller uses this to decide whether to *keep* the auth
 * file, and losing a valid session to a transient network error would be worse
 * than proceeding with one that turns out to be dead. Use verifyAuth() when you
 * need to distinguish the two — that is what the CLI does, because an exit code
 * cannot be hedged.
 *
 * One thing did change, and it is a bug fix rather than a hedge: an auth file
 * that is not a Playwright storage state is now reported as not authenticated.
 * It previously fell into the network-error branch and read as `true`, on the
 * grounds that it might be a good session. It cannot be — nothing can open a
 * context from it.
 *
 * @param {string} [targetUrl]
 * @param {string} [authFilePath]
 * @returns {Promise<boolean>}
 */
async function checkAuth(targetUrl = ONENOTE_URL, authFilePath) {
    // Still positional, on purpose: this function's contract is documented as
    // unchanged for the library callers that depend on it, and the CLI — the
    // only surface that exposes the dump flags — calls verifyAuth() directly.
    const status = await verifyAuth({ targetUrl, authFilePath });
    return status.authenticated || status.reason === 'unverifiable';
}

module.exports = {
    login,
    // Exported so a caller can assert the number-match behaviour directly, and
    // so the path is reachable from a test rather than only from inside login().
    waitForPhoneApproval,
    getAuthenticatedContext,
    checkAuth,
    verifyAuth,
    getAuthMeta,
    logout,
    // Exported for tests and for callers that need to assert on the reason a
    // login failed, since login() reports a boolean and the details go to the log.
    verifyAuthStateFile,
    // Exported for tests: how `check` tells a live session from a dead one,
    // with a caller-supplied timeout so the assertion does not pay the
    // production wait. The old fixed 2 s sleep reported an empty auth file as
    // authenticated; this is the assertion that says it does not any more.
    settleSessionProbe,
    // Exported for tests: clears the consent/interrupt screens that Microsoft can
    // inject mid-login (e.g. the Terms of Use update at account.live.com/tou/accrue).
    clearBlockingScreens,
    // Exported for tests: the loop that polls for those screens after the early
    // pass, so "a late screen is dumped too" can be asserted without a real login.
    watchBlockingScreens,
    // Exported for tests: walks the "how do you want to sign in?" screens
    // (passwordless "Get a code to sign in", "Other ways to sign in", method
    // list) and lands on the password box.
    reachPasswordScreen,
    // Exported for tests: presses the sign-in submit control on both the legacy
    // (<input type="submit" value="Sign in">) and Fluent (<button>) pages.
    submitSignInForm,
    // Exported for tests: the single path every --dodump write goes through, so
    // that credentials cannot reach a dump file.
    dumpPage,
    // Exported for tests: the same success detection with a caller-supplied
    // timeout, so "is this the authenticated app?" can be asserted on directly
    // rather than through a 60 s wait.
    waitForAuthSuccessProbe,
    // Exported for tests: the production wait itself, with a caller-supplied
    // timeout, so the failure report can be asserted on. The probe above only
    // says whether it arrived; this is what a user actually sees when it did not.
    waitForAuthSuccess
};
