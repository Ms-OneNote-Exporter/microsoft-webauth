/**
 * @fileoverview This file handles user authentication logic.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const { chromium } = require('playwright');
const fs = require('fs-extra');
const logger = require('./utils/logger');
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
 * Waits for successful authentication based on target URL.
 * @param {import('playwright').Page} page - Playwright page object
 * @param {string} targetUrl - The target URL (ONENOTE_URL or OUTLOOK_URL)
 */
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
 */
const ONENOTE_APP_PATH = /\/copilotnotebooks\b|\/notebooks\b/;

/** UI markers that only render once the session is actually signed in. */
const ONENOTE_SIGNED_IN_MARKERS = [
    'text="My notebooks"',
    'text="Create new notebook"',
    'text="All Notebooks"',
    'text="Welcome, "'
];

/**
 * Waits for the authenticated app, bounded by `timeoutMs` instead of the full
 * production timeout. Returns whether it arrived, so it can be asserted on.
 * @returns {Promise<boolean>}
 */
async function waitForAuthSuccessProbe(page, targetUrl, timeoutMs) {
    const isOutlook = targetUrl && targetUrl.includes('outlook.cloud.microsoft');
    const attempts = isOutlook
        ? [
            page.waitForSelector('[aria-label*="message list"], [role="grid"][aria-label*="mail"], .messageList', { state: 'visible', timeout: timeoutMs }),
            page.waitForSelector('text=/Inbox|Sent Mail|Drafts/i', { state: 'visible', timeout: timeoutMs }),
            page.waitForSelector('div[role="row"]', { state: 'visible', timeout: timeoutMs }),
        ]
        : [
            page.waitForURL(url => ONENOTE_APP_PATH.test(url.toString()), { timeout: timeoutMs }),
            ...ONENOTE_SIGNED_IN_MARKERS.map(marker =>
                page.waitForSelector(marker, { state: 'visible', timeout: timeoutMs })),
        ];

    const won = await Promise.any(attempts.map(p => p.then(() => true))).catch(() => false);
    return won;
}

async function waitForAuthSuccess(page, targetUrl) {
    const isOutlook = targetUrl && targetUrl.includes('outlook.cloud.microsoft');

    if (isOutlook) {
        logger.info('Waiting for redirection to Outlook mail...');
        await Promise.any([
            // Outlook: wait for email message list (table with emails)
            page.waitForSelector('[aria-label*="message list"], [role="grid"][aria-label*="mail"], .messageList', { state: 'visible', timeout: 60000 }),
            // Fallback: wait for folder navigation (Inbox, Sent, etc.)
            page.waitForSelector('text=/Inbox|Sent Mail|Drafts/i', { state: 'visible', timeout: 60000 }),
            // Fallback: wait for any email-like content
            page.waitForSelector('div[role="row"]', { state: 'visible', timeout: 60000 }),
        ]);
        logger.success('Outlook mail interface detected.');
    } else {
        logger.info('Waiting for redirection to authenticated notebooks interface...');
        await Promise.any([
            // Primary: the URL must be the authenticated app, not the marketing page
            page.waitForURL(url => ONENOTE_APP_PATH.test(url.toString()), { timeout: 60000 }),
            // Fallback UI elements that only appear when actually signed in
            ...ONENOTE_SIGNED_IN_MARKERS.map(marker =>
                page.waitForSelector(marker, { state: 'visible', timeout: 60000 })),
        ]);
        logger.success('Authenticated notebooks interface detected.');
    }
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
            const displayPath = await dumpPage(page, `debug_blocking_screen_${i + 1}.html`);
            logger.debug(`[dodump] Blocking screen state dumped to ${displayPath}/debug_blocking_screen_${i + 1}.html`);
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
                if (el.getAttribute('aria-hidden') === 'true') return false;
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
 * @param {string} [options.dumpFile]        basename written under the dump dir
 * @returns {Promise<{ reached: boolean, reason: string, steps: number, state: object|null }>}
 */
async function reachPasswordScreen(page, options = {}) {
    const {
        maxSteps = 4,
        stateTimeout = 15000,
        transitionTimeout = 10000,
        dodump = false,
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
            const displayPath = await dumpPage(page, `${dumpFile}.html`);
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
 * @param {import('playwright').Page} page
 * @param {string} fileName  basename, e.g. debug_after_password.html
 */
async function dumpPage(page, fileName) {
    const dumpDir = await logger.getDumpDir();
    const displayPath = logger.getDumpDisplayPath();
    await fs.writeFile(path.join(dumpDir, fileName), await redactedPageContent(page));
    return displayPath;
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

async function login(credentials = {}) {
    const { email, password, targetUrl, authFile } = credentials;
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

    const browser = await chromium.launch({ headless: !!headless });
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

    try {
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
                    const displayPath = await dumpPage(page, 'debug_login_error_email.html');
                    logger.error(`Email submission failed. HTML dumped to ${displayPath}/debug_login_error_email.html`);
                }
                throw e;
            }

            // Proactive dump after email step (before MFA detection)
            if (credentials.dodump) {
                const displayPath = await dumpPage(page, 'debug_after_email.html');
                logger.debug(`[dodump] Post-email state dumped to ${displayPath}/debug_after_email.html`);
            }

            // 1.5. Get from the email step to the password box. This screen has
            // no "Other ways to sign in" step on it — the "Use your password" link
            // in its footer is the only route to the password — so the state is
            // read from the DOM rather than guessed from a race between text
            // selectors. See reachPasswordScreen() above.
            try {
                const nav = await reachPasswordScreen(page, { dodump: credentials.dodump });

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
                    const displayPath = await dumpPage(page, 'debug_login_error_password.html');
                    logger.error(`Password entry failed. HTML dumped to ${displayPath}/debug_login_error_password.html`);
                }
                throw e;
            }

            // Proactive dump after password submission (before post-password MFA check)
            if (credentials.dodump) {
                const displayPath = await dumpPage(page, 'debug_after_password.html');
                logger.debug(`[dodump] Post-password state dumped to ${displayPath}/debug_after_password.html`);
            }

            // 2.5a. Clear blocking screens (consent, proof freshness, FIDO, "Stay signed
            // in?"). All of them hijack the navigation after the password is accepted.
            // This first pass catches the ones that appear immediately; the watcher in
            // step 4 covers the rest, which is where they usually turn up.
            try {
                const cleared = await clearBlockingScreens(page, {
                    progress: blockerProgress,
                    dodump: credentials.dodump
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
                    const displayPath = await dumpPage(page, 'debug_post_password_mfa.html');
                    logger.debug(`[dodump] Post-password MFA screen state dumped to ${displayPath}/debug_post_password_mfa.html`);
                }

                if (verificationScreen === 'number_match') {
                    logger.warn('Number Matching MFA detected ("Approve sign in request" screen).');

                    let matchNumber = '??';
                    try {
                        matchNumber = await page.locator(NUMBER_MATCH).first().textContent().catch(() => null) || '??';
                    } catch (_) {
                        logger.debug('Could not extract the number-match code — user may still see it if --notheadless is used.');
                    }

                    logger.step('══════════════════════════════════════════════════════');
                    logger.step(`  ACTION REQUIRED: Open Microsoft Authenticator on your phone.`);
                    logger.step(`  Enter the number:  ${matchNumber.trim()}`);
                    logger.step(`  Then tap "Yes" / "Approve" in the app.`);
                    logger.step('══════════════════════════════════════════════════════');
                    logger.info('Waiting for phone approval (up to 120 seconds)...');

                    await Promise.race([
                        page.waitForSelector(NUMBER_MATCH, { state: 'hidden', timeout: 120000 }),
                        page.waitForURL(url => !url.toString().includes('login.microsoftonline.com'), { timeout: 120000 }),
                        page.waitForSelector('text=/Stay signed in/i', { timeout: 120000 }),
                    ]);

                    logger.success('Phone approval received. Continuing login flow...');

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
            let stopWatcher = false;
            const blockerWatcher = (async () => {
                while (!stopWatcher) {
                    try {
                        await clearBlockingScreens(page, {
                            progress: blockerProgress,
                            maxScreens: 2,
                            shouldStop: () => stopWatcher
                        });
                    } catch (e) {
                        logger.debug(`Blocking screen watcher error: ${e.message}`);
                    }
                    await page.waitForTimeout(1000).catch(() => {});
                }
            })();

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
                    const displayPath = await dumpPage(page, 'debug_login_error_success.html');
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

        logger.success(`Authentication successful! State saved to ${filePath}`);
    } catch (error) {
        logger.error('Authentication failed or cancelled:', error);
        if (isAutomated) {
            logger.debug('Possible cause: incorrect credentials, MFA requirement, or selector change.');
        }
    } finally {
        await browser.close();
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

async function checkAuth(targetUrl = ONENOTE_URL, authFilePath) {
    const filePath = getAuthFilePath(authFilePath);
    
    if (!(await fs.pathExists(filePath))) {
        return false;
    }

    let browser;
    try {
        logger.debug('Verifying authentication session...');
        browser = await chromium.launch({ headless: true });
        const context = await browser.newContext({ storageState: filePath });
        const page = await context.newPage();

        await page.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: 15000 });

        logger.info('Will wait 2 sec to allow client-side redirects to Microsoft login pages if session is dead/');
        await page.waitForTimeout(2000);

        const url = page.url();
        const isLoginUrl = url.includes('login.live.com') || url.includes('login.microsoftonline.com');

        if (isLoginUrl) {
            logger.warn('Authentication session has expired. Deleting stale auth state.');
            await logout(authFilePath);
            return false;
        }

        return true;
    } catch (e) {
        logger.debug(`Session verification encountered an error (timeout/network): ${e.message}`);
        return true;
    } finally {
        logger.debug(`Looks like user is logged in.`);
        if (browser) {
            await browser.close();
        }
    }
}

module.exports = {
    login,
    getAuthenticatedContext,
    checkAuth,
    getAuthMeta,
    logout,
    // Exported for tests: clears the consent/interrupt screens that Microsoft can
    // inject mid-login (e.g. the Terms of Use update at account.live.com/tou/accrue).
    clearBlockingScreens,
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
    waitForAuthSuccessProbe
};
