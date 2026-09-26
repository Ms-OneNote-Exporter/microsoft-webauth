/**
 * @fileoverview This file handles user authentication logic.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const { chromium } = require('playwright');
const fs = require('fs-extra');
const logger = require('./utils/logger');
const { DEFAULT_AUTH_FILE, getAuthMetaFilePath, ensureAuthDir, ONENOTE_URL } = require('./config');
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
        // OneNote default behavior — wait for the authenticated notebooks interface.
        // IMPORTANT: We must require /notebooks in the URL to avoid matching the
        // unauthenticated marketing landing page (onenote.cloud.microsoft/en-us)
        // which also matches the old hostname-only check and caused premature auth saving.
        logger.info('Waiting for redirection to authenticated notebooks interface...');
        await Promise.any([
            // Primary: URL must contain /notebooks (authenticated app)
            page.waitForURL(url => url.toString().includes('/notebooks'), { timeout: 60000 }),
            // Fallback UI elements that only appear when actually signed in
            page.waitForSelector('text="My notebooks"', { state: 'visible', timeout: 60000 }),
            page.waitForSelector('text="Create new notebook"', { state: 'visible', timeout: 60000 }),
            page.waitForSelector('text="Welcome, "', { state: 'visible', timeout: 60000 }),
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
            const dumpDir = await logger.getDumpDir();
            const displayPath = logger.getDumpDisplayPath();
            const debugFile = path.join(dumpDir, `debug_blocking_screen_${i + 1}.html`);
            await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
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

    // Added to verify version on user's machine
    logger.debug('Authentication Module: Version 4.5-DEBUG starting...');

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

                logger.debug('Waiting for email field to disappear...');
                await page.waitForSelector('input[name="loginfmt"]', { state: 'hidden', timeout: 15000 }).catch(() => {
                    logger.debug('Email field still present, proceeding with caution.');
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
                    const dumpDir = await logger.getDumpDir();
                    const displayPath = logger.getDumpDisplayPath();
                    const debugFile = path.join(dumpDir, 'debug_login_error_email.html');
                    await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
                    logger.error(`Email submission failed. HTML dumped to ${displayPath}/debug_login_error_email.html`);
                }
                throw e;
            }

            // Proactive dump after email step (before MFA detection)
            if (credentials.dodump) {
                const dumpDir = await logger.getDumpDir();
                const displayPath = logger.getDumpDisplayPath();
                const debugFile = path.join(dumpDir, 'debug_after_email.html');
                await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
                logger.debug(`[dodump] Post-email state dumped to ${displayPath}/debug_after_email.html`);
            }

            // 1.5. Handle intermediate screens (MFA selection, "Other ways to sign in")
            try {
                const pageTitle = (await page.title()).trim();
                const pageHeading = (await page.locator('h1, [role="heading"]').first().textContent().catch(() => '')).trim();

                logger.debug(`Settled State: Title="${pageTitle}" | Heading="${pageHeading}"`);
                logger.debug('Checking for intermediate MFA/Sign-in option screens...');

                const result = await Promise.race([
                    page.waitForSelector('text=/Other ways to sign in/i', { state: 'visible', timeout: 15000 }).then(() => 'other_ways'),
                    page.waitForSelector('text=/Get a code to sign in/i', { state: 'visible', timeout: 15000 }).then(() => 'other_ways'),
                    page.waitForSelector('text=/Verify your identity/i', { state: 'visible', timeout: 15000 }).then(() => 'other_ways'),
                    page.waitForSelector('text=/Use your password/i', { state: 'visible', timeout: 15000 }).then(() => 'use_password'),
                    page.waitForSelector('text=/Approve a request on my Microsoft Authenticator app/i', { state: 'visible', timeout: 5000 }).then(() => 'approve_app'),
                    page.waitForSelector('input[name="passwd"]', { state: 'visible', timeout: 15000 }).then(() => 'password'),
                    page.waitForFunction(() => {
                        const h = document.querySelector('h1, [role="heading"]')?.textContent || '';
                        return h.includes('Get a code') || h.includes('Verify your identity');
                    }, { timeout: 15000 }).then(() => 'other_ways'),
                ]).catch((err) => {
                    logger.debug(`Detection race timed out or failed: ${err.message}`);
                    return 'timeout';
                });

                logger.debug(`Intermediate screen detection result: ${result}`);

                if (credentials.dodump) {
                    const dumpDir = await logger.getDumpDir();
                    const displayPath = logger.getDumpDisplayPath();
                    const debugFile = path.join(dumpDir, 'debug_intermediate_screen.html');
                    await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
                    logger.debug(`[dodump] Intermediate screen state dumped to ${displayPath}/debug_intermediate_screen.html`);
                }

                if (result === 'other_ways' || pageHeading.includes('Get a code') || pageHeading.includes('Verify your identity')) {
                    logger.info('Detected MFA/Verification screen. Attempting to locate "Other ways to sign in"...');

                    const otherWays = page.getByRole('button', { name: /Other ways to sign in|Sign in another way/i })
                        .or(page.getByText(/Other ways to sign in|Sign in another way/i))
                        .first();

                    try {
                        logger.debug('Waiting for "Other ways" link to appear in DOM...');
                        await otherWays.waitFor({ state: 'attached', timeout: 15000 });

                        const isVisible = await otherWays.isVisible();
                        logger.debug(`"Other ways" link visibility: ${isVisible}`);

                        logger.info('Clicking "Other ways to sign in"...');
                        try {
                            await otherWays.click({ timeout: 5000 });
                        } catch (e) {
                            logger.debug(`Standard click failed, trying forced: ${e.message}`);
                            await otherWays.click({ force: true, timeout: 5000 });
                        }
                    } catch (e) {
                        logger.warn(`MFA link interaction failed: ${e.message}`);

                        logger.debug('Attempting final fallback: JavaScript-based click...');
                        const clicked = await page.evaluate(() => {
                            const elements = Array.from(document.querySelectorAll('span, a, button'));
                            const target = elements.find(el =>
                                el.textContent.toLowerCase().includes('other ways to sign in') ||
                                el.textContent.toLowerCase().includes('sign in another way')
                            );
                            if (target) {
                                target.click();
                                return true;
                            }
                            return false;
                        });

                        if (clicked) {
                            logger.info('Successfully triggered click via JavaScript fallback.');
                        } else if (pageHeading.includes('Get a code')) {
                            throw new Error('STUCK: "Other ways to sign in" link not found even via JS scan.');
                        }
                    }

                    logger.debug('Waiting for method selection screen ("Use your password")...');
                    const subResult = await Promise.race([
                        page.waitForSelector('text=/Use your password/i', { state: 'visible', timeout: 15000 }).then(() => 'use_password'),
                        page.waitForSelector('#idA_PWD_SwitchToPassword', { state: 'visible', timeout: 15000 }).then(() => 'use_password'),
                        page.waitForSelector('text=/Select a verification method/i', { state: 'visible', timeout: 15000 }).then(() => 'other_ways_list'),
                    ]).catch(() => 'timeout');

                    logger.debug(`Sub-screen detection result: ${subResult}`);

                    if (subResult === 'use_password') {
                        logger.info('Selecting "Use your password" option...');
                        await page.click('text=/Use your password/i');
                    } else if (subResult === 'other_ways_list') {
                        logger.info('Selection list detected. Looking for "Password"...');
                        await page.click('text=/Password|Use your password/i');
                    }
                } else if (result === 'use_password') {
                    logger.info('Detected "Use your password" option. Clicking...');
                    await page.click('text="Use your password"');
                } else if (result === 'approve_app') {
                    logger.warn('MFA notification already sent. Attempting to switch to password...');
                    const otherLink = page.locator('text="Other ways to sign in", #signInAnotherWay').first();
                    if (await otherLink.isVisible()) {
                        await otherLink.click();
                        await page.waitForSelector('text="Use your password"', { state: 'visible', timeout: 10000 });
                        await page.click('text="Use your password"');
                    }
                } else if (result === 'password') {
                    logger.debug('Direct password field detected.');
                } else if (result === 'timeout') {
                    logger.debug('No intermediate screen detected within timeout. Proceeding to password entry.');
                }
            } catch (e) {
                logger.debug(`Intermediate screen handler encountered a fatal issue: ${e.message}`);
            }

            // 2. Enter Password
            try {
                await page.waitForSelector('input[name="passwd"]', { state: 'visible', timeout: 30000 });
                await page.fill('input[name="passwd"]', password);

                const submitButton = page.locator('input[type="submit"], button[type="submit"]').filter({ hasText: /Sign in|Next|Finish/i }).first();

                logger.debug('Waiting for submit button to be enabled...');
                await submitButton.waitFor({ state: 'visible', timeout: 10000 });
                if (await submitButton.isDisabled()) {
                    logger.debug('Submit button is disabled. It might be the wrong one or the password field is not considered filled.');
                    logger.info('Will wait 1 seconds to let the submit button load properly');
                    await page.waitForTimeout(1000);
                }

                await submitButton.click();

                const passwordError = page.locator('#passwordError');
                if (await passwordError.isVisible({ timeout: 2000 })) {
                    const errorMsg = await passwordError.textContent();
                    throw new Error(`Login Error (Password): ${errorMsg?.trim()}`);
                }
            } catch (e) {
                if (credentials.dodump) {
                    const dumpDir = await logger.getDumpDir();
                    const displayPath = logger.getDumpDisplayPath();
                    const debugFile = path.join(dumpDir, 'debug_login_error_password.html');
                    await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
                    logger.error(`Password entry failed. HTML dumped to ${displayPath}/debug_login_error_password.html`);
                }
                throw e;
            }

            // Proactive dump after password submission (before post-password MFA check)
            if (credentials.dodump) {
                const dumpDir = await logger.getDumpDir();
                const displayPath = logger.getDumpDisplayPath();
                const debugFile = path.join(dumpDir, 'debug_after_password.html');
                await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
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
                const verificationScreen = await Promise.race([
                    page.waitForSelector('text="Verify your identity"', { timeout: 10000 }).then(() => 'verify'),
                    page.waitForSelector('text="Enter code"', { timeout: 10000 }).then(() => 'enter_code'),
                    page.waitForSelector('input[name="otc"]', { timeout: 10000 }).then(() => 'otc_input'),
                    page.waitForSelector('text=/Approve sign in request/i', { timeout: 10000 }).then(() => 'number_match'),
                    page.waitForSelector('.displaySign', { timeout: 10000 }).then(() => 'number_match'),
                ]).catch(() => null);

                if (credentials.dodump) {
                    const dumpDir = await logger.getDumpDir();
                    const displayPath = logger.getDumpDisplayPath();
                    const debugFile = path.join(dumpDir, 'debug_post_password_mfa.html');
                    await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
                    logger.debug(`[dodump] Post-password MFA screen state dumped to ${displayPath}/debug_post_password_mfa.html`);
                }

                if (verificationScreen === 'number_match') {
                    logger.warn('Number Matching MFA detected ("Approve sign in request" screen).');

                    let matchNumber = '??';
                    try {
                        matchNumber = await page.$eval('.displaySign', el => el.textContent.trim());
                    } catch (_) {
                        logger.debug('Could not extract number from .displaySign — user may still see it if --notheadless is used.');
                    }

                    logger.step('══════════════════════════════════════════════════════');
                    logger.step(`  ACTION REQUIRED: Open Microsoft Authenticator on your phone.`);
                    logger.step(`  Enter the number:  ${matchNumber}`);
                    logger.step(`  Then tap "Yes" / "Approve" in the app.`);
                    logger.step('══════════════════════════════════════════════════════');
                    logger.info('Waiting for phone approval (up to 120 seconds)...');

                    await Promise.race([
                        page.waitForSelector('.displaySign', { state: 'hidden', timeout: 120000 }),
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

                    await page.click('input[type="submit"]');
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
                    const dumpDir = await logger.getDumpDir();
                    const displayPath = logger.getDumpDisplayPath();
                    const debugFile = path.join(dumpDir, 'debug_login_error_success.html');
                    await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
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
    clearBlockingScreens
};
