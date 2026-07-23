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
        // OneNote default behavior
        logger.info('Waiting for redirection to notebooks list...');
        await Promise.any([
            page.waitForURL(url => url.toString().includes('/notebooks') || url.hostname.includes('onenote.cloud.microsoft') || url.hostname.includes('onenote.com'), { timeout: 60000 }),
            page.waitForSelector('text="My notebooks"', { state: 'visible', timeout: 60000 }),
            page.waitForSelector('text="Create new notebook"', { state: 'visible', timeout: 60000 }),
            page.waitForSelector('text="Welcome, "', { state: 'visible', timeout: 60000 })
        ]);
        logger.success('Notebooks list detected.');
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
        await page.evaluate(() => {
            const btns = Array.from(document.querySelectorAll('button, input[type="button"], input[type="submit"]'));
            const cancel = btns.find(b => /^cancel$/i.test((b.textContent || b.value || '').trim()));
            if (cancel) cancel.click();
        });
    }

    // Wait for navigation away from the FIDO page (up to 8 s)
    try {
        await page.waitForURL(url => !url.toString().includes('/fido/'), { timeout: 8000 });
        logger.debug('FIDO: navigated away from FIDO page successfully.');
    } catch (_) {
        logger.warn('FIDO: still on FIDO URL after Cancel — continuing anyway.');
    }
}

async function login(credentials = {}) {
    const { email, password, targetUrl, authFile } = credentials;
    const isAutomated = !!(email && password);
    const headless = !credentials.notheadless && isAutomated;
    // Use targetUrl if provided, otherwise default to ONENOTE_URL for backward compatibility
    const finalTargetUrl = targetUrl || ONENOTE_URL;

    // Get the auth file path (use provided or default)
    const filePath = getAuthFilePath(authFile);
    const metaPath = getAuthMetaFilePath(filePath);

    // Added to verify version on user's machine
    logger.debug('Authentication Module: Version 4.4-DEBUG starting...');

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

            // 2.5b. Handle FIDO/security key page (login.microsoft.com/consumers/fido/create)
            // The addInitScript above makes navigator.credentials.create() reject immediately,
            // so the native OS WebAuthn dialog never appears. We only need to click the
            // page-level "Cancel" button and wait for navigation away from the FIDO URL.
            try {
                const currentUrl = page.url();
                const onFidoPage = currentUrl.includes('/fido/');

                if (onFidoPage) {
                    logger.info(`Already on FIDO page (${currentUrl}). Clicking page-level Cancel...`);
                    await dismissFidoPage(page, logger);
                } else {
                    // Race: either we navigate to fido, or 8 s passes (no fido page)
                    const fidoHandled = await Promise.race([
                        page.waitForURL(url => url.toString().includes('/fido/'), { timeout: 8000 })
                            .then(async () => {
                                logger.info(`Navigated to FIDO page: ${page.url()}. Dismissing...`);
                                await dismissFidoPage(page, logger);
                                return 'fido_cancelled';
                            }),
                        page.waitForTimeout(8000).then(() => 'no_fido'),
                    ]);
                    logger.debug(`FIDO check result: ${fidoHandled}`);
                }
            } catch (e) {
                logger.debug(`FIDO popup handler skipped: ${e.message}`);
            }

            // 2.5. Handle post-password MFA/Verification if needed
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

            // 2.7. Handle "Help protect your account" interrupt screen
            try {
                const interruptPrompt = page.getByText(/Help protect your account/i).first();
                if (await interruptPrompt.isVisible({ timeout: 5000 }) || page.url().includes('account.live.com/interrupt/')) {
                    logger.info('Detected "Help protect your account" interrupt screen.');
                    const skipButton = page.getByRole('button', { name: /Skip for now/i })
                        .or(page.getByText(/Skip for now/i))
                        .first();
                    if (await skipButton.isVisible()) {
                        logger.info('Clicking "Skip for now"...');
                        await skipButton.click();
                    }
                }
            } catch (e) {
                logger.debug(`Help protect your account interrupt screen did not appear: ${e.message}`);
            }

            // 3. Handle "Stay signed in?" prompt if it appears
            try {
                logger.debug('Checking for "Stay signed in?" prompt...');

                const staySignedIn = page.getByText(/Stay signed in?/i)
                    .or(page.locator('#KmsiDescription'))
                    .first();

                await staySignedIn.waitFor({ state: 'visible', timeout: 7000 });

                logger.info('Detected "Stay signed in?" prompt.');

                const dontShowAgain = page.locator('input[name="DontShowAgain"], #KmsiCheckboxField').first();
                if (await dontShowAgain.isVisible()) {
                    logger.debug('Checking "Don\'t show this again" checkbox...');
                    await dontShowAgain.check().catch(() => { });
                }

                const yesButton = page.getByRole('button', { name: /^Yes$/i })
                    .or(page.locator('button[data-testid="primaryButton"]'))
                    .or(page.locator('#idSIButton9'))
                    .first();

                logger.info('Clicking "Yes" to stay signed in...');
                await yesButton.click();
            } catch (e) {
                logger.debug(`Stay signed in prompt did not appear or was not recognized: ${e.message}`);
            }

            // 4. Wait for redirection to target interface (notebooks or mail)
            try {
                await waitForAuthSuccess(page, finalTargetUrl);
            } catch (e) {
                if (credentials.dodump) {
                    const dumpDir = await logger.getDumpDir();
                    const displayPath = logger.getDumpDisplayPath();
                    const debugFile = path.join(dumpDir, 'debug_login_error_success.html');
                    await fs.writeFile(debugFile, await page.content().catch(e => `<!-- Error: ${e.message} -->`));
                    logger.error(`Success detection failed. HTML dumped to ${displayPath}/debug_login_error_success.html`);
                }
                throw e;
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
    logout
};
