/**
 * Waiting for a phone approval on the "Approve sign in request" screen.
 *
 * Extracted from the middle of `login()` so it can be tested. It was inline and
 * anonymous until now, which is the whole reason nobody knew whether it worked:
 *
 *   - the only test touching an approval screen asserted the **pre-password**
 *     reason (`approver_prompt` from `reachPasswordScreen`), which is a
 *     different code path entirely
 *   - so the post-password wait — the one that decides whether a user with
 *     Authenticator number matching can sign in — had no coverage at all
 *
 * A capability nobody has executed is a claim nobody has checked, and this one
 * ended up being described as unsupported in a user-facing document on the
 * strength of reading `dismissFidoPage` rather than reading this.
 *
 * ## What this does, and what it deliberately does not do
 *
 * It extracts the number, shows it, and **waits**. It does not type the number
 * anywhere and it does not click Approve — both of those are the user's job, on
 * their phone, which is the entire point of the flow.
 *
 * `dismissFidoPage()` cancels the FIDO2 hardware-key screen, and that is correct
 * and unrelated: a page asking for a security key has nothing to wait for.
 */

/** How long to wait for the user to act on their phone. */
const PHONE_APPROVAL_TIMEOUT_MS = 120000;

/**
 * The number-match marker, as a locator.
 *
 * `.displaySign` is the legacy server-rendered element; the Fluent pages put the
 * same number under a `data-testid`. Both are matched because Microsoft ships
 * both shapes and has no versioning between them.
 */
const NUMBER_MATCH_SELECTOR = '.displaySign, [data-testid="displaySign"]';

/**
 * waitForPhoneApproval shows the number and waits for the user to act.
 *
 * Resolves once the screen clears, the login leaves the Microsoft host, or the
 * "Stay signed in" interstitial appears — whichever happens first. A timeout is
 * **not** an error: the caller continues to the redirect wait, which is the
 * honest outcome for a user who never answered their phone.
 *
 * `onChallenge` is called once, as soon as `shown` is known and before the wait
 * begins. It is a hook rather than a return value because the number has to be
 * announced while the challenge is still outstanding: a caller that learns it
 * afterwards has already missed the window it needed to display it in.
 *
 * @param {(detail: {shown: string}) => void} [options.onChallenge]
 * @returns {Promise<{ shown: string, waited: boolean }>} `shown` is the number
 *   displayed, or `'??'` when it could not be read.
 */
async function waitForPhoneApproval(page, { logger, onChallenge } = {}) {
    const log = logger || console;

    log.warn?.('Number Matching MFA detected ("Approve sign in request" screen).');

    // Best-effort read of the number. A failure here must not abandon the wait:
    // the user can still see the number on their own phone, which is exactly what
    // the `--notheadless` path relies on.
    let shown = '??';
    try {
        const text = await page
            .locator(NUMBER_MATCH_SELECTOR)
            .first()
            .textContent()
            .catch(() => null);
        if (typeof text === 'string' && text.trim() !== '') shown = text.trim();
    } catch (_) {
        log.debug?.('Could not extract the number-match code — the user may still see it on their phone.');
    }

    log.step?.('══════════════════════════════════════════════════════');
    log.step?.('  ACTION REQUIRED: Open Microsoft Authenticator on your phone.');
    log.step?.(`  Enter the number:  ${shown}`);
    log.step?.('  Then tap "Yes" / "Approve" in the app.');
    log.step?.('══════════════════════════════════════════════════════');
    log.info?.(`Waiting for phone approval (up to ${PHONE_APPROVAL_TIMEOUT_MS / 1000} seconds)...`);

    // Announced after the log block so the terminal still leads with the number
    // for an interactive user. Swallowed for the same reason as the read above:
    // an observer's bug must not cancel a wait the user is about to satisfy.
    try {
        onChallenge?.({ shown });
    } catch (_) {
        log.debug?.('Challenge observer threw; continuing to wait for approval.');
    }

    let waited = false;
    try {
        await Promise.race([
            page.waitForSelector(NUMBER_MATCH_SELECTOR, { state: 'hidden', timeout: PHONE_APPROVAL_TIMEOUT_MS }),
            page.waitForURL((url) => !url.toString().includes('login.microsoftonline.com'), { timeout: PHONE_APPROVAL_TIMEOUT_MS }),
            page.waitForSelector('text=/Stay signed in/i', { timeout: PHONE_APPROVAL_TIMEOUT_MS }),
        ]);
        waited = true;
    } catch (_) {
        // Timed out. Not an error: the user did not answer their phone within two
        // minutes. The caller carries on to the redirect wait, which reports the
        // real outcome rather than this guess.
        waited = false;
    }

    if (waited) log.success?.('Phone approval received. Continuing login flow...');
    return { shown, waited };
}

module.exports = { waitForPhoneApproval, NUMBER_MATCH_SELECTOR, PHONE_APPROVAL_TIMEOUT_MS };