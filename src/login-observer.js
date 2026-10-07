/**
 * Login reasons and the observer callback.
 *
 * Two things a caller of `login()` currently cannot get, and both are additive:
 *
 * 1. **Why a login failed.** `login()` resolves a boolean and every detail goes
 *    to the log, so a caller cannot tell a wrong password from a Microsoft
 *    outage. It matters because the two need different advice, and guessing is
 *    worse than saying "something failed" — telling someone their correct
 *    password is wrong sends them to reset it.
 *
 * 2. **That a challenge is pending.** A verification-code prompt is currently
 *    only visible on stdout. A non-interactive caller has no way to know one is
 *    waiting, so the only thing it can do is hang or fail.
 *
 * ## Why `login()` still returns a boolean
 *
 * The obvious design is `login()` resolving `{ ok, reason }`. **That would break
 * the package's own CLI and every library caller**, because an object is truthy
 * whether or not it is ok:
 *
 *     const ok = await login(credentials);
 *     if (!ok) reportFailure(...);   // auth.js's index.js:102
 *
 * Returning `{ ok: false }` makes that check pass, so **a failed login is
 * reported as a success** — silently, with an exit code that says otherwise.
 *
 * So the boolean stays, and the structured result is delivered through
 * `onEvent` as a terminal `login-result`. One mechanism rather than two, and
 * zero chance of a caller that stops working by being upgraded.
 *
 * ## Why the reasons are exported
 *
 * `LOGIN_REASONS` is the closed union, frozen, for the same reason the api
 * exports its `EVENT_TYPES`: a caller can assert mechanically that every reason
 * it can observe has a mapping on its side. When a reason is added here, that
 * caller's test goes red and tells it, rather than a new cause being silently
 * folded into a generic message.
 */

const logger = require('./utils/logger');

const LOGIN_REASONS = Object.freeze([
    /** Reached the password field. The happy path's intermediate state. */
    'password_field',
    /** A screen was present but could not be read at all. */
    'unreadable',
    /** The screen did not change across a poll round. */
    'unchanged',
    /** Gave up after `maxSteps` attempts to clear the screen. */
    'max_steps',
    /** Microsoft is asking for a one-time code. */
    'code_prompt',
    /** Microsoft Authenticator is asking for a push approval. */
    'approver_prompt',
    /** No route from this screen to a password field. */
    'no_password_route',
    /** The page settled but the service did not accept the sign-in. */
    'credentials_rejected',
    /** The interface loaded but no auth state could be written or read back. */
    'auth_state_unusable',
    /** The page could not be reached at all. */
    'network',
    /** The overall attempt exceeded its deadline. */
    'timed_out',
    /** A run reached a screen that is neither a password nor a challenge. */
    'interstitial',
    /**
     * An error escaped that this package does not recognise.
     *
     * **Deliberately not a nicer name for one of the reasons above.** Every
     * other value names a screen that was actually observed, and a caller can act
     * on it — retry, tell the user to check their password, stop asking. This one
     * says only "it failed in a way I cannot describe", and the honest response
     * to that is a generic message.
     *
     * It exists because the alternative is picking the *closest* reason for a
     * failure that matched none of them, which turns one generic failure into
     * several specific lies. A caller that has no mapping for it should have a
     * default path anyway; this makes the default reachable.
     *
     * It was added before publication rather than after, because the set is
     * frozen and exported: adding to it later is the kind of change that reads as
     * additive and breaks a caller that switches exhaustively.
     */
    'unknown',
]);

/**
 * The `kind` a `challenge` can carry.
 *
 * **Two values, not three, on purpose.** The three things a user may have to do
 * — read a number, type a code, or just tap approve — are not three kinds of
 * screen. Reading a number and tapping approve are the *same* screen: Microsoft
 * shows the number on the push prompt and in Authenticator, and webauth cannot
 * tell them apart even in principle, because `waitForPhoneApproval` returns the
 * sentinel `'??'` both when the screen carries no number and when it carried one
 * that could not be read (`phone-approval.js`). A three-value union would assert
 * a distinction this package does not have, and it would assert it on exactly
 * the path nobody has run.
 *
 * So the number's presence is a *field*, and the user's action is derived from it:
 * `number !== null` means "read this and enter it in Authenticator", `null`
 * means "there is nothing to read, tap approve".
 */
const CHALLENGE_KINDS = Object.freeze([
    /** Microsoft wants a one-time code typed into the page. */
    'code',
    /** Microsoft Authenticator is waiting; the user acts on their own phone. */
    'phone-approval',
]);

/**
 * An error that carries the reason `login()` failed.
 *
 * `login()` throws ordinary `Error`s in a dozen places, so the reason cannot be
 * recovered from the message — classifying on wording is how a reworded log line
 * silently becomes "wrong password", which is the one misclassification that
 * sends a user to reset a password they typed correctly. The throw sites tag
 * themselves instead.
 *
 * Extending `Error` rather than replacing it keeps `instanceof Error` true for
 * the existing catch in `login()`, and for any caller that inspects it.
 */
class LoginError extends Error {
    /**
     * @param {string} reason  one of `LOGIN_REASONS`
     * @param {string} message  the human-readable text, unchanged from what the
     *   log already prints
     */
    constructor(reason, message) {
        super(message);
        this.name = 'LoginError';
        this.reason = reason;
    }
}

/**
 * The reason to report for an error that escaped `login()`.
 *
 * A tagged `LoginError` is believed. Anything else becomes `'unknown'` rather
 * than a guess: the log already carries the message and the stack, so nothing is
 * lost by declining to name a cause we cannot evidence.
 *
 * @param {unknown} error
 * @returns {string} a member of `LOGIN_REASONS`
 */
function reasonForError(error) {
    if (error instanceof LoginError && LOGIN_REASONS.includes(error.reason)) {
        return error.reason;
    }
    return 'unknown';
}

/**
 * An observer callback. Every event is optional and omitted-safe: a caller that
 * passes nothing sees today's behaviour exactly, including the stdin prompt.
 *
 * Emitted by `login()`:
 *
 *   challenge          { kind, label, timeoutMs, number }
 *                        a prompt is waiting. `number` is the value to enter on
 *                        the user's phone, or null when there is nothing to read
 *                        — see `CHALLENGE_KINDS` for why that is one kind and
 *                        not two. `timeoutMs` is a real deadline, or null when
 *                        the wait genuinely has none.
 *   challenge-seen     {}                             the prompt was answered
 *   challenge-expired  {}                             it was not answered
 *   login-result       { ok, reason }                 terminal; always fires.
 *                        `reason` is null when `ok`, and a member of
 *                        `LOGIN_REASONS` when not — never a value that is both.
 * Deliberately NOT emitted here: `login-started`, `login-success`, `login-failed`
 * and `auth-state`. Those are statements about a *session*, which this package
 * does not own — it knows what it saw on screen, and the caller decides what that
 * means. Emitting them here would be inventing vocabulary for a state we cannot
 * see.
 *
 * `kind` is one of `CHALLENGE_KINDS`.
 */
const LOGIN_EVENT_TYPES = Object.freeze([
    'challenge',
    'challenge-seen',
    'challenge-expired',
    'login-result',
]);

/**
 * `emit` wraps the caller's callback so a throwing or absent observer cannot
 * break a login.
 *
 * An observer that throws is a bug in the observer, and the worst possible
 * outcome is that it turns a working sign-in into a failed one. So the throw is
 * swallowed and logged — the login continues, because the observer is watching,
 * not participating.
 */
function makeEmitter(onEvent) {
    if (typeof onEvent !== 'function') return () => {};
    return (type, payload) => {
        try {
            onEvent({ type, ...payload });
        } catch (error) {
            logger.warn(`onEvent observer threw for "${type}"; continuing.`, error);
        }
    };
}

module.exports = {
    LOGIN_REASONS,
    LOGIN_EVENT_TYPES,
    CHALLENGE_KINDS,
    LoginError,
    reasonForError,
    makeEmitter
};
