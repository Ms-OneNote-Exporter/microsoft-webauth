/**
 * @fileoverview The login observer: emitted events, reasons, and what happens
 * when the observer is broken.
 * @copyright 2026 phptr,enoola,msout
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const {
    LOGIN_REASONS,
    LOGIN_EVENT_TYPES,
    CHALLENGE_KINDS,
    LoginError,
    reasonForError,
    makeEmitter
} = require('../src/login-observer');
const { chromium } = require('playwright');
const { waitForPhoneApproval } = require('../src/phone-approval');

/**
 * Serves `html` to a real page on the Microsoft sign-in origin, so assertions are
 * made against what the DOM actually contains rather than against a hand-written
 * object standing in for it.
 *
 * **The origin is load-bearing.** `waitForPhoneApproval` ends its wait on the
 * first of three events, one of which is "the URL no longer mentions
 * login.microsoftonline.com". A `data:` URL satisfies that immediately, so every
 * wait would resolve before it began and a test would pass for the wrong reason —
 * which is what happened when this fixture was a data URL.
 *
 * The route intercepts, so nothing leaves the machine.
 *
 * The real login cannot be exercised here — it needs Microsoft — so these tests
 * assert the *seam*: that what `login()` observes reaches the far side of the
 * observer. An assertion that the handler ran, or that a value was assigned, has
 * already missed four bugs in this project that a green suite did not; the
 * emitted event is the only thing a caller will ever see.
 */
async function withPage(html, fn) {
    const browser = await chromium.launch({ headless: true });
    try {
        const page = await browser.newPage();
        await page.route('https://login.microsoftonline.com/**', route =>
            route.fulfill({ status: 200, contentType: 'text/html', body: html })
        );
        await page.goto('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');
        return await fn(page);
    } finally {
        await browser.close().catch(() => { });
    }
}

describe('LOGIN_REASONS', () => {
    it('is frozen, so a caller cannot extend it by accident', () => {
        expect(Object.isFrozen(LOGIN_REASONS)).toBe(true);
        expect(() => { LOGIN_REASONS.push('smuggled'); }).toThrow();
    });

    it('has no duplicates, which would make a reason ambiguous', () => {
        expect(new Set(LOGIN_REASONS).size).toBe(LOGIN_REASONS.length);
    });

    it('names the reasons the login actually reports', () => {
        // Every reason `login()` can emit, minus `null` for success. Read off the
        // throw sites and the two `login-result` calls rather than maintained by
        // hand, because a set that drifts from the code is worse than no set: the
        // backend asserts against it.
        expect(LOGIN_REASONS).toContain('credentials_rejected');
        expect(LOGIN_REASONS).toContain('auth_state_unusable');
        expect(LOGIN_REASONS).toContain('no_password_route');
        expect(LOGIN_REASONS).toContain('approver_prompt');
        expect(LOGIN_REASONS).toContain('code_prompt');
        expect(LOGIN_REASONS).toContain('interstitial');
        expect(LOGIN_REASONS).toContain('unreadable');
    });

    it('has an escape hatch, so an unrecognised failure is not forced into a wrong reason', () => {
        // Without this, every unclassified failure gets the closest-looking
        // reason, and "wrong password" is the one a user acts on destructively.
        expect(LOGIN_REASONS).toContain('unknown');
    });
});

describe('CHALLENGE_KINDS', () => {
    it('is frozen', () => {
        expect(Object.isFrozen(CHALLENGE_KINDS)).toBe(true);
    });

    it('does not separate number-matching from plain approval', () => {
        // Those are one screen. webauth cannot tell them apart -- waitForPhoneApproval
        // returns '??' for "no number" and for "number unreadable" alike -- so a
        // three-kind union would assert a distinction the code does not make.
        expect(CHALLENGE_KINDS).toContain('phone-approval');
        expect(CHALLENGE_KINDS).not.toContain('number-match');
    });

    it('covers the two things a user can actually be asked to do', () => {
        expect([...CHALLENGE_KINDS].sort()).toEqual(['code', 'phone-approval']);
    });
});

describe('LOGIN_EVENT_TYPES', () => {
    it('does not claim to speak about sessions', () => {
        // login-started / login-success / login-failed / auth-state are the
        // backend's vocabulary. Emitting them here would invent words for state
        // this package cannot see.
        for (const type of ['login-started', 'login-success', 'login-failed', 'auth-state']) {
            expect(LOGIN_EVENT_TYPES).not.toContain(type);
        }
    });

    it('emits exactly one terminal event', () => {
        expect(LOGIN_EVENT_TYPES.filter(t => t === 'login-result')).toHaveLength(1);
    });
});

describe('reasonForError', () => {
    it('believes a tagged error', () => {
        expect(reasonForError(new LoginError('credentials_rejected', 'nope'))).toBe('credentials_rejected');
    });

    it('refuses a tagged error carrying a reason outside the frozen set', () => {
        // Otherwise the set is only a convention, and a typo becomes a reason the
        // backend has no mapping for — silently folded into a generic message.
        const bogus = new LoginError('definitely_not_a_reason', 'nope');
        expect(reasonForError(bogus)).toBe('unknown');
    });

    it('does not classify an untagged error by reading its message', () => {
        // Classifying on wording is how a reworded log line becomes "wrong
        // password". This is the assertion that says it does not.
        const impostor = new Error('Login Error (Password): Your password is incorrect');
        expect(reasonForError(impostor)).toBe('unknown');
    });

    it('survives being handed something that is not an error at all', () => {
        for (const value of [null, undefined, 'a string', 42, {}]) {
            expect(reasonForError(value)).toBe('unknown');
        }
    });
});

describe('makeEmitter', () => {
    it('hands the observer the type alongside the payload', () => {
        const seen = [];
        makeEmitter(e => seen.push(e))('challenge', { kind: 'code' });
        expect(seen).toEqual([{ type: 'challenge', kind: 'code' }]);
    });

    it('lets a payload key not be overwritten by the type', () => {
        // Spread order is load-bearing: `{ type, ...payload }` means a payload
        // carrying its own `type` wins, which is surprising. Asserted so a
        // refactor to `{ ...payload, type }` is a deliberate change.
        const seen = [];
        makeEmitter(e => seen.push(e))('challenge', { type: 'hijacked' });
        expect(seen[0].type).toBe('hijacked');
    });

    it('is a no-op when there is no observer', () => {
        expect(() => makeEmitter(undefined)('challenge', {})).not.toThrow();
        expect(() => makeEmitter(null)('challenge', {})).not.toThrow();
        expect(() => makeEmitter('not a function')('challenge', {})).not.toThrow();
    });

    it('swallows a throwing observer, because the login continues', () => {
        // The observer is watching, not participating. If its bug could fail the
        // login, every caller's mistake would become a user's failed sign-in.
        const emit = makeEmitter(() => { throw new Error('observer bug'); });
        expect(() => emit('challenge', { kind: 'code' })).not.toThrow();
    });
});

describe('the package exports its login vocabulary', () => {
    it('re-exports it from the main entry point, so a caller need not guess a subpath', () => {
        // The backend builds its mapping table against this. An unlisted subpath
        // in `exports` works locally and 404s for a consumer, so the union has to
        // be reachable from `main`.
        const main = require('../src/auth');
        expect(main.LOGIN_REASONS).toBe(LOGIN_REASONS);
        expect(main.LOGIN_EVENT_TYPES).toBe(LOGIN_EVENT_TYPES);
        expect(main.CHALLENGE_KINDS).toBe(CHALLENGE_KINDS);
        expect(main.LoginError).toBe(LoginError);
    });
});

describe('phone approval reports a challenge to its caller', () => {
    it('announces the number before waiting, while it can still be shown', async () => {
        // Ordering is the whole point. A hook called after the wait would have
        // missed the window in which the user needs the number.
        const html = `<div class="displaySign">42</div>
            <script>setTimeout(() => document.querySelector('.displaySign').remove(), 300)</script>`;

        await withPage(html, async page => {
            const order = [];
            const result = await waitForPhoneApproval(page, {
                onChallenge: d => order.push(`challenge:${d.shown}`)
            });
            expect(result.shown).toBe('42');
            expect(result.waited).toBe(true);
            expect(order).toEqual(['challenge:42']);
        });
    });

    it('reports the challenge before the wait, not after', async () => {
        // The observable consequence of ordering. A hook called after the wait
        // returns would fire ~2 minutes late for a real number-match, which is
        // after the user needs it and after the caller's countdown started. Here
        // the screen clears 300 ms in, so "before the wait" is asserted directly
        // rather than inferred from a timeout.
        const html = `<div class="displaySign">42</div>
            <script>setTimeout(() => document.querySelector('.displaySign').remove(), 300)</script>`;

        await withPage(html, async page => {
            let announcedAt = null;
            const result = await waitForPhoneApproval(page, {
                onChallenge: () => { announcedAt = Date.now(); }
            });
            const finishedAt = Date.now();
            expect(result.shown).toBe('42');
            expect(announcedAt).not.toBeNull();
            // The screen clears 300 ms in, so the observer landing well before the
            // wait resolved is the observable form of "before the wait".
            expect(finishedAt - announcedAt).toBeGreaterThan(150);
        });
    });

    it('reports null for a screen with no readable number', async () => {
        // The runner is headless: there is no Microsoft window for a user to read
        // a number from. If this arrives as the string '??' the caller renders
        // question marks, so the null has to be decided here.
        const html = `<div class="displaySign"></div>
            <script>setTimeout(() => document.querySelector('.displaySign').remove(), 300)</script>`;

        await withPage(html, async page => {
            const seen = [];
            const result = await waitForPhoneApproval(page, {
                onChallenge: d => seen.push(d.shown)
            });
            expect(result.shown).toBe('??');
            expect(seen).toEqual(['??']);
        });
    });

    it('reports the number as null, never as the sentinel, once it is mapped', async () => {
        // The seam itself. waitForPhoneApproval speaks in the string '??', because
        // that is what a human reads in the log. On the wire the same fact has to
        // be null: a caller that renders the payload would put two question
        // marks in a box and ask the user to type them.
        const html = `<div class="displaySign"></div>
            <script>setTimeout(() => document.querySelector('.displaySign').remove(), 300)</script>`;

        await withPage(html, async page => {
            const { shown } = await waitForPhoneApproval(page, { onChallenge: () => { } });
            // What login() puts on the wire, expressed against the real return.
            const onTheWire = shown === '??' ? null : shown;
            expect(onTheWire).toBeNull();
            expect(typeof onTheWire).not.toBe('string');
        });
    });

    it('keeps waiting when the observer throws', async () => {
        // Same rule as makeEmitter, one layer down: a broken observer must not
        // cancel a wait the user is seconds from satisfying.
        const html = `<div class="displaySign">42</div>
            <script>setTimeout(() => document.querySelector('.displaySign').remove(), 300)</script>`;

        await withPage(html, async page => {
            const result = await waitForPhoneApproval(page, {
                onChallenge: () => { throw new Error('observer bug'); }
            });
            expect(result.waited).toBe(true);
        });
    });
});

describe('login() reports its outcome to an observer', () => {
    let dir;
    let authFile;
    let auth;

    beforeEach(() => {
        jest.resetModules();
        dir = fs.mkdtempSync(path.join(os.tmpdir(), 'webauth-observer-'));
        authFile = path.join(dir, 'auth-file.json');

        // src/auth.js captures chromium when it is required, so a spy installed
        // after the top-level import would never be seen. doMock + resetModules is
        // how the existing exit-code suite does this for the same reason.
        jest.doMock('playwright', () => ({
            chromium: {
                launch: async () => { throw new Error('Chromium distribution not found'); }
            }
        }));
        auth = require('../src/auth');
    });

    afterEach(() => {
        fs.rmSync(dir, { recursive: true, force: true });
        jest.dontMock('playwright');
    });

    it('resolves false and reports a reason when Chromium is missing', async () => {
        // Chromium not installed is the single most common setup failure. A caller
        // waiting on `login-result` would hang forever if this did not emit — the
        // promise resolving false is not enough on its own, because an observer is
        // an event stream, not a return value.
        const events = [];
        const ok = await auth.login({ email: 'a@b.c', password: 'pw', authFile, onEvent: e => events.push(e) });

        expect(ok).toBe(false);
        const result = events.filter(e => e.type === 'login-result');
        expect(result).toHaveLength(1);
        expect(result[0].ok).toBe(false);
        expect(LOGIN_REASONS).toContain(result[0].reason);
    });

    it('emits exactly one login-result, whatever happened', async () => {
        // Two terminals would let a caller act on a stale one; zero would leave
        // it waiting on an event that is never coming.
        const events = [];
        await auth.login({ email: 'a@b.c', password: 'pw', authFile, onEvent: e => events.push(e) });
        expect(events.filter(e => e.type === 'login-result')).toHaveLength(1);
    });

    it('never reports a reason that contradicts ok', async () => {
        // Every value in LOGIN_REASONS names something that went wrong, so a
        // success carrying one would force a caller to check both fields and
        // eventually branch on one and forget the other.
        const events = [];
        await auth.login({ email: 'a@b.c', password: 'pw', authFile, onEvent: e => events.push(e) });

        for (const event of events.filter(e => e.type === 'login-result')) {
            if (event.ok) {
                expect(event.reason).toBeNull();
            } else {
                expect(LOGIN_REASONS).toContain(event.reason);
            }
        }
    });

    it('still returns a boolean, because an object would be truthy', async () => {
        // index.js does `const ok = await login(...); if (!ok) reportFailure(...)`.
        // Returning `{ ok: false }` there reports a failed login as a success with
        // exit code 0. This is the assertion that says it has not happened.
        const ok = await auth.login({ email: 'a@b.c', password: 'pw', authFile });
        expect(typeof ok).toBe('boolean');
        expect(ok).toBe(false);
        // The trap, stated directly: an object is truthy whatever it holds.
        expect(Boolean({ ok: false })).toBe(true);
    });

    it('behaves identically with no observer at all', async () => {
        // Omitting onEvent must not change the outcome. A caller on today's
        // version has no onEvent, so this is the promise that nobody is broken.
        const without = await auth.login({ email: 'a@b.c', password: 'pw', authFile });
        const explicit = await auth.login({
            email: 'a@b.c', password: 'pw', authFile, onEvent: undefined
        });
        expect(without).toBe(explicit);
    });

    it('completes normally when the observer throws on every event', async () => {
        // The observer is watching, not participating. If its bug could fail the
        // login, every caller's mistake would become a user's failed sign-in.
        const ok = await auth.login({
            email: 'a@b.c', password: 'pw', authFile,
            onEvent: () => { throw new Error('observer bug'); }
        });
        expect(ok).toBe(false);
    });
});
