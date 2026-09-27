/**
 * @fileoverview Tests for the OAuth2 PKCE helpers.
 * Everything here runs offline: no network, no browser, no credentials.
 * @author msout
 */
const crypto = require('crypto');
const path = require('path');
const os = require('os');
const fs = require('fs-extra');
const {
    generatePkcePair,
    generateState,
    buildAuthorizeUrl,
    getEndpoints,
    decodeJwtClaims,
    waitForAuthorizationCode,
    saveTokenFile,
    resolveTokenFilePath,
} = require('../src/pkce');

describe('generatePkcePair', () => {
    test('produces a verifier within the RFC 7636 length range', () => {
        const { verifier } = generatePkcePair();
        expect(verifier.length).toBeGreaterThanOrEqual(43);
        expect(verifier.length).toBeLessThanOrEqual(128);
    });

    test('verifier uses only unreserved characters', () => {
        const { verifier } = generatePkcePair();
        expect(verifier).toMatch(/^[A-Za-z0-9._~-]+$/);
    });

    test('challenge is BASE64URL(SHA256(verifier))', () => {
        const { verifier, challenge } = generatePkcePair();
        const expected = crypto.createHash('sha256').update(verifier, 'ascii').digest('base64url');
        expect(challenge).toBe(expected);
    });

    test('challenge is not the verifier itself', () => {
        const { verifier, challenge } = generatePkcePair();
        expect(challenge).not.toBe(verifier);
    });

    test('is not deterministic', () => {
        expect(generatePkcePair().verifier).not.toBe(generatePkcePair().verifier);
    });
});

describe('generateState', () => {
    test('is url-safe and unique across calls', () => {
        const a = generateState();
        const b = generateState();
        expect(a).toMatch(/^[A-Za-z0-9_-]+$/);
        expect(a).not.toBe(b);
    });
});

describe('getEndpoints', () => {
    test('defaults to the consumers tenant used by the recommended registration', () => {
        const { authorize, token } = getEndpoints();
        expect(authorize).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize');
        expect(token).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/token');
    });

    test('builds v2.0 endpoints for the common tenant', () => {
        const { authorize, token } = getEndpoints('common');
        expect(authorize).toBe('https://login.microsoftonline.com/common/oauth2/v2.0/authorize');
        expect(token).toBe('https://login.microsoftonline.com/common/oauth2/v2.0/token');
    });

    test('honours an explicit tenant', () => {
        expect(getEndpoints('contoso.onmicrosoft.com').authorize)
            .toContain('/contoso.onmicrosoft.com/');
    });
});

describe('buildAuthorizeUrl', () => {
    const base = { clientId: 'fake-client-id', challenge: 'challenge', state: 'state' };

    test('sets every parameter PKCE requires', () => {
        const url = new URL(buildAuthorizeUrl(base));
        expect(url.searchParams.get('client_id')).toBe('fake-client-id');
        expect(url.searchParams.get('response_type')).toBe('code');
        expect(url.searchParams.get('code_challenge')).toBe('challenge');
        expect(url.searchParams.get('code_challenge_method')).toBe('S256');
        expect(url.searchParams.get('state')).toBe('state');
        expect(url.searchParams.get('response_mode')).toBe('query');
    });

    test('requests a code, not tokens, and never a secret', () => {
        const qs = buildAuthorizeUrl(base);
        expect(qs).toContain('response_type=code');
        // A public client has no secret; sending one would be a design error.
        expect(qs).not.toContain('client_secret');
    });

    test('includes the redirect uri and scopes', () => {
        const url = new URL(buildAuthorizeUrl({ ...base, redirectUri: 'http://localhost:8400/callback', scopes: 'Notes.Read' }));
        expect(url.searchParams.get('redirect_uri')).toBe('http://localhost:8400/callback');
        expect(url.searchParams.get('scope')).toBe('Notes.Read');
    });

    test('adds login_hint and prompt only when provided', () => {
        const plain = new URL(buildAuthorizeUrl(base));
        expect(plain.searchParams.get('login_hint')).toBeNull();
        expect(plain.searchParams.get('prompt')).toBeNull();

        const hinted = new URL(buildAuthorizeUrl({ ...base, loginHint: 'a@b.com', prompt: 'select_account' }));
        expect(hinted.searchParams.get('login_hint')).toBe('a@b.com');
        expect(hinted.searchParams.get('prompt')).toBe('select_account');
    });

    test('refuses to build an incomplete request', () => {
        expect(() => buildAuthorizeUrl({ challenge: 'c', state: 's' })).toThrow(/clientId/);
        expect(() => buildAuthorizeUrl({ clientId: 'c', state: 's' })).toThrow(/challenge/);
        expect(() => buildAuthorizeUrl({ clientId: 'c', challenge: 'c' })).toThrow(/state/);
    });
});

describe('decodeJwtClaims', () => {
    test('reads claims for display', () => {
        const payload = Buffer.from(JSON.stringify({ preferred_username: 'a@b.com' })).toString('base64url');
        expect(decodeJwtClaims(`h.${payload}.s`).preferred_username).toBe('a@b.com');
    });

    test('returns null instead of throwing on junk', () => {
        expect(decodeJwtClaims('not-a-jwt')).toBeNull();
        expect(decodeJwtClaims(undefined)).toBeNull();
    });
});

describe('waitForAuthorizationCode', () => {
    const redirectUri = 'http://localhost:8451/callback';

    test('resolves with the code when state matches', async () => {
        const pending = waitForAuthorizationCode({ redirectUri, expectedState: 'st4te', timeoutMs: 5000 });
        await new Promise((r) => setTimeout(r, 100));
        const res = await fetch(`${redirectUri}?code=the-code&state=st4te`);
        expect(res.status).toBe(200);
        await expect(pending).resolves.toEqual({ code: 'the-code' });
    });

    test('rejects a mismatched state', async () => {
        const pending = waitForAuthorizationCode({ redirectUri, expectedState: 'expected', timeoutMs: 5000 });
        // Attach the expectation before triggering the request: the server
        // rejects as soon as it sees the forged state, so a handler attached
        // afterwards would be too late.
        const assertion = expect(pending).rejects.toThrow(/State mismatch/);
        await new Promise((r) => setTimeout(r, 100));
        const res = await fetch(`${redirectUri}?code=the-code&state=forged`);
        expect(res.status).toBe(400);
        await assertion;
    });

    test('surfaces a denied consent as an error', async () => {
        const pending = waitForAuthorizationCode({ redirectUri, expectedState: 'st4te', timeoutMs: 5000 });
        const assertion = expect(pending).rejects.toThrow(/access_denied/);
        await new Promise((r) => setTimeout(r, 100));
        await fetch(`${redirectUri}?error=access_denied&error_description=User+declined`);
        await assertion;
    });

    test('times out rather than waiting forever', async () => {
        const pending = waitForAuthorizationCode({ redirectUri, expectedState: 'st4te', timeoutMs: 250 });
        await expect(pending).rejects.toThrow(/Timed out/);
    });

    test('404s an unrelated path instead of treating it as the callback', async () => {
        const pending = waitForAuthorizationCode({ redirectUri, expectedState: 'st4te', timeoutMs: 800 });
        await new Promise((r) => setTimeout(r, 100));
        const res = await fetch('http://localhost:8451/somewhere-else?code=x&state=st4te');
        expect(res.status).toBe(404);
        // The real callback must still work afterwards.
        const ok = await fetch(`${redirectUri}?code=real&state=st4te`);
        expect(ok.status).toBe(200);
        await expect(pending).resolves.toEqual({ code: 'real' });
    });
});

describe('saveTokenFile', () => {
    let dir;

    beforeEach(async () => {
        dir = await fs.mkdtemp(path.join(os.tmpdir(), 'pkce-test-'));
    });

    afterEach(async () => {
        await fs.remove(dir);
    });

    test('writes the file with owner-only permissions', async () => {
        const target = path.join(dir, 'auth-file-token.json');
        await saveTokenFile(target, { access_token: 'a' });
        const mode = (await fs.stat(target)).mode & 0o777;
        expect(mode).toBe(0o600);
    });

    test('tightens permissions on a pre-existing loose file', async () => {
        const target = path.join(dir, 'auth-file-token.json');
        await fs.writeJson(target, { access_token: 'old' });
        await fs.chmod(target, 0o644);
        await saveTokenFile(target, { access_token: 'new' });
        const mode = (await fs.stat(target)).mode & 0o777;
        expect(mode).toBe(0o600);
    });

    test('creates the parent directory', async () => {
        const target = path.join(dir, 'nested', 'deeper', 'token.json');
        await saveTokenFile(target, { access_token: 'a' });
        expect(await fs.pathExists(target)).toBe(true);
    });
});

describe('resolveTokenFilePath', () => {
    test('derives the token path from the auth file', () => {
        expect(resolveTokenFilePath(null, '/tmp/auth.json')).toBe('/tmp/auth-token.json');
        expect(resolveTokenFilePath(null, '/tmp/auth-file.json')).toBe('/tmp/auth-file-token.json');
    });

    test('an explicit path always wins', () => {
        expect(resolveTokenFilePath('/tmp/custom.json', '/tmp/auth.json')).toBe('/tmp/custom.json');
    });
});
