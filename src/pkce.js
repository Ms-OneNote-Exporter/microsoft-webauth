/**
 * @fileoverview OAuth2 Authorization Code + PKCE login (experiment).
 *
 * This is an alternative to the Playwright password flow in ./auth.js. The
 * difference that matters: the password is typed on Microsoft's own page in the
 * user's normal browser, so this process never sees it. What it does see is an
 * authorization code, exchanged for tokens.
 *
 * IMPORTANT (scope of the experiment): the tokens obtained here are Graph
 * tokens. They are NOT OneNote web session cookies. The downstream tools
 * (list-notebooks, export-notebook) drive the OneNote SPA and need
 * `storageState`, so whether these tokens can bootstrap that is a separate
 * question this module deliberately does not answer. See README "OAuth2 PKCE".
 *
 * No dependencies beyond Node builtins: `https` is used instead of `fetch` so
 * the module works identically on Node 20 with no experimental warnings.
 * @author msout
 * @copyright 2026 msout
 */
const crypto = require('crypto');
const https = require('https');
const http = require('http');
const os = require('os');
const { spawn } = require('child_process');
const fs = require('fs-extra');
const path = require('path');
const logger = require('./utils/logger');
const { ensureAuthDir } = require('./config');

const DEFAULT_PORT = 8400;
// No path component: for an app registered as "Public client/native (mobile and
// desktop)" with `http://localhost`, Microsoft matches on host and port only and
// rejects any request carrying a path. The port is free to vary; the path is not.
const DEFAULT_REDIRECT_URI = `http://localhost:${DEFAULT_PORT}`;
// Matches the recommended app registration ("Personal Microsoft accounts"),
// which is single-tenant and therefore consentable without a verified publisher.
// A multi-tenant app ('common') would require Partner Network verification before
// end users could grant consent.
const DEFAULT_TENANT = 'consumers';
const DEFAULT_SCOPES = 'openid profile email User.Read Notes.Read offline_access';
const DEFAULT_TIMEOUT_MS = 5 * 60 * 1000;

/**
 * Builds the tenant-specific endpoints.
 * @param {string} [tenant] 'common', 'organizations', 'consumers' or a tenant id
 */
function getEndpoints(tenant = DEFAULT_TENANT) {
    const base = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0`;
    return { authorize: `${base}/authorize`, token: `${base}/token` };
}

/**
 * Generates a PKCE verifier and its S256 challenge.
 *
 * RFC 7636: the verifier is 43-128 chars of unreserved characters, and the
 * challenge is BASE64URL(SHA256(ASCII(verifier))).
 * @returns {{ verifier: string, challenge: string, method: 'S256' }}
 */
function generatePkcePair() {
    // 32 random bytes -> 43 base64url chars, the RFC 7636 minimum length.
    const verifier = crypto.randomBytes(32).toString('base64url');
    const challenge = crypto.createHash('sha256').update(verifier, 'ascii').digest('base64url');
    return { verifier, challenge, method: 'S256' };
}

/**
 * Generates an opaque `state` value, used to bind the callback to this request
 * and to blunt CSRF against the local callback server.
 * @returns {string}
 */
function generateState() {
    return crypto.randomBytes(16).toString('base64url');
}

/**
 * Builds the Microsoft authorization URL the user must open.
 * @returns {string}
 */
function buildAuthorizeUrl({
    clientId,
    tenant = DEFAULT_TENANT,
    scopes = DEFAULT_SCOPES,
    challenge,
    state,
    redirectUri = DEFAULT_REDIRECT_URI,
    loginHint,
    prompt,
}) {
    if (!clientId) throw new Error('clientId is required');
    if (!challenge) throw new Error('challenge is required');
    if (!state) throw new Error('state is required');

    const url = new URL(getEndpoints(tenant).authorize);
    url.searchParams.set('client_id', clientId);
    url.searchParams.set('response_type', 'code');
    url.searchParams.set('redirect_uri', redirectUri);
    url.searchParams.set('response_mode', 'query');
    url.searchParams.set('scope', scopes);
    url.searchParams.set('code_challenge', challenge);
    url.searchParams.set('code_challenge_method', 'S256');
    url.searchParams.set('state', state);

    // `prompt=select_account` stops Microsoft from silently reusing a browser
    // session, which is what we want when verifying who actually signed in.
    if (prompt) url.searchParams.set('prompt', prompt);
    if (loginHint) url.searchParams.set('login_hint', loginHint);

    return url.toString();
}

/**
 * Opens a URL in the user's default browser.
 *
 * Deliberately the *system* browser, not Playwright: no automation should touch
 * the page where the password is typed.
 * @returns {Promise<void>}
 */
function openInDefaultBrowser(url) {
    return new Promise((resolve) => {
        let cmd;
        let args;
        if (os.platform() === 'darwin') {
            cmd = 'open';
            args = [url];
        } else if (os.platform() === 'win32') {
            cmd = 'cmd';
            // The empty "" is the window title `start` would otherwise eat.
            args = ['/c', 'start', '""', url];
        } else {
            cmd = 'xdg-open';
            args = [url];
        }
        try {
            const child = spawn(cmd, args, { detached: true, stdio: 'ignore' });
            child.on('error', (e) => {
                logger.warn(`Could not open a browser automatically (${e.message}).`);
                logger.warn('Open this URL manually to continue:');
                logger.warn(url);
            });
            child.unref();
        } catch {
            logger.warn('Open this URL manually to continue:');
            logger.warn(url);
        }
        resolve();
    });
}

/**
 * Starts a one-shot HTTP server on the loopback interface to receive the
 * redirect. Binds to 127.0.0.1 only, and is destroyed after the first hit.
 *
 * @returns {Promise<{code: string}>} resolves with the authorization code
 */
function waitForAuthorizationCode({
    redirectUri = DEFAULT_REDIRECT_URI,
    expectedState,
    timeoutMs = DEFAULT_TIMEOUT_MS,
    onUrl,
}) {
    const target = new URL(redirectUri);

    return new Promise((resolve, reject) => {
        let settled = false;
        // server.close() only stops new connections; it waits for existing
        // keep-alive sockets to drain. A browser that just loaded our success
        // page would hold the process open for seconds, so sockets are tracked
        // and destroyed on finish.
        const sockets = new Set();

        const server = http.createServer((req, res) => {
            const requestUrl = new URL(req.url, `http://${req.headers.host}`);

            // Answer once, then close: never leave this connection reusable.
            res.setHeader('Connection', 'close');

            // Once the page has been fully written, drop the socket. A keep-alive
            // connection would otherwise hold the process open after sign-in.
            res.on('finish', () => {
                if (res.socket) res.socket.destroy();
            });

            // Only answer the redirect path; 404 anything else so a stray
            // request cannot be mistaken for the real callback.
            if (requestUrl.pathname !== target.pathname) {
                res.writeHead(404, { 'Content-Type': 'text/plain' });
                res.end('Not found');
                return;
            }

            const code = requestUrl.searchParams.get('code');
            const state = requestUrl.searchParams.get('state');
            const error = requestUrl.searchParams.get('error');
            const errorDescription = requestUrl.searchParams.get('error_description');

            // Unblock the caller immediately. Waiting on server.close() would mean
            // waiting for the browser to drop the socket, which delays the CLI for
            // no benefit; the teardown happens in the 'finish' handler above.
            const finish = (fn, value) => {
                if (settled) return;
                settled = true;
                clearTimeout(timer);
                server.close();
                fn(value);
            };

            if (error) {
                res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
                res.end(errorPage('Sign-in was not completed', errorDescription || error));
                finish(reject, new Error(`Authorization failed: ${error} ${errorDescription || ''}`.trim()));
                return;
            }

            // Bind the response to this request. Without this, any local process
            // could feed us a code, and a leftover tab from an earlier sign-in
            // would be accepted as if it belonged to this one.
            if (state !== expectedState) {
                res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
                res.end(stateMismatchPage());
                finish(reject, new Error([
                    'State mismatch: the sign-in response did not match this request.',
                    'Nothing was signed in and no token was issued.',
                    'This is almost always a leftover browser tab: the sign-in you completed',
                    'was started by an earlier run of this command. Close every Microsoft',
                    'sign-in tab, then run the command once and finish the sign-in in the tab it opens.',
                ].join('\n  ')));
                return;
            }

            if (!code) {
                res.writeHead(400, { 'Content-Type': 'text/html; charset=utf-8' });
                res.end(errorPage('No code received', 'Microsoft did not return an authorization code.'));
                finish(reject, new Error('No authorization code in the callback.'));
                return;
            }

            res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
            res.end(successPage());
            finish(resolve, { code });
        });

        const timer = setTimeout(() => {
            if (settled) return;
            settled = true;
            for (const socket of sockets) socket.destroy();
            sockets.clear();
            server.close();
            reject(new Error(`Timed out after ${Math.round(timeoutMs / 1000)}s waiting for the sign-in redirect.`));
        }, timeoutMs);

        server.on('connection', (socket) => {
            sockets.add(socket);
            socket.on('close', () => sockets.delete(socket));
        });
        server.on('error', (e) => {
            if (settled) return;
            settled = true;
            clearTimeout(timer);
            reject(new Error(
                `Could not listen on ${target.port} (${e.code || e.message}). ` +
                'Pass a free port with --port if something else is using it.'
            ));
        });

        server.listen(parseInt(target.port, 10) || DEFAULT_PORT, '127.0.0.1', () => {
            logger.debug(`Callback server listening on ${redirectUri}`);
            if (onUrl) onUrl(`${target.origin}${target.pathname}`);
        });
    });
}

/**
 * Exchanges an authorization code for tokens. Public client: no secret, which
 * is what makes PKCE necessary rather than optional.
 * @returns {Promise<object>} raw token endpoint response
 */
function exchangeCodeForToken({ code, codeVerifier, clientId, tenant = DEFAULT_TENANT, redirectUri = DEFAULT_REDIRECT_URI, scopes = DEFAULT_SCOPES }) {
    if (!code) throw new Error('code is required');
    if (!codeVerifier) throw new Error('codeVerifier is required');
    if (!clientId) throw new Error('clientId is required');

    const body = new URLSearchParams({
        client_id: clientId,
        grant_type: 'authorization_code',
        code,
        redirect_uri: redirectUri,
        code_verifier: codeVerifier,
        scope: scopes,
    }).toString();

    return postForm(getEndpoints(tenant).token, body, 'application/x-www-form-urlencoded');
}

/**
 * Minimal form POST over https. Chosen over fetch so behaviour is identical on
 * Node 18/20/22 with no experimental warnings.
 * @returns {Promise<object>} parsed JSON
 */
function postForm(urlString, body, contentType) {
    return new Promise((resolve, reject) => {
        const url = new URL(urlString);
        const req = https.request(
            {
                method: 'POST',
                hostname: url.hostname,
                path: url.pathname + url.search,
                headers: {
                    'Content-Type': contentType,
                    'Content-Length': Buffer.byteLength(body),
                    Accept: 'application/json',
                },
            },
            (res) => {
                const chunks = [];
                res.on('data', (c) => chunks.push(c));
                res.on('end', () => {
                    const text = Buffer.concat(chunks).toString('utf8');
                    let parsed;
                    try {
                        parsed = JSON.parse(text);
                    } catch {
                        return reject(new Error(`Token endpoint returned non-JSON (HTTP ${res.statusCode}): ${text.slice(0, 200)}`));
                    }
                    if (res.statusCode < 200 || res.statusCode >= 300) {
                        const desc = parsed.error_description || parsed.error || `HTTP ${res.statusCode}`;
                        return reject(new Error(`Token exchange failed: ${desc}`));
                    }
                    resolve(parsed);
                });
            }
        );
        req.on('error', reject);
        req.setTimeout(30_000, () => req.destroy(new Error('Token request timed out.')));
        req.end(body);
    });
}

/**
 * Reads the claims of a JWT without verifying the signature.
 *
 * This is for display only ("who did we just sign in as"). It is NOT
 * authentication: an unverified token proves nothing, and nothing here should
 * trust its contents.
 * @returns {object|null}
 */
function decodeJwtClaims(jwt) {
    try {
        const part = String(jwt).split('.')[1];
        if (!part) return null;
        return JSON.parse(Buffer.from(part, 'base64url').toString('utf8'));
    } catch {
        return null;
    }
}

/**
 * Writes the token file with owner-only permissions.
 *
 * The refresh token is password-equivalent: anyone who reads this file can
 * mint access tokens as the signed-in user until it is revoked. So 0600, and a
 * loud warning if the mode drifted.
 * @returns {Promise<string>} the path written
 */
async function saveTokenFile(tokenFilePath, payload) {
    await ensureAuthDir(tokenFilePath);
    await fs.writeJson(tokenFilePath, payload, { spaces: 2, mode: 0o600 });
    // writeJson only applies mode when creating the file; enforce it either way
    // so an existing loose file is tightened.
    await fs.chmod(tokenFilePath, 0o600);
    return tokenFilePath;
}

/** Warns when a token file is readable by group or others. */
async function checkTokenFilePermissions(tokenFilePath) {
    try {
        const stat = await fs.stat(tokenFilePath);
        const mode = stat.mode & 0o777;
        if (mode & 0o077) {
            logger.warn(`WARNING: ${tokenFilePath} is mode ${mode.toString(8).padStart(3, '0')}.`);
            logger.warn('A refresh token is password-equivalent. Run: chmod 600 ' + tokenFilePath);
        }
    } catch {
        // No file yet; nothing to check.
    }
}

/** Resolves the token file path, defaulting next to the auth file. */
function resolveTokenFilePath(tokenFilePath, authFilePath) {
    if (tokenFilePath) return tokenFilePath;
    const base = authFilePath || path.join(os.homedir(), '.microsoft-webauth', 'auth-file.json');
    const dir = path.dirname(base);
    const name = path.basename(base).replace(/\.json$/, '');
    return path.join(dir, `${name}-token.json`);
}

function errorPage(title, detail) {
    return `<!doctype html><meta charset="utf-8"><title>${title}</title>
<body style="font:16px system-ui;margin:4rem auto;max-width:34rem">
<h1 style="font-size:1.4rem">${title}</h1>
<p>${detail}</p>
<p>You can close this tab.</p></body>`;
}

/**
 * The overwhelmingly common cause of a state mismatch is a stale tab, not an
 * attack, so the page leads with the fix instead of the jargon.
 * @returns {string}
 */
function stateMismatchPage() {
    return `<!doctype html><meta charset="utf-8"><title>Start the sign-in again</title>
<body style="font:16px system-ui;margin:4rem auto;max-width:34rem;line-height:1.5">
<h1 style="font-size:1.4rem">This sign-in tab is out of date</h1>
<p>The sign-in you just completed was started by an <strong>earlier run</strong> of the
command, so this tool is not expecting it and has discarded it. Nothing was signed in
and no token was issued.</p>
<p>To continue:</p>
<ol>
  <li>Close every Microsoft sign-in tab in this browser.</li>
  <li>Run the command again.</li>
  <li>Finish the sign-in in the tab it opens, without opening a second one.</li>
</ol>
<p>You can close this tab.</p></body>`;
}

function successPage() {
    return `<!doctype html><meta charset="utf-8"><title>Signed in</title>
<body style="font:16px system-ui;margin:4rem auto;max-width:34rem">
<h1 style="font-size:1.4rem">Signed in</h1>
<p>You can close this tab and return to the terminal.</p></body>`;
}

/**
 * Fails fast on a redirect URI shape that Microsoft will reject.
 *
 * A native/public client registered as `http://localhost` is matched on host
 * and port only, so a path component is invalid. That rejection happens in the
 * browser, after the CLI is already waiting on the callback, which otherwise
 * means a silent multi-minute timeout with no explanation.
 *
 * @returns {string|null} an error message, or null when the URI is acceptable
 */
function validateRedirectUri(redirectUri) {
    let url;
    try {
        url = new URL(redirectUri);
    } catch {
        return `Redirect URI is not a valid URL: ${redirectUri}`;
    }

    if (url.protocol !== 'http:' && url.protocol !== 'https:') {
        return `Redirect URI must be http or https, got "${url.protocol}".`;
    }

    // url.pathname is '/' for a bare origin and for an explicit "/".
    if (url.pathname && url.pathname !== '/') {
        return [
            `Redirect URI must not contain a path (found "${url.pathname}").`,
            'An app registered as "Public client/native (mobile and desktop)" with http://localhost',
            'is matched on host and port only, so Microsoft rejects any path. Use',
            `http://localhost:${url.port || DEFAULT_PORT} instead.`,
        ].join('\n  ');
    }

    if (url.search) {
        return `Redirect URI must not contain a query string (found "${url.search}").`;
    }

    return null;
}

/**
 * Runs the full authorization-code + PKCE login.
 *
 * Sequence: generate a verifier/challenge pair, open the authorize URL in the
 * system browser, wait for the local callback, then exchange the code for
 * tokens and write them to a 0600 file.
 *
 * @param {object} options
 * @param {string} options.clientId        public client id of the registered app
 * @param {string} [options.tenant]
 * @param {string} [options.scopes]
 * @param {string} [options.redirectUri]
 * @param {string} [options.tokenFile]     where to write tokens
 * @param {string} [options.authFile]      used to derive the default token path
 * @param {string} [options.loginHint]     pre-fills the account picker
 * @param {number} [options.timeoutMs]
 * @param {boolean} [options.openBrowser]  set false when driving a remote headless box
 * @returns {Promise<object>} the token response
 */
async function loginWithPkce({
    clientId,
    tenant = DEFAULT_TENANT,
    scopes = DEFAULT_SCOPES,
    redirectUri = DEFAULT_REDIRECT_URI,
    tokenFile,
    authFile,
    loginHint,
    timeoutMs = DEFAULT_TIMEOUT_MS,
    openBrowser = true,
}) {
    if (!clientId) {
        throw new Error(
            'Missing client id. Pass --client-id <id> or set MSOUT_CLIENT_ID.\n' +
            'Register a public client (Mobile and desktop applications) at ' +
            'https://portal.azure.com -> Microsoft Entra ID -> App registrations -> New registration.'
        );
    }

    const redirectProblem = validateRedirectUri(redirectUri);
    if (redirectProblem) {
        throw new Error(
            `Invalid --redirect-uri.\n  ${redirectProblem}\n` +
            'Microsoft validates this on the sign-in page, so a bad value shows up as a\n' +
            'browser error while this command waits for a redirect that never arrives.'
        );
    }

    const tokenFilePath = resolveTokenFilePath(tokenFile, authFile);
    const { verifier, challenge } = generatePkcePair();
    const state = generateState();

    const authorizeUrl = buildAuthorizeUrl({
        clientId,
        tenant,
        scopes,
        challenge,
        state,
        redirectUri,
        loginHint,
    });

    logger.info('OAuth2 PKCE login (experimental).');
    logger.info("A browser tab will open on Microsoft's own sign-in page.");
    logger.info('Your password is entered on Microsoft and is never seen by this tool.');
    logger.info(`Waiting for the redirect on ${redirectUri} (timeout ${Math.round(timeoutMs / 1000)}s)...`);
    logger.warn('Complete the sign-in in the tab that opens below. If you left a Microsoft');
    logger.warn('sign-in tab open from an earlier attempt, close it first: this run only');
    logger.warn('accepts the response to the sign-in it starts now.');

    if (openBrowser) {
        await openInDefaultBrowser(authorizeUrl);
    } else {
        logger.info('Open this URL manually:');
        logger.info(authorizeUrl);
    }

    // Start listening around the moment the browser opens so a fast redirect is
    // not missed.
    const { code } = await waitForAuthorizationCode({ redirectUri, expectedState: state, timeoutMs });
    logger.success('Authorization code received. Exchanging for tokens...');

    const token = await exchangeCodeForToken({ code, codeVerifier: verifier, clientId, tenant, redirectUri, scopes });

    const now = Date.now();
    await saveTokenFile(tokenFilePath, {
        acquiredAt: now,
        tenant,
        clientId,
        scopes,
        // Access tokens are short-lived. This expiry is a convenience for a human
        // reading the file, never something to trust for an auth decision.
        expiresAt: token.expires_in ? now + (token.expires_in * 1000) : null,
        ext_expires_in: token.ext_expires_in || null,
        scope: token.scope || null,
        token_type: token.token_type || 'Bearer',
        access_token: token.access_token,
        refresh_token: token.refresh_token || null,
        id_token: token.id_token || null,
    });
    await checkTokenFilePermissions(tokenFilePath);

    const claims = decodeJwtClaims(token.id_token);
    if (claims) {
        const who = claims.preferred_username || claims.email || claims.upn || claims.name;
        if (who) logger.success(`Signed in as: ${who}`);
    }

    logger.success(`Tokens written to ${tokenFilePath} (mode 600).`);
    if (!token.refresh_token) {
        logger.warn('No refresh token was issued. The "offline_access" scope is required for one.');
    }
    if (!token.id_token) {
        logger.warn('No id_token was issued. The "openid profile email" scopes are required for one.');
    }

    return token;
}

module.exports = {
    DEFAULT_PORT,
    DEFAULT_REDIRECT_URI,
    DEFAULT_TENANT,
    DEFAULT_SCOPES,
    DEFAULT_TIMEOUT_MS,
    getEndpoints,
    generatePkcePair,
    generateState,
    buildAuthorizeUrl,
    validateRedirectUri,
    openInDefaultBrowser,
    waitForAuthorizationCode,
    exchangeCodeForToken,
    postForm,
    decodeJwtClaims,
    saveTokenFile,
    checkTokenFilePermissions,
    resolveTokenFilePath,
    loginWithPkce,
};
