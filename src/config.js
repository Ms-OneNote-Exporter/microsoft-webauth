/**
 * @fileoverview Returns the directory where auth files (auth.json, auth-meta.json) are stored.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const path = require('path');
const os = require('os');
const fs = require('fs-extra');

/**
 * Returns the default auth file path (~/.microsoft-webauth/auth-file.json)
 */
function getDefaultAuthFilePath() {
    const homeDir = os.homedir();
    const appDir = path.join(homeDir, '.microsoft-webauth');
    return path.join(appDir, 'auth-file.json');
}

/**
 * Generates the meta file path from auth file path
 * e.g., /path/to/auth-file.json -> /path/to/auth-file-meta.json
 */
function getAuthMetaFilePath(authFilePath) {
    const dir = path.dirname(authFilePath);
    const name = path.basename(authFilePath);
    // Remove .json extension if present, then add -meta.json
    const baseName = name.replace(/\.json$/, '');
    return path.join(dir, `${baseName}-meta.json`);
}

/**
 * Ensures the directory for the auth file exists, creates it if needed
 */
async function ensureAuthDir(authFilePath) {
    const dir = path.dirname(authFilePath);
    await fs.ensureDir(dir);
}

const DEFAULT_AUTH_FILE = getDefaultAuthFilePath();
/**
 * Where an automated login starts.
 *
 * This is the pre-rebrand path, and that is deliberate. Microsoft 365 Copilot
 * moved the authenticated app to /copilotnotebooks, but /notebooks still
 * redirects there, so this stays the entry point: a login entered here is
 * detected as authenticated a few seconds later. Confirmed on v0.1.4 — the
 * 2026-09-29 run in src/logs/app.log entered at this URL, cleared a "Stay
 * signed in?" prompt and reported "Authenticated notebooks interface detected"
 * seven seconds later.
 *
 * Success detection accepts both spellings, because it has to: see
 * ONENOTE_APP_PATH in auth.js, where the same move is the reason a successful
 * login once timed out for a full 60 s and never saved its auth state. Only
 * the entry point is pinned here. Switching it to /copilotnotebooks would work
 * today too, but it would trade a working alias for a path that can move again.
 */
const ONENOTE_URL = 'https://onenote.cloud.microsoft/notebooks';
const OUTLOOK_URL = 'https://outlook.cloud.microsoft/mail/';

module.exports = {
    DEFAULT_AUTH_FILE,
    getAuthMetaFilePath,
    ensureAuthDir,
    ONENOTE_URL,
    OUTLOOK_URL,
};
