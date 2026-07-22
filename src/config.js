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
const ONENOTE_URL = 'https://onenote.cloud.microsoft/en-us';
const OUTLOOK_URL = 'https://outlook.cloud.microsoft/mail/';

module.exports = {
    DEFAULT_AUTH_FILE,
    getAuthMetaFilePath,
    ensureAuthDir,
    ONENOTE_URL,
    OUTLOOK_URL,
};
