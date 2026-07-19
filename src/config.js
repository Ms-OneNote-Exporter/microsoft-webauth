/**
 * @fileoverview Returns the directory where auth files (auth.json, auth-meta.json) are stored.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const path = require('path');

/**
 * Returns the directory where auth files (auth.json, auth-meta.json) are stored.
 * For this standalone CLI tool, we use the project root.
 */
function getUserDataDir() {
    return path.resolve(__dirname, '..');
}

const USER_DATA_DIR = getUserDataDir();

const AUTH_FILE = path.join(USER_DATA_DIR, 'auth.json');
const ONENOTE_URL = 'https://onenote.cloud.microsoft/en-us';
const OUTLOOK_URL = 'https://outlook.cloud.microsoft/mail/';

module.exports = {
    AUTH_FILE,
    ONENOTE_URL,
    OUTLOOK_URL,
    USER_DATA_DIR,
};
