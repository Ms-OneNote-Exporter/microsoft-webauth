#!/usr/bin/env node
/**
 * @fileoverview Main CLI interface to handle authentication commands.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const { program } = require('commander');
const logger = require('./utils/logger');
const { login, checkAuth, getAuthMeta, logout } = require('./auth');
const { loginWithPkce, DEFAULT_REDIRECT_URI, DEFAULT_TENANT, DEFAULT_SCOPES, DEFAULT_TIMEOUT_MS } = require('./pkce');
const { DEFAULT_AUTH_FILE, ONENOTE_URL, OUTLOOK_URL } = require('./config');
const { version: PKG_VERSION } = require('../package.json');

program
    .name('webauth')
    .description('Microsoft web authentication via Playwright — extracted from MSOneNote Exporter')
    .version(PKG_VERSION);

program
    .command('login')
    .description('Authenticate with Microsoft Account')
    .option('--email <email>', 'Microsoft account email')
    .option('--password <password>', 'Microsoft account password')
    .option('--notheadless', 'Run in visible browser mode for debugging')
    .option('--dodump', 'Dump HTML content to files for debugging')
    .option('--against <target>', 'Target service: onenote (default) or outlook', 'onenote')
    .option('--auth-file <path>', 'Path to auth file (default: ~/.microsoft-webauth/auth-file.json)', DEFAULT_AUTH_FILE)
    .action(async (options) => {
        const targetUrl = options.against === 'outlook' ? OUTLOOK_URL : ONENOTE_URL;
        await login({ ...options, targetUrl });
    });

program
    .command('login-pkce')
    .description('Experimental: OAuth2 authorization-code + PKCE login (tokens, not Playwright session cookies)')
    .option('--client-id <id>', 'Public client id of your app registration (or set MSOUT_CLIENT_ID)')
    .option('--tenant <tenant>', `Directory tenant: common, organizations, consumers, or a tenant id`, DEFAULT_TENANT)
    .option('--scopes <scopes>', 'Space-separated scopes', DEFAULT_SCOPES)
    .option('--redirect-uri <uri>', 'Must be registered on the app', DEFAULT_REDIRECT_URI)
    .option('--token-file <path>', `Where to write tokens (default: alongside the auth file, name-token.json)`)
    .option('--auth-file <path>', 'Auth file, only used to derive the default token path', DEFAULT_AUTH_FILE)
    .option('--login-hint <email>', 'Pre-fills the account picker')
    .option('--timeout <seconds>', 'How long to wait for the sign-in redirect', String(Math.round(DEFAULT_TIMEOUT_MS / 1000)))
    .option('--no-open-browser', 'Print the URL instead of opening a browser (remote/headless machines)')
    .action(async (options) => {
        try {
            await loginWithPkce({
                clientId: options.clientId || process.env.MSOUT_CLIENT_ID,
                tenant: options.tenant,
                scopes: options.scopes,
                redirectUri: options.redirectUri,
                tokenFile: options.tokenFile,
                authFile: options.authFile,
                loginHint: options.loginHint,
                timeoutMs: parseInt(options.timeout, 10) * 1000,
                openBrowser: options.openBrowser !== false,
            });
        } catch (e) {
            logger.error('PKCE login failed:', e.message);
            process.exit(1);
        }
    });

program
    .command('check')
    .description('Check if authenticated')
    .option('--against <target>', 'Target service: onenote (default) or outlook', 'onenote')
    .option('--auth-file <path>', 'Path to auth file (default: ~/.microsoft-webauth/auth-file.json)', DEFAULT_AUTH_FILE)
    .action(async (options) => {
        const targetUrl = options.against === 'outlook' ? OUTLOOK_URL : ONENOTE_URL;
        const isAuth = await checkAuth(targetUrl, options.authFile);
        if (isAuth) {
            logger.success('Authentication file found. You are authenticated.');
            const meta = await getAuthMeta(options.authFile);
            if (meta && meta.email) {
                const loginTime = new Date(meta.loginTime).toLocaleString();
                logger.info(`Logged in as: ${meta.email}`);
                logger.debug(`Session started at: ${loginTime}`);
            }
        } else {
            logger.error('Authentication file NOT found or invalid. Run "login" first.');
        }
    });

program
    .command('logout')
    .description('Clear authentication state')
    .option('--auth-file <path>', 'Path to auth file (default: ~/.microsoft-webauth/auth-file.json)', DEFAULT_AUTH_FILE)
    .action(async (options) => {
        await logout(options.authFile);
        logger.success('Logged out successfully. Authentication state cleared.');
    });

program.parse();
