#!/usr/bin/env node
const { program } = require('commander');
const logger = require('./utils/logger');
const { login, checkAuth, getAuthMeta, logout } = require('./auth');
const { ONENOTE_URL, OUTLOOK_URL } = require('./config');

program
    .name('webauth')
    .description('Microsoft web authentication via Playwright — extracted from MSOneNote Exporter')
    .version('1.0.0');

program
    .command('login')
    .description('Authenticate with Microsoft Account')
    .option('--email <email>', 'Microsoft account email')
    .option('--password <password>', 'Microsoft account password')
    .option('--notheadless', 'Run in visible browser mode for debugging')
    .option('--dodump', 'Dump HTML content to files for debugging')
    .option('--against <target>', 'Target service: onenote (default) or outlook', 'onenote')
    .action(async (options) => {
        const targetUrl = options.against === 'outlook' ? OUTLOOK_URL : ONENOTE_URL;
        await login({ ...options, targetUrl });
    });

program
    .command('check')
    .description('Check if authenticated')
    .option('--against <target>', 'Target service: onenote (default) or outlook', 'onenote')
    .action(async (options) => {
        const targetUrl = options.against === 'outlook' ? OUTLOOK_URL : ONENOTE_URL;
        const isAuth = await checkAuth(targetUrl);
        if (isAuth) {
            logger.success('Authentication file found. You are authenticated.');
            const meta = await getAuthMeta();
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
    .action(async () => {
        await logout();
        logger.success('Logged out successfully. Authentication state cleared.');
    });

program.parse();
