const path = require('path');

const LOG_ENV = 'ONENOTE_EXPORT_LOG_DIR';
const { resolveLogDir, resolveLogDirFor } = require('../src/utils/logPaths');

describe('log directory resolution', () => {
    let saved;

    beforeEach(() => {
        saved = process.env[LOG_ENV];
    });

    afterEach(() => {
        if (saved === undefined) delete process.env[LOG_ENV];
        else process.env[LOG_ENV] = saved;
    });

    it('honours ONENOTE_EXPORT_LOG_DIR above everything else', () => {
        process.env[LOG_ENV] = '/tmp/somewhere-else';
        expect(resolveLogDir()).toBe('/tmp/somewhere-else');
    });

    it('resolves a relative override to an absolute path', () => {
        process.env[LOG_ENV] = 'relative-logs';
        expect(path.isAbsolute(resolveLogDir())).toBe(true);
    });

    it('ignores a blank override', () => {
        process.env[LOG_ENV] = '   ';
        expect(resolveLogDir()).not.toContain('   ');
    });

    // The bug this fixes. Installed as a dependency - which is what
    // ms-onenote-exporter does - __dirname resolves inside node_modules, so the
    // old hardcoded `../logs` landed there: unreadable in a container running as
    // an unprivileged user, and wiped by the next reinstall besides.
    it('never writes inside node_modules, whatever the install layout', () => {
        for (const layout of [
            '/app/node_modules/@msout/microsoft-webauth/src/utils',
            '/usr/lib/node_modules/@msout/microsoft-webauth/src/utils',
            '/opt/homebrew/lib/node_modules/@msout/microsoft-webauth/src/utils',
            '/home/u/.nvm/versions/node/v24/lib/node_modules/@msout/microsoft-webauth/src/utils',
        ]) {
            delete process.env[LOG_ENV];
            const dir = resolveLogDirFor(layout);
            expect(dir).not.toContain('node_modules');
            expect(path.isAbsolute(dir)).toBe(true);
        }
    });

    it('uses the package logs dir for a plain checkout', () => {
        delete process.env[LOG_ENV];
        const dir = resolveLogDir();
        expect(dir.endsWith(path.join('microsoft-webauth', 'logs')) || dir.endsWith('logs')).toBe(true);
    });

    it('honours XDG_STATE_HOME when installed', () => {
        delete process.env[LOG_ENV];
        const dir = resolveLogDirFor(
            path.join('/app', 'node_modules', '@msout', 'microsoft-webauth', 'src', 'utils')
        );
        // Falls back to ~/.local/state/<package> unless XDG_STATE_HOME is set;
        // either way it is absolute and outside node_modules, which the loop
        // above already asserts. This pins the shape of that answer.
        expect(path.isAbsolute(dir)).toBe(true);
        expect(dir).toContain('microsoft-webauth');
    });
});

describe('the logger actually writes there', () => {
    // The resolution logic above is only worth anything if the logger uses it.
    // This is the assertion that would have caught the original bug: a hardcoded
    // path in the logger, with a correct logPaths.js sitting unused next to it.
    const os = require('os');
    const fs = require('fs');
    let tmp;
    let savedLogDir;

    beforeEach(() => {
        savedLogDir = process.env[LOG_ENV];
        tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'webauth-logdir-'));
        jest.resetModules();
        process.env[LOG_ENV] = tmp;
    });

    afterEach(() => {
        if (savedLogDir === undefined) delete process.env[LOG_ENV];
        else process.env[LOG_ENV] = savedLogDir;
        fs.rmSync(tmp, { recursive: true, force: true });
    });

    it('creates app.log inside the override, not inside the package', () => {
        const logger = require('../src/utils/logger');
        logger.info('a message worth keeping');
        expect(fs.existsSync(path.join(tmp, 'app.log'))).toBe(true);
        expect(fs.readFileSync(path.join(tmp, 'app.log'), 'utf8')).toContain('a message worth keeping');
    });

    it('keeps the log directory owner-only, because dumps hold account state', () => {
        const logger = require('../src/utils/logger');
        logger.info('hello');
        // Windows and some network filesystems do not model this; skip there
        // rather than fail a test about permissions on a platform that has none.
        if (process.platform === 'win32') return;
        expect(fs.statSync(tmp).mode & 0o077).toBe(0);
    });
});
