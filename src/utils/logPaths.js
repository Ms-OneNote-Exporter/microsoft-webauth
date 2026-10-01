const os = require('os');
const path = require('path');

/**
 * Decides the log directory for a given package location.
 *
 * Split out from resolveLogDir so the install-layout decision can be tested
 * against paths that do not exist on the test machine - a global install cannot
 * be reproduced locally.
 *
 * NOTE: this file is deliberately identical to the one in
 * microsoft-onenote-list-notebooks and microsoft-onenote-export-notebook, which
 * are separate packages that cannot share source. Only PACKAGE_NAME differs.
 * If you change it here, change it there too, or the next release of this
 * package will disagree with its siblings about where an install writes logs.
 *
 * @param {string} packageRoot - Absolute path of the package root
 * @returns {string} Absolute path to the directory holding app.log and dumps/
 */
function resolveLogDirFor(packageRoot) {
    const isGlobalInstall = packageRoot.split(path.sep).includes('node_modules');

    if (!isGlobalInstall) {
        return path.join(packageRoot, 'logs');
    }

    const xdgState = process.env.XDG_STATE_HOME;
    if (xdgState && xdgState.trim()) {
        return path.join(path.resolve(xdgState.trim()), 'microsoft-webauth');
    }

    const home = os.homedir();
    if (home) {
        return path.join(home, '.local', 'state', 'microsoft-webauth');
    }

    return path.join(os.tmpdir(), 'microsoft-webauth');
}

/**
 * Resolves where logs and HTML dumps are written.
 *
 * The previous code hardcoded `path.resolve(__dirname, '../logs')`, which is
 * correct for a checkout but wrong the moment this package is a dependency:
 *
 *   npm install @msout/microsoft-webauth
 *
 * puts it inside a node_modules tree, so the logger wrote to
 * `node_modules/@msout/microsoft-webauth/logs/app.log` - inside node_modules,
 * where it is liable to be read-only and wiped by the next reinstall, and where
 * nobody looks for it. In a container running as an unprivileged user it is
 * worse than misplaced: the directory cannot be created at all, and the logger
 * throws before the login starts.
 *
 * ONENOTE_EXPORT_LOG_DIR exists for that case: ms-onenote-exporter installs
 * this package as a dependency and points all three of its steps at one log
 * directory, so a single pipeline run produces a single app.log.
 *
 * Precedence:
 *   1. ONENOTE_EXPORT_LOG_DIR  - explicit override, wins over everything
 *   2. a local checkout        - <package>/logs, which is gitignored and where a
 *                                developer expects to find it
 *   3. an installed dependency - XDG state dir, else ~/.local/state/<package>
 *   4. os.tmpdir()             - last resort, so logging never breaks a login
 *
 * @returns {string} Absolute path to the directory holding app.log and dumps/
 */
function resolveLogDir() {
    const override = process.env.ONENOTE_EXPORT_LOG_DIR;
    if (override && override.trim()) {
        return path.resolve(override.trim());
    }

    return resolveLogDirFor(path.resolve(__dirname, '..', '..'));
}

module.exports = { resolveLogDir, resolveLogDirFor };
