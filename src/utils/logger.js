const chalk = require('chalk');
const fs = require('fs-extra');
const path = require('path');
const { resolveLogDir } = require('./logPaths');

class Logger {
    constructor() {
        this.months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        this.logDir = resolveLogDir();
        this.logFilePath = path.join(this.logDir, 'app.log');

         // Initialize dump directory name once per execution
        const now = new Date();
        const yyyy = now.getFullYear();
        const mm = String(now.getMonth() + 1).padStart(2, '0');
        const dd = String(now.getDate()).padStart(2, '0');
        const hh = String(now.getHours()).padStart(2, '0');
        const min = String(now.getMinutes()).padStart(2, '0');

         // Format: YYYY-MM-DD_HHhMM
        this.dumpSubDir = `${yyyy}-${mm}-${dd}_${hh}h${min}`;

         // Ensure logs directory exists. Restricted permissions because the dump
        // files written next to app.log contain the authenticated DOM of a real
        // Microsoft account: cookies, tenant hostnames and note titles. The dump
        // path redacts known credential fields, but the surrounding page does
        // not, so the directory is owner-only rather than merely gitignored.
        this._ensurePrivateDir(this.logDir);
        this._tightenExistingLogFile();
    }

    /**
     * Brings an existing app.log down to owner-only.
     *
     * Files this process creates are 0600 from the start, but a log written by
     * an earlier version - which had no `mode` at all - is still 0644, and
     * appendFileSync's `mode` only applies at creation. One chmod at startup is
     * enough to close that off.
     */
    _tightenExistingLogFile() {
        try {
            const stats = fs.statSync(this.logFilePath);
            if ((stats.mode & 0o077) !== 0) {
                fs.chmodSync(this.logFilePath, 0o600);
            }
        } catch (e) {
            // No log file yet, or chmod unsupported: not fatal.
        }
    }

    /**
     * Creates `dir` and forces it to owner-only.
     *
     * fs.ensureDirSync honours the process umask, so on a permissive umask the
     * dumps would end up world-readable. chmod makes it explicit rather than
     * dependent on the environment.
     *
     * @param {string} dir - Directory to create
     */
    _ensurePrivateDir(dir) {
        fs.ensureDirSync(dir);
        try {
            fs.chmodSync(dir, 0o700);
        } catch (e) {
            // A filesystem that does not support chmod is not a reason to fail
            // a login; the log is a diagnostic aid, not the product.
        }
    }

    _getTimestamp() {
        const now = new Date();
        const month = this.months[now.getMonth()];
        const day = String(now.getDate()).padStart(2, '0');
        const time = now.toTimeString().split(' ')[0];
        return `[${month} ${day} ${time}]`;
    }

    /**
     * Returns the absolute path to the current session's dump directory.
     * Ensures the directory exists, owner-only.
     * @returns {Promise<string>}
     */
    async getDumpDir() {
        const dumpDir = path.join(this.logDir, 'dumps', this.dumpSubDir);
        fs.ensureDirSync(dumpDir);
        try {
            fs.chmodSync(dumpDir, 0o700);
        } catch (e) {
            // See _ensurePrivateDir: not fatal.
        }
        return dumpDir;
    }

    /**
     * Returns a user-friendly relative path for logging.
     *
     * Relative to the cwd rather than a hardcoded `logs/dumps/...`, because the
     * log directory is no longer always <package>/logs: an installed copy writes
     * to the XDG state dir, and ms-onenote-exporter points every step at one
     * shared directory. A fixed string would name a path that does not exist.
     * @returns {string}
     */
    getDumpDisplayPath() {
        return path.relative(process.cwd(), path.join(this.logDir, 'dumps', this.dumpSubDir)) || '.';
    }

    _stripColors(str) {
         // eslint-disable-next-line no-control-regex
        return str.replace(/\u001b\[[0-9;]*m/g, '');
    }

    _formatMessage(level, message, colorFunc = (m) => m) {
        const timestamp = this._getTimestamp();
        const coloredTimestamp = chalk.gray(timestamp);
        const levelTag = `[${level}]`;
        const coloredLevelTag = colorFunc(levelTag);

         // Handle multi-line messages
        let formattedMessage = '';
        if (typeof message === 'string' && message.includes('\n')) {
            formattedMessage = message.split('\n').map(line => `${coloredTimestamp} ${coloredLevelTag} ${line}`).join('\n');
        } else if (typeof message !== 'string') {
             // Handle objects/errors
            try {
                const stringified = JSON.stringify(message, null, 2);
                formattedMessage = `${coloredTimestamp} ${coloredLevelTag} ${stringified}`;
            } catch (e) {
                formattedMessage = `${coloredTimestamp} ${coloredLevelTag} [Complex Object]`;
            }
        } else {
            formattedMessage = `${coloredTimestamp} ${coloredLevelTag} ${message}`;
        }

         // Write to log file (no colors)
        const plainTimestamp = timestamp;
        const plainLevelTag = levelTag;
        let plainMessage = '';

        if (typeof message === 'string' && message.includes('\n')) {
            plainMessage = message.split('\n').map(line => `${plainTimestamp} ${plainLevelTag} ${line}`).join('\n');
        } else if (typeof message !== 'string') {
            try {
                const stringified = JSON.stringify(message, null, 2);
                plainMessage = `${plainTimestamp} ${plainLevelTag} ${stringified}`;
            } catch (e) {
                plainMessage = `${plainTimestamp} ${plainLevelTag} [Complex Object]`;
            }
        } else {
            plainMessage = `${plainTimestamp} ${plainLevelTag} ${message}`;
        }

         // Append to log file. `mode` only applies at creation, so an app.log
        // written by an earlier version stays 0644; _tightenExistingLogFile
        // brings that down to owner-only on the next run.
        fs.appendFileSync(this.logFilePath, plainMessage + '\n', { mode: 0o600 });

        return formattedMessage;
    }

    /** Generic log method for programmatic use */
    log(level, message) {
        const lv = (level || 'info').toLowerCase();
        if (this[lv] && typeof this[lv] === 'function') {
            this[lv](message);
        } else {
            this.info(message);
        }
    }

    info(message) {
        process.stdout.write(this._formatMessage('INFO', message, chalk.blue) + '\n');
    }

    warn(message) {
        process.stdout.write(this._formatMessage('WARN', message, chalk.yellow) + '\n');
    }

    error(message, error = null) {
        process.stderr.write(this._formatMessage('ERROR', message, chalk.red) + '\n');
        if (error) {
            if (error.stack) {
                const stack = chalk.red(error.stack);
                process.stderr.write(stack + '\n');
                 // Also write stack to file
                fs.appendFileSync(this.logFilePath, this._stripColors(stack) + '\n');
            } else {
                process.stderr.write(this._formatMessage('ERROR', error, chalk.red) + '\n');
            }
        }
    }

    success(message) {
        process.stdout.write(this._formatMessage('SUCCESS', message, chalk.green) + '\n');
    }

    debug(message) {
        process.stdout.write(this._formatMessage('DEBUG', message, chalk.gray) + '\n');
    }

    step(message) {
        process.stdout.write(this._formatMessage('STEP', message, chalk.magenta) + '\n');
    }
}

module.exports = new Logger();
