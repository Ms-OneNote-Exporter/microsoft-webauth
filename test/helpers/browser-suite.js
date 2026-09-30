/**
 * @fileoverview Shared guard for the Playwright-backed test suites.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
const fs = require('fs');
const { chromium } = require('playwright');

/**
 * These suites drive a real browser, which is the expensive part of the setup
 * and the reason the guard exists at all: a contributor who has run
 * `npm install` but not `npx playwright install chromium` should not have a
 * failing test run, so the suites skip rather than fail.
 *
 * Skipping is only safe if it cannot hide a real failure, and that is the part
 * this file exists to get right. A missing browser used to produce a
 * `console.warn` and a skip: 28 of 38 tests vanished and jest still exited 0.
 * On a CI runner with no browser that is a green build that ran none of the
 * auth regressions — the exact thing the workflow added in #11 is meant to
 * prevent, and it was not preventing it.
 *
 * So the behaviour depends on who is asking:
 *
 *   - local, no PW_TESTS_REQUIRE_BROWSER: skip, as before, but say so loudly
 *     enough that it cannot scroll past unnoticed.
 *   - CI, PW_TESTS_REQUIRE_BROWSER=1: fail. The workflow installs Chromium
 *     explicitly, so a skip in CI can only mean the install step broke.
 *
 * @param {string} suiteName  used in both the skip warning and the failure
 * @param {Function} fn       the suite body, normally run against a real browser
 */
function describeWithBrowser(suiteName, fn) {
    const problem = chromiumProblem();
    if (!problem) return describe(suiteName, fn);

    const remedy = 'Run `npx playwright install chromium`.';

    if (process.env.PW_TESTS_REQUIRE_BROWSER === '1') {
        // One failing test rather than a beforeAll on an empty suite: jest
        // rejects an empty suite with "must contain at least one test", which
        // would bury the only message that says what to do about it. The suite
        // body is deliberately not registered, since its tests would each fail
        // with the same text and bury it in noise.
        return describe(suiteName, () => {
            it('requires a chromium install to run', () => {
                throw new Error(
                    `${suiteName} could not run: ${problem}.\n`
                    + `PW_TESTS_REQUIRE_BROWSER=1 is set, so this is a failure and not a skip. ${remedy}`
                );
            });
        });
    }

    console.warn(
        `\n  !! SKIPPING "${suiteName}"\n`
        + `     ${problem}.\n`
        + `     ${remedy}\n`
        + `     Set PW_TESTS_REQUIRE_BROWSER=1 to make this a failure instead.\n`
    );
    return describe.skip(suiteName, fn);
}

/**
 * Why Chromium cannot be launched, or null when it can.
 * @returns {string|null}
 */
function chromiumProblem() {
    let executable = null;
    try {
        executable = chromium.executablePath();
    } catch (_) {
        // Older playwright throws when nothing is downloaded; newer versions
        // return a path regardless. Both mean the same thing here.
        return 'playwright could not resolve a chromium path';
    }
    if (!executable || !fs.existsSync(executable)) {
        return `chromium is not installed (looked in ${executable || 'an unresolvable path'})`;
    }
    return null;
}

module.exports = { chromium, describeWithBrowser, chromiumProblem };
