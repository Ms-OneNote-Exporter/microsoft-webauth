/**
 * @fileoverview Jest configuration.
 * @author phptr,enoola,msout
 * @copyright 2026 phptr,enoola,msout
 */
module.exports = {
    testEnvironment: 'node',
    testPathIgnorePatterns: ['/node_modules/', '/dist/'],
    collectCoverageFrom: ['src/**/*.js'],
    coverageDirectory: 'coverage',
    verbose: true
};
