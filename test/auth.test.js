const path = require('path');
const fs = require('fs-extra');

// Mock the modules before requiring auth
jest.mock('../src/config', () => {
    const actual = jest.requireActual('../src/config');
    return {
        ...actual,
        DEFAULT_AUTH_FILE: '/tmp/test-auth.json'
    };
});

jest.mock('../src/utils/logger', () => ({
    info: jest.fn(),
    warn: jest.fn(),
    error: jest.fn(),
    success: jest.fn(),
    debug: jest.fn(),
    step: jest.fn(),
    getDumpDir: async () => '/tmp/dumps',
    getDumpDisplayPath: () => 'logs/dumps/test'
}));

describe('Auth Module', () => {
    let auth;
    let config;

    beforeEach(() => {
        jest.clearAllMocks();
        // Clean up test files before each test
        fs.removeSync('/tmp/test-auth.json');
        fs.removeSync('/tmp/test-auth-meta.json');
        
        // Re-import after mocks are set up
        config = require('../src/config');
        auth = require('../src/auth');
    });

    afterAll(async () => {
        // Clean up after all tests
        fs.removeSync('/tmp/test-auth.json');
        fs.removeSync('/tmp/test-auth-meta.json');
    });

    describe('getAuthMeta', () => {
        it('should return null when no auth meta file exists', async () => {
            const { getAuthMeta } = require('../src/auth');
            const meta = await getAuthMeta();
            expect(meta).toBeNull();
        });

        it('should return metadata when auth meta file exists', async () => {
            // Create a test meta file
            await fs.writeJson('/tmp/test-auth-meta.json', {
                email: 'test@example.com',
                loginTime: new Date().toISOString()
            });
            
            const { getAuthMeta } = require('../src/auth');
            const meta = await getAuthMeta();
            
            expect(meta).toBeDefined();
            expect(meta.email).toBe('test@example.com');
            expect(meta.loginTime).toBeDefined();
        });
    });

    describe('logout', () => {
        it('should remove auth files', async () => {
            // Create test files
            await fs.writeJson('/tmp/test-auth.json', { test: true });
            await fs.writeJson('/tmp/test-auth-meta.json', { email: 'test@example.com' });
            
            const { logout } = require('../src/auth');
            await logout();
            
            expect(await fs.pathExists('/tmp/test-auth.json')).toBe(false);
            expect(await fs.pathExists('/tmp/test-auth-meta.json')).toBe(false);
        });
    });

    describe('checkAuth', () => {
        it('should return false when no auth file exists', async () => {
            const { checkAuth } = require('../src/auth');
            const isAuth = await checkAuth();
            expect(isAuth).toBe(false);
        });

        it('should handle network errors gracefully', async () => {
            // Create a dummy auth file
            await fs.writeJson('/tmp/test-auth.json', { test: true });
            
            const { checkAuth } = require('../src/auth');
            // Use mock URL that won't actually connect
            const isAuth = await checkAuth('https://mock.test/nonexistent');
            
            // Should return true on error (conservative approach)
            expect(isAuth).toBe(true);
        }, 10000);
    });

    describe('getAuthenticatedContext', () => {
        it('should throw error when no auth file exists', async () => {
            const { getAuthenticatedContext } = require('../src/auth');
            
            await expect(getAuthenticatedContext(null)).rejects.toThrow(
                'No authentication state found. Please run "login" command first.'
            );
        });
    });
});

describe('config', () => {
    let config;
    
    beforeEach(() => {
        jest.clearAllMocks();
        fs.removeSync('/tmp/test-auth.json');
        fs.removeSync('/tmp/test-auth-meta.json');
        config = require('../src/config');
    });

    it('should export DEFAULT_AUTH_FILE', () => {
        expect(config.DEFAULT_AUTH_FILE).toBe('/tmp/test-auth.json');
    });

    it('should export ONENOTE_URL', () => {
        expect(config.ONENOTE_URL).toBe('https://onenote.cloud.microsoft/en-us');
    });

    it('should export OUTLOOK_URL', () => {
        expect(config.OUTLOOK_URL).toBe('https://outlook.cloud.microsoft/mail/');
    });

    it('should export getAuthMetaFilePath', () => {
        expect(config.getAuthMetaFilePath('/tmp/test-auth.json')).toBe('/tmp/test-auth-meta.json');
    });
});
