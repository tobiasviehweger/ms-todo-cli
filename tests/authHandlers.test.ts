import * as manager from '../src/auth/authManager';
import * as output from '../src/output';
import { withProfile } from '../src/profile';
import { handleAuthLogin, handleAuthLogout, handleAuthStatus, handleAuthPrintAccount, handleAuthProfiles } from '../src/auth';
import { AppError, ErrorCodes } from '../src/errors';

jest.mock('../src/auth/authManager');

beforeEach(() => {
  jest.resetAllMocks();
  jest.spyOn(output, 'printSuccess').mockImplementation(() => undefined);
  jest.spyOn(output, 'printError').mockImplementation(() => undefined);
});

afterEach(() => jest.restoreAllMocks());

test('login forwards replacement and includes the selected profile', async () => {
  jest.mocked(manager.getAccount).mockResolvedValue(null);
  await withProfile('work', () => handleAuthLogin({ replace: true }));
  expect(manager.login).toHaveBeenCalledWith({ replace: true });
  expect(output.printSuccess).toHaveBeenCalledWith(expect.objectContaining({ profile: 'work' }));
});

test('status reports cached identity explicitly without validating a token', async () => {
  jest.mocked(manager.getAccount).mockResolvedValue({
    homeAccountId: 'id', environment: 'cloud', tenantId: 'tenant', username: 'user', name: 'User',
  });
  await withProfile('work', () => handleAuthStatus());
  expect(output.printSuccess).toHaveBeenCalledWith(expect.objectContaining({
    profile: 'work', authenticated: true, tokenVerified: false, status: 'cached',
  }));
  expect(manager.getAccessToken).not.toHaveBeenCalled();
  jest.mocked(manager.getAccount).mockResolvedValue(null);
  await handleAuthStatus();
  expect(output.printSuccess).toHaveBeenLastCalledWith({
    profile: 'default', authenticated: false, tokenVerified: false, status: 'unauthenticated',
  });
});

test('logout, account hints and profile listing identify the selected profile', async () => {
  await withProfile('work', () => handleAuthLogout());
  expect(output.printSuccess).toHaveBeenCalledWith(expect.objectContaining({ profile: 'work' }));
  jest.mocked(manager.getAccount).mockResolvedValue(null);
  await withProfile('work', () => handleAuthPrintAccount());
  expect(output.printError).toHaveBeenCalledWith(ErrorCodes.AUTH_REQUIRED, expect.stringContaining('--profile work auth login'));
  jest.mocked(manager.listProfiles).mockResolvedValue([{ profile: 'work', account: null }]);
  await withProfile('work', () => handleAuthProfiles());
  expect(output.printSuccess).toHaveBeenLastCalledWith({
    profile: 'work', profiles: [{ profile: 'work', account: null }],
  });
});

test('profile-aware errors are not prefixed twice', async () => {
  const message = 'Profile "work": token refresh failed';
  jest.mocked(manager.logout).mockRejectedValue(new AppError(ErrorCodes.AUTH_EXPIRED, message));
  await withProfile('work', () => handleAuthLogout());
  expect(output.printError).toHaveBeenCalledWith(ErrorCodes.AUTH_EXPIRED, message);
});

test.each([handleAuthLogin, handleAuthLogout, handleAuthStatus, handleAuthPrintAccount, handleAuthProfiles])(
  'handler errors include profile and preserve structured error codes',
  async handler => {
    const error = new AppError(ErrorCodes.AUTH_EXPIRED, 'cached token expired');
    jest.mocked(manager.login).mockRejectedValue(error);
    jest.mocked(manager.logout).mockRejectedValue(error);
    jest.mocked(manager.getAccount).mockRejectedValue(error);
    jest.mocked(manager.listProfiles).mockRejectedValue(error);
    await withProfile('work', () => handler());
    expect(output.printError).toHaveBeenCalledWith(ErrorCodes.AUTH_EXPIRED, expect.stringContaining('Profile "work"'));
  },
);
