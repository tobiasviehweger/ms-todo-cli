import { login, getAccount, logout, listProfiles } from './authManager';
import { printSuccess, printError } from '../output';
import { ErrorCodes } from '../errors';
import { getProfile } from '../profile';

function reportError(err: unknown): void {
  const e = err as { code?: string; message?: string };
  const prefix = `Profile "${getProfile()}": `;
  const message = e.message || 'Unknown error';
  printError(e.code || ErrorCodes.GRAPH_ERROR, message.startsWith(prefix) ? message : prefix + message);
}

export async function handleAuthLogin(options?: { replace?: boolean }): Promise<void> {
  try {
    await login(options);
    const account = await getAccount();
    printSuccess({ profile: getProfile(), message: 'Logged in successfully', account });
  } catch (err: unknown) {
    reportError(err);
  }
}

export async function handleAuthStatus(): Promise<void> {
  try {
    const account = await getAccount();
    printSuccess({ profile: getProfile(), authenticated: account !== null, tokenVerified: false, status: account ? 'cached' : 'unauthenticated', ...(account ? { account } : {}) });
  } catch (err: unknown) {
    reportError(err);
  }
}

export async function handleAuthLogout(): Promise<void> {
  try {
    await logout();
    printSuccess({ profile: getProfile(), message: 'Logged out successfully' });
  } catch (err: unknown) {
    reportError(err);
  }
}

export async function handleAuthPrintAccount(): Promise<void> {
  try {
    const account = await getAccount();
    if (!account) {
      printError(ErrorCodes.AUTH_REQUIRED, `Profile "${getProfile()}": not authenticated. Run: ms-todo-cli --profile ${getProfile()} auth login`);
      return;
    }
    printSuccess({ profile: getProfile(), account });
  } catch (err: unknown) {
    reportError(err);
  }
}

export async function handleAuthProfiles(): Promise<void> {
  try {
    printSuccess({ profile: getProfile(), profiles: await listProfiles() });
  } catch (err: unknown) {
    reportError(err);
  }
}
