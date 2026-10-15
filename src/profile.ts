import { AsyncLocalStorage } from 'async_hooks';
import { AppError, ErrorCodes } from './errors';

export interface ProfileAccountIdentity {
  homeAccountId: string;
  environment: string;
  tenantId: string;
}

interface ProfileContext {
  readonly profile: string;
  readonly auth: { accountIdentity?: Readonly<ProfileAccountIdentity> };
}

const profiles = new AsyncLocalStorage<ProfileContext>();

export function resolveProfile(explicit?: string): string {
  const profile = explicit ?? process.env.MS_TODO_PROFILE ?? 'default';
  if (!/^[a-z0-9][a-z0-9_-]{0,63}$/.test(profile)) {
    throw new AppError(
      ErrorCodes.VALIDATION_ERROR,
      'Profile must be 1–64 lowercase letters, digits, hyphens or underscores, starting with a letter or digit.',
    );
  }
  return profile;
}

export function getProfile(): string {
  return profiles.getStore()?.profile ?? 'default';
}

export function getPinnedAccountIdentity(): Readonly<ProfileAccountIdentity> | undefined {
  return profiles.getStore()?.auth.accountIdentity;
}

export function pinAccountIdentity(identity: ProfileAccountIdentity): void {
  const context = profiles.getStore();
  if (context && !context.auth.accountIdentity) {
    context.auth.accountIdentity = Object.freeze({ ...identity });
  }
}

export function withProfile<T>(profile: string, callback: () => T): T {
  const resolved = resolveProfile(profile);
  const current = profiles.getStore();
  const context = current?.profile === resolved
    ? current
    : Object.freeze({ profile: resolved, auth: {} });
  return profiles.run(context, callback);
}
