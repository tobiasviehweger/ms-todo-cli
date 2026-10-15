import { AsyncLocalStorage } from 'async_hooks';
import { AppError, ErrorCodes } from './errors';

const profiles = new AsyncLocalStorage<Readonly<{ profile: string }>>();

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

export function withProfile<T>(profile: string, callback: () => T): T {
  return profiles.run(Object.freeze({ profile: resolveProfile(profile) }), callback);
}
