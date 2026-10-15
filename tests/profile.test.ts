import { getProfile, resolveProfile, withProfile } from '../src/profile';

const originalProfile = process.env.MS_TODO_PROFILE;

afterEach(() => {
  if (originalProfile === undefined) delete process.env.MS_TODO_PROFILE;
  else process.env.MS_TODO_PROFILE = originalProfile;
});

test('selection precedence is flag, environment, then default', () => {
  delete process.env.MS_TODO_PROFILE;
  expect(resolveProfile()).toBe('default');
  process.env.MS_TODO_PROFILE = 'personal';
  expect(resolveProfile()).toBe('personal');
  expect(resolveProfile('work')).toBe('work');
});

test.each(['', '.', '..', '../work', '/work', 'Work', 'a/b', 'a\\b', 'a b', 'a'.repeat(65)])(
  'rejects unsafe or ambiguous alias %j',
  (alias) => expect(() => resolveProfile(alias)).toThrow('Profile must be'),
);

test('does not silently replace an invalid environment profile with default', () => {
  process.env.MS_TODO_PROFILE = '../work';
  expect(() => resolveProfile()).toThrow();
  expect(resolveProfile('work')).toBe('work');
});

test('concurrent asynchronous operations retain independent immutable profiles', async () => {
  const results = await Promise.all(['work', 'personal'].map((profile) =>
    withProfile(profile, async () => {
      await new Promise((resolve) => setTimeout(resolve, 5));
      expect(getProfile()).toBe(profile);
      return withProfile('nested', async () => {
        await Promise.resolve();
        expect(getProfile()).toBe('nested');
        return profile;
      });
    }),
  ));
  expect(results).toEqual(['work', 'personal']);
  expect(getProfile()).toBe('default');
});
