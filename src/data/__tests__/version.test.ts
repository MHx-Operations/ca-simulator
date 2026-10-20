// version.test.ts — APP_VERSION and package.json must move together.
//
// They are two halves of one release step, and nothing else reads both. The
// public mirror's commit is named from package.json, so the two have to agree.
//
// package.json is imported rather than read from disk: resolveJsonModule is
// already on, and this project deliberately carries no @types/node, so the
// node:fs route would not typecheck.

import { describe, it, expect } from 'vitest';
import pkg from '../../../package.json';
import { APP_VERSION } from '../theme';

describe('release version', () => {
  it('APP_VERSION carries a semver and the beta marker', () => {
    expect(APP_VERSION).toMatch(/^v\d+\.\d+\.\d+ beta$/);
  });

  it('matches package.json, which names the public mirror commit', () => {
    expect(`v${pkg.version} beta`).toBe(APP_VERSION);
  });
});
