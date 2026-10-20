import { describe, it, expect } from 'vitest';
import { isGuid, redactIds } from '../guid';

describe('isGuid', () => {
  it('accepts canonical GUIDs in either case', () => {
    expect(isGuid('00000003-0000-0000-c000-000000000000')).toBe(true);
    expect(isGuid('2793995E-0A7D-40D7-BD35-6968BA142197')).toBe(true);
  });

  it('rejects braces, whitespace, special values and near-misses', () => {
    expect(isGuid('{00000003-0000-0000-c000-000000000000}')).toBe(false);
    expect(isGuid(' 00000003-0000-0000-c000-000000000000')).toBe(false);
    expect(isGuid('All')).toBe(false);
    expect(isGuid('00000003-0000-0000-c000-00000000000')).toBe(false);
    expect(isGuid('')).toBe(false);
  });
});

describe('redactIds', () => {
  it('replaces every GUID in a line, in either case', () => {
    const line =
      'Endpoint resolution failed for https://login.microsoftonline.com/AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE/v2.0 ' +
      '(correlation 00000003-0000-0000-c000-000000000000)';
    expect(redactIds(line)).toBe(
      'Endpoint resolution failed for https://login.microsoftonline.com/<id>/v2.0 (correlation <id>)',
    );
  });

  it('leaves a line without a GUID unchanged', () => {
    const line = 'Authority https://login.microsoftonline.com/organizations did not resolve';
    expect(redactIds(line)).toBe(line);
  });
});
