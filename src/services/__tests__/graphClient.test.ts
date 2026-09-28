// services/__tests__/graphClient.test.ts — Where the bearer token may go.
//
// The token is a delegated credential for a real tenant. These tests pin the
// one rule that protects it in transit: it is attached to requests for Graph's
// own host and nothing else, and a refused URL is refused BEFORE the request,
// not after.

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { graphFetch, graphPost, fetchAllPages } from '../graphClient';

const TOKEN = 'test-token';

function json(body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { 'Content-Type': 'application/json' },
  });
}

// The network boundary is mocked; our own logic never is.
const mockFetch = vi.fn((_url: string, _init?: RequestInit) => Promise.resolve(json({ value: [] })));

/** Every URL a request was actually made to. */
function requested(): string[] {
  return mockFetch.mock.calls.map((call) => call[0]);
}

beforeEach(() => {
  mockFetch.mockClear();
  vi.stubGlobal('fetch', mockFetch);
});
afterEach(() => vi.unstubAllGlobals());

describe('graphFetch — destination of the bearer token', () => {
  it('resolves a relative endpoint against Graph v1.0', async () => {
    await graphFetch('/organization?$select=displayName', TOKEN);
    expect(requested()).toEqual(['https://graph.microsoft.com/v1.0/organization?$select=displayName']);
    expect(mockFetch.mock.calls[0][1]?.headers).toMatchObject({ Authorization: `Bearer ${TOKEN}` });
  });

  it('accepts an absolute URL on the Graph host — the beta endpoint needs one', async () => {
    const beta = 'https://graph.microsoft.com/beta/identity/conditionalAccess/policies';
    await graphFetch(beta, TOKEN);
    expect(requested()).toEqual([beta]);
  });

  it.each([
    ['another host', 'https://example.net/v1.0/users'],
    ['a host that shares the prefix', 'https://graph.microsoft.com.example.net/v1.0/users'],
    ['Graph as the userinfo of another host', 'https://graph.microsoft.com@example.net/v1.0/users'],
  ])('refuses %s without making a request', async (_label, url) => {
    await expect(graphFetch(url, TOKEN)).rejects.toThrow('Untrusted Graph URL');
    expect(mockFetch).not.toHaveBeenCalled();
  });

  it('never treats a non-https absolute URL as a destination', async () => {
    // Anything that is not https:// is a path on Graph, so the worst outcome
    // is a 404 from Graph — the token still goes nowhere else.
    await graphFetch('http://example.net/users', TOKEN);
    expect(requested()[0].startsWith('https://graph.microsoft.com/v1.0')).toBe(true);
  });
});

describe('graphPost — destination of the bearer token', () => {
  it('posts to Graph v1.0', async () => {
    await graphPost('/directoryObjects/getByIds', TOKEN, { ids: [] });
    expect(requested()).toEqual(['https://graph.microsoft.com/v1.0/directoryObjects/getByIds']);
  });

  it('refuses another host without making a request', async () => {
    await expect(graphPost('https://example.net/collect', TOKEN, {})).rejects.toThrow('Untrusted Graph URL');
    expect(mockFetch).not.toHaveBeenCalled();
  });
});

describe('fetchAllPages — following @odata.nextLink', () => {
  it('follows a next link on the Graph host and joins the pages', async () => {
    const next = 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc';
    mockFetch
      .mockResolvedValueOnce(json({ value: [{ id: '1' }], '@odata.nextLink': next }))
      .mockResolvedValueOnce(json({ value: [{ id: '2' }] }));

    const rows = await fetchAllPages<{ id: string }>('/users', TOKEN);

    expect(rows).toEqual([{ id: '1' }, { id: '2' }]);
    expect(requested()).toEqual(['https://graph.microsoft.com/v1.0/users', next]);
  });

  it.each([
    ['another host', 'https://example.net/v1.0/users?$skiptoken=abc'],
    ['a host that shares the prefix', 'https://graph.microsoft.com.example.net/users'],
    ['a relative link', '/users?$skiptoken=abc'],
  ])('refuses a next link to %s and stops after the first page', async (_label, next) => {
    mockFetch.mockResolvedValueOnce(json({ value: [{ id: '1' }], '@odata.nextLink': next }));

    await expect(fetchAllPages('/users', TOKEN)).rejects.toThrow('Untrusted @odata.nextLink URL');
    expect(mockFetch).toHaveBeenCalledOnce();
  });

  it('sends extra headers on every page', async () => {
    const next = 'https://graph.microsoft.com/v1.0/servicePrincipals?$skiptoken=abc';
    mockFetch
      .mockResolvedValueOnce(json({ value: [], '@odata.nextLink': next }))
      .mockResolvedValueOnce(json({ value: [] }));

    await fetchAllPages('/servicePrincipals', TOKEN, { ConsistencyLevel: 'eventual' });

    for (const call of mockFetch.mock.calls) {
      expect(call[1]?.headers).toMatchObject({ ConsistencyLevel: 'eventual' });
    }
  });
});
