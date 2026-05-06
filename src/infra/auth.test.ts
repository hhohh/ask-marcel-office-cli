import { afterEach, describe, expect, it } from 'bun:test';
import type { AccessToken } from '../domain/access-token.ts';
import { accessTokenUnsafe } from '../domain/access-token.ts';
import { ok } from '../domain/result.ts';
import { installFetchMock } from '../test-helpers/fetch-mock.ts';
import { createFileSystemFake } from '../test-helpers/filesystem-fake.ts';
import { createLoggerFake } from '../test-helpers/logger-fake.ts';
import { createAuthManager, createAuthManagerFromApi } from './auth.ts';
import type { BrowserAuth, BrowserTokenResult } from './browser-auth.ts';

const CACHE_PATH = '/virtual/token-cache.json';

const fakeBrowserAuth = (config?: { acquireResult?: BrowserTokenResult | null; acquireError?: Error; elevatedResult?: AccessToken | null }): BrowserAuth => ({
  acquireToken: async () => {
    if (config?.acquireError) throw config.acquireError;
    return config?.acquireResult ?? null;
  },
  acquireElevatedToken: async () => config?.elevatedResult ?? null,
  close: async () => {},
});

const futureToken = (): BrowserTokenResult => {
  const future = Math.floor(Date.now() / 1000) + 3600;
  const header = btoa(JSON.stringify({ alg: 'RS256' }));
  const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com', tid: 'tenant-1' }));
  return { accessToken: accessTokenUnsafe(`${header}.${payload}.sig`), refreshToken: 'new-refresh' };
};

const futureElevated = (): AccessToken => {
  const future = Math.floor(Date.now() / 1000) + 3600;
  const header = btoa(JSON.stringify({ alg: 'RS256' }));
  const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com', appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1' }));
  return accessTokenUnsafe(`${header}.${payload}.sig`);
};

describe('auth manager recovery ladder', () => {
  it('returns cached token when fresh and valid', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_on: future, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result).toEqual(ok(accessTokenUnsafe(`${header}.${payload}.sig`)));
    expect(logger.calls.some((l) => l.event === 'auth.ladder.rung' && (l.meta as Record<string, unknown>)?.rung === 'cache')).toBe(true);
  });

  it('refreshes expired token when refresh_token exists', async () => {
    const mock = installFetchMock([
      {
        match: (url, init) => {
          let bodyStr = '';
          if (typeof init?.body === 'string') bodyStr = init.body;
          else if (init?.body instanceof URLSearchParams) bodyStr = init.body.toString();
          return url.includes('/token') && bodyStr.includes('refresh_token=old-refresh');
        },
        respond: () => {
          const future = Math.floor(Date.now() / 1000) + 3600;
          const header = btoa(JSON.stringify({ alg: 'RS256' }));
          const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com' }));
          return new Response(JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_in: 3600, refresh_token: 'new-refresh' }));
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
  });

  it('falls to browser when refresh fails', async () => {
    const mock = installFetchMock([
      {
        match: () => true,
        respond: () => new Response(JSON.stringify({ error: 'invalid_grant' }), { status: 400 }),
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const browserToken = futureToken();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('acquires token via browser when no cache exists', async () => {
    const fs = createFileSystemFake();
    const browserToken = futureToken();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('returns auth_cancelled when browser returns null', async () => {
    const fs = createFileSystemFake();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: null }), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_cancelled');
  });

  it('returns auth_failed when browser throws', async () => {
    const fs = createFileSystemFake();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireError: new Error('browser launch failed') }), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.error.type).toBe('auth_failed');
  });

  it('skips browser when cached token has wrong audience', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'management.core.windows.net' }));
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: `${header}.${payload}.sig`, expires_on: future, refresh_token: '' }));

    const browserToken = futureToken();
    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: browserToken }), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.value).toBe(browserToken.accessToken);
  });

  it('clears cache file on logout', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'token', expires_on: future, refresh_token: 'refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, logger, fs);

    const result = await auth.logout();
    expect(result.ok).toBe(true);
    expect(fs.has(CACHE_PATH)).toBe(false);
  });

  it('still clears cache on logout when browser close fails', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'token', expires_on: future, refresh_token: 'refresh' }));

    const failingBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => null,
      close: async () => {
        throw new Error('close failed');
      },
    };

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(failingBrowser, CACHE_PATH, logger, fs);

    const result = await auth.logout();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('close failed');
    }
    expect(fs.has(CACHE_PATH)).toBe(false);
  });

  it('returns auth_failed when the refresh-token fetch rejects with a network error', async () => {
    const mock = installFetchMock([
      {
        match: () => true,
        respond: () => {
          throw new Error('connection reset');
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: null }), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('connection reset');
    }
  });

  it('exposes an AuthManager port shape from the real createAuthManager production-wiring factory', () => {
    const logger = createLoggerFake();
    const auth = createAuthManager({ cachePath: CACHE_PATH, logger });
    expect(typeof auth.getAccessToken).toBe('function');
    expect(typeof auth.logout).toBe('function');
  });

  it('returns auth_failed with the stringified value when the refresh-token fetch rejects with a non-Error', async () => {
    const mock = installFetchMock([
      {
        match: () => true,
        respond: () => {
          throw 'tcp_reset';
        },
      },
    ]);
    afterEach(() => mock.restore());

    const past = Math.floor(Date.now() / 1000) - 100;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'expired-token', expires_on: past, refresh_token: 'old-refresh' }));

    const logger = createLoggerFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: null }), CACHE_PATH, logger, fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toBe('tcp_reset');
    }
  });
});

describe('auth manager elevated token', () => {
  it('returns the cached elevated token when fresh and valid', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const header = btoa(JSON.stringify({ alg: 'RS256' }));
    const payload = btoa(JSON.stringify({ exp: future, aud: 'https://graph.microsoft.com', appid: 'c0ab8ce9-e9a0-42e7-b064-33d422df41f1' }));
    const elevatedToken = `${header}.${payload}.sig`;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r', elevated_access_token: elevatedToken, elevated_expires_on: future }));

    const auth = createAuthManagerFromApi(fakeBrowserAuth(), CACHE_PATH, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(accessTokenUnsafe(elevatedToken)));
  });

  it('re-captures the elevated token via the browser when the cache is missing it', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: captured }), CACHE_PATH, createLoggerFake(), fs);

    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(captured));
    // Cache should now contain the freshly-captured elevated token.
    const persisted = await fs.readJson<{ elevated_access_token?: string }>(CACHE_PATH);
    expect(persisted.ok && persisted.value.elevated_access_token).toBe(captured);
  });

  it('re-captures when the cached elevated token is expired', async () => {
    const past = Math.floor(Date.now() / 1000) - 3600;
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(
      CACHE_PATH,
      JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r', elevated_access_token: 'expired-elevated', elevated_expires_on: past })
    );

    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: captured }), CACHE_PATH, createLoggerFake(), fs);

    const result = await auth.getElevatedAccessToken();
    expect(result).toEqual(ok(captured));
  });

  it('returns auth_failed with actionable guidance when re-capture returns null', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const auth = createAuthManagerFromApi(fakeBrowserAuth({ elevatedResult: null }), CACHE_PATH, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated token capture failed');
      expect(result.error.message).toContain('ask-marcel logout');
    }
  });

  it('returns auth_failed with the wrapper message when the browser throws during elevated capture', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const throwingBrowser: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => {
        throw new Error('playwright not installed');
      },
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(throwingBrowser, CACHE_PATH, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated capture threw');
      expect(result.error.message).toContain('playwright not installed');
    }
  });

  it('persists the elevated token alongside the Teams token at login time (best-effort)', async () => {
    const fs = createFileSystemFake();
    const captured = futureElevated();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: futureToken(), elevatedResult: captured }), CACHE_PATH, createLoggerFake(), fs);

    await auth.getAccessToken();
    const cacheRead = await fs.readJson<{ elevated_access_token?: string }>(CACHE_PATH);
    expect(cacheRead.ok).toBe(true);
    if (cacheRead.ok) {
      expect(cacheRead.value.elevated_access_token).toBe(captured);
    }
  });

  it('login still succeeds even when elevated capture returns null (the 80+ non-version commands do not need it)', async () => {
    const fs = createFileSystemFake();
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: futureToken(), elevatedResult: null }), CACHE_PATH, createLoggerFake(), fs);

    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
    const cached = await fs.readJson<{ elevated_access_token?: string }>(CACHE_PATH);
    expect(cached.ok && cached.value.elevated_access_token).toBeUndefined();
  });

  it('login still succeeds even when elevated capture throws (best-effort, error is logged not propagated)', async () => {
    const fs = createFileSystemFake();
    const throwingElevated: BrowserAuth = {
      acquireToken: async () => futureToken(),
      acquireElevatedToken: async () => {
        throw new Error('headless edge crashed');
      },
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(throwingElevated, CACHE_PATH, createLoggerFake(), fs);
    const result = await auth.getAccessToken();
    expect(result.ok).toBe(true);
  });

  it('persists elevated_expires_on as 0 when the elevated JWT has no exp claim', async () => {
    const fs = createFileSystemFake();
    const noExpJwt = `${btoa(JSON.stringify({ alg: 'RS256' }))}.${btoa(JSON.stringify({ aud: 'https://graph.microsoft.com', appid: 'c0ab8ce9' }))}.sig`;
    const auth = createAuthManagerFromApi(fakeBrowserAuth({ acquireResult: futureToken(), elevatedResult: accessTokenUnsafe(noExpJwt) }), CACHE_PATH, createLoggerFake(), fs);
    await auth.getAccessToken();
    const cached = await fs.readJson<{ elevated_access_token?: string; elevated_expires_on?: number }>(CACHE_PATH);
    expect(cached.ok && cached.value.elevated_expires_on).toBe(0);
  });

  it('elevated capture throwing a non-Error value is still wrapped (covers the String(e) fallback)', async () => {
    const future = Math.floor(Date.now() / 1000) + 3600;
    const fs = createFileSystemFake();
    fs.seed(CACHE_PATH, JSON.stringify({ access_token: 'teams-tok', expires_on: future, refresh_token: 'r' }));

    const stringThrower: BrowserAuth = {
      acquireToken: async () => null,
      acquireElevatedToken: async () => {
        throw 'edge process killed by SIGKILL';
      },
      close: async () => {},
    };
    const auth = createAuthManagerFromApi(stringThrower, CACHE_PATH, createLoggerFake(), fs);
    const result = await auth.getElevatedAccessToken();
    expect(result.ok).toBe(false);
    if (!result.ok && result.error.type === 'auth_failed') {
      expect(result.error.message).toContain('elevated capture threw');
      expect(result.error.message).toContain('edge process killed');
    }
  });
});
