import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { AuthManager } from '../infra/auth.ts';

type GraphError =
  | { type: 'api_error'; status: number; message: string }
  | { type: 'auth_failed'; message: string }
  | { type: 'network_error'; message: string }
  | { type: 'validation_error'; message: string };

type GraphClient = {
  get: (path: string) => Promise<Result<unknown, GraphError>>;
  post: (path: string, body: unknown) => Promise<Result<unknown, GraphError>>;
  getBinary: (path: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Auth-less fetch of an arbitrary URL whose host MUST be on the
   * Microsoft allow-list. Used to follow `@microsoft.graph.downloadUrl`
   * 302 redirects (CDN-signed URLs) that the format-conversion
   * commands sometimes get back from Graph instead of inline bytes.
   */
  fetchUrl: (url: string) => Promise<Result<unknown, GraphError>>;
  /**
   * Upload bytes to a drive item. `basePath` is the bare driveItem
   * path (e.g. `/me/drive/root:/.ask-marcel-temp/abc.rtf`) — `put()`
   * appends `:/content` for the simple ≤4 MiB sync path or
   * `:/createUploadSession` for the chunked-session path internally
   * based on `body.byteLength`. No upper file-size limit beyond the
   * user's OneDrive quota.
   */
  put: (basePath: string, body: Uint8Array, contentType?: string) => Promise<Result<unknown, GraphError>>;
  delete: (path: string) => Promise<Result<unknown, GraphError>>;
};

const ALLOWED_FETCH_URL_HOSTS: ReadonlyArray<RegExp> = [
  /\.sharepoint\.com$/i,
  /\.onedrive\.com$/i,
  /\.live\.com$/i,
  /\.officeapps\.live\.com$/i,
  /\.1drv\.com$/i,
  /^graph\.microsoft\.com$/i,
  /\.svc\.ms$/i,
];

const isAllowedFetchUrlHost = (host: string): boolean => ALLOWED_FETCH_URL_HOSTS.some((re) => re.test(host));

type FetchFn = (url: string, init?: RequestInit) => Promise<Response>;

const REQUEST_TIMEOUT_MS = 60_000;
const SIMPLE_PUT_THRESHOLD = 4 * 1024 * 1024; // 4 MiB
const CHUNK_SIZE = 5 * 1024 * 1024; // 5 MiB — Graph requires multiples of 320 KiB; 5 MiB is 16 × 320 KiB

const isJson = (contentType: string | null): boolean => contentType !== null && contentType.toLowerCase().includes('application/json');

const isText = (contentType: string | null): boolean => {
  if (contentType === null) return false;
  const lower = contentType.toLowerCase();
  return lower.startsWith('text/') || lower.includes('+xml') || lower.includes('application/xml') || lower.includes('application/javascript');
};

const toBase64 = (bytes: Uint8Array): string => {
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
};

const networkErrorMessage = (e: unknown): string => {
  if (e instanceof Error && e.name === 'TimeoutError') return 'request timed out after 60s';
  if (e instanceof Error && e.name === 'AbortError') return 'request aborted';
  if (e instanceof Error) return e.message;
  if (typeof e === 'string') return e;
  return 'network request failed';
};

const emptyOnJsonFailure = (): { error?: { message?: string } } => ({});

const apiErrorFrom = async (res: Response): Promise<GraphError> => {
  const errBody = (await res.json().catch(emptyOnJsonFailure)) as { error?: { message?: string } };
  return { type: 'api_error', status: res.status, message: errBody.error?.message ?? res.statusText };
};

const createGraphClient = (auth: AuthManager, fetchFn: FetchFn = globalThis.fetch): GraphClient => {
  const authHeaders = async (): Promise<Result<{ Authorization: string }, GraphError>> => {
    const tokenResult = await auth.getAccessToken();
    if (!tokenResult.ok) {
      const msg = tokenResult.error.type === 'auth_cancelled' ? 'Auth cancelled' : tokenResult.error.message;
      return err({ type: 'auth_failed', message: msg });
    }
    return ok({ Authorization: `Bearer ${tokenResult.value}` });
  };

  const request = async (method: 'GET' | 'POST', path: string, body?: unknown): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;

    try {
      const res = await fetchFn(`https://graph.microsoft.com/v1.0${path}`, {
        method,
        headers: { ...headers.value, 'content-type': 'application/json' },
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
        ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      });
      if (!res.ok) return err(await apiErrorFrom(res));
      return ok(await res.json());
    } catch (e: unknown) {
      return err({ type: 'network_error', message: networkErrorMessage(e) });
    }
  };

  const getBinary = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;

    try {
      const res = await fetchFn(`https://graph.microsoft.com/v1.0${path}`, {
        method: 'GET',
        headers: headers.value,
        redirect: 'manual',
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (res.status >= 300 && res.status < 400) {
        const location = res.headers.get('location');
        if (location !== null) return ok({ '@microsoft.graph.downloadUrl': location });
      }
      if (!res.ok) return err(await apiErrorFrom(res));
      const contentType = res.headers.get('content-type');
      if (isJson(contentType)) return ok(await res.json());
      if (isText(contentType)) {
        const text = await res.text();
        return ok({ contentType: contentType ?? 'text/plain', size: text.length, text });
      }
      const buffer = await res.arrayBuffer();
      return ok({ contentType: contentType ?? 'application/octet-stream', size: buffer.byteLength, base64: toBase64(new Uint8Array(buffer)) });
    } catch (e: unknown) {
      return err({ type: 'network_error', message: networkErrorMessage(e) });
    }
  };

  const fetchUrl = async (url: string): Promise<Result<unknown, GraphError>> => {
    let host: string;
    try {
      host = new URL(url).host;
    } catch {
      return err({ type: 'network_error', message: `fetchUrl rejected: invalid URL ${url}` });
    }
    if (!isAllowedFetchUrlHost(host)) {
      return err({ type: 'network_error', message: `fetchUrl rejected: host ${host} not in Microsoft allow-list` });
    }

    try {
      const res = await fetchFn(url, {
        method: 'GET',
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) {
        return err({ type: 'api_error', status: res.status, message: res.statusText });
      }
      const contentType = res.headers.get('content-type');
      if (isJson(contentType)) return ok(await res.json());
      if (isText(contentType)) {
        const text = await res.text();
        return ok({ contentType: contentType ?? 'text/plain', size: text.length, text });
      }
      const buffer = await res.arrayBuffer();
      return ok({ contentType: contentType ?? 'application/octet-stream', size: buffer.byteLength, base64: toBase64(new Uint8Array(buffer)) });
    } catch (e: unknown) {
      return err({ type: 'network_error', message: networkErrorMessage(e) });
    }
  };

  const simplePut = async (path: string, body: Uint8Array, contentType?: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;
    try {
      const res = await fetchFn(`https://graph.microsoft.com/v1.0${path}`, {
        method: 'PUT',
        headers: { ...headers.value, 'content-type': contentType ?? 'application/octet-stream' },
        body: body as unknown as BodyInit,
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res));
      return ok(await res.json());
    } catch (e: unknown) {
      return err({ type: 'network_error', message: networkErrorMessage(e) });
    }
  };

  const chunkedPut = async (basePath: string, body: Uint8Array): Promise<Result<unknown, GraphError>> => {
    // 1. Create upload session via the authenticated request() helper.
    const session = await request('POST', `${basePath}:/createUploadSession`, {
      item: { '@microsoft.graph.conflictBehavior': 'replace' },
    });
    if (!session.ok) return session;
    const uploadUrl = (session.value as { uploadUrl?: string }).uploadUrl;
    if (typeof uploadUrl !== 'string') {
      return err({ type: 'api_error', status: 500, message: 'createUploadSession returned no uploadUrl' });
    }

    // Hardening #3: validate the host before any chunk PUT.
    let host: string;
    try {
      host = new URL(uploadUrl).host;
    } catch {
      return err({ type: 'network_error', message: 'createUploadSession returned an invalid uploadUrl' });
    }
    if (!isAllowedFetchUrlHost(host)) {
      return err({ type: 'network_error', message: `uploadUrl host ${host} not in Microsoft allow-list` });
    }

    // 2. PUT chunks to the pre-signed upload URL — no auth header.
    const total = body.byteLength;
    for (let start = 0; start < total; start += CHUNK_SIZE) {
      const end = Math.min(start + CHUNK_SIZE, total) - 1;
      const chunk = body.slice(start, end + 1);
      try {
        const res = await fetchFn(uploadUrl, {
          method: 'PUT',
          headers: { 'Content-Range': `bytes ${start}-${end}/${total}` },
          body: chunk as unknown as BodyInit,
          signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
        });
        if (!res.ok) {
          // Best-effort session cancellation; ignore failure.
          try {
            await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
          } catch {
            /* ignore */
          }
          return err({ type: 'api_error', status: res.status, message: `chunk PUT failed at byte ${start}` });
        }
        if (res.status === 200 || res.status === 201) {
          return ok(await res.json());
        }
        // 202 Accepted — continue uploading.
      } catch (e: unknown) {
        try {
          await fetchFn(uploadUrl, { method: 'DELETE', signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS) });
        } catch {
          /* ignore */
        }
        return err({ type: 'network_error', message: networkErrorMessage(e) });
      }
    }
    return err({ type: 'api_error', status: 500, message: 'chunked upload completed without final response' });
  };

  const put = async (basePath: string, body: Uint8Array, contentType?: string): Promise<Result<unknown, GraphError>> => {
    if (body.byteLength <= SIMPLE_PUT_THRESHOLD) {
      return simplePut(`${basePath}:/content`, body, contentType);
    }
    return chunkedPut(basePath, body);
  };

  const deleteResource = async (path: string): Promise<Result<unknown, GraphError>> => {
    const headers = await authHeaders();
    if (!headers.ok) return headers;
    try {
      const res = await fetchFn(`https://graph.microsoft.com/v1.0${path}`, {
        method: 'DELETE',
        headers: headers.value,
        signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
      });
      if (!res.ok) return err(await apiErrorFrom(res));
      return ok(undefined);
    } catch (e: unknown) {
      return err({ type: 'network_error', message: networkErrorMessage(e) });
    }
  };

  return {
    get: (path) => request('GET', path),
    post: (path, body) => request('POST', path, body),
    getBinary,
    fetchUrl,
    put,
    delete: deleteResource,
  };
};

export { createGraphClient };
export type { FetchFn, GraphClient, GraphError };
