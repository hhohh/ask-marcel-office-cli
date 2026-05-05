/**
 * Pure helpers for `extract-sharepoint-links-in-mail`.
 *
 * `extractSharepointUrls(htmlBody)` finds every `https://*.sharepoint.com/...`
 * URL inside an HTML mail body — both `<a href=...>` and bare-text
 * occurrences — and returns the deduplicated list.
 *
 * `buildShareToken(url)` encodes a URL for Graph's `/shares/{token}`
 * resolver per [shares-get](https://learn.microsoft.com/en-us/graph/api/shares-get):
 * `u!` + base64url(url) with no padding.
 */

const SP_URL_PATTERN = /https:\/\/[\w-]+(?:\.[\w-]+)*\.sharepoint\.com[^\s"'<>)]*/gi;

const stripFragment = (url: string): string => {
  const hash = url.indexOf('#');
  return hash === -1 ? url : url.slice(0, hash);
};

const extractSharepointUrls = (htmlBody: string): ReadonlyArray<string> => {
  const matches = htmlBody.match(SP_URL_PATTERN) ?? [];
  const seen = new Set<string>();
  const out: string[] = [];
  for (const raw of matches) {
    const cleaned = stripFragment(raw);
    if (seen.has(cleaned)) continue;
    seen.add(cleaned);
    out.push(cleaned);
  }
  return out;
};

const buildShareToken = (url: string): string => {
  const b64 = btoa(url).replaceAll('+', '-').replaceAll('/', '_').replaceAll('=', '');
  return `u!${b64}`;
};

export { buildShareToken, extractSharepointUrls };
