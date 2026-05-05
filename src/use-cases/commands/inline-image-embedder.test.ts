import { describe, expect, it } from 'bun:test';
import { embedInlineImages } from './inline-image-embedder.ts';

describe('embedInlineImages — replace cid: refs in mail body HTML with self-contained data: URIs', () => {
  it('replaces a single cid: ref with a base64 data URI built from the matching attachment', () => {
    const html = '<p>See logo: <img src="cid:logo123" alt="logo"></p>';
    const out = embedInlineImages(html, [{ contentId: 'logo123', contentType: 'image/png', contentBytes: 'iVBORw0=' }]);
    expect(out).toContain('data:image/png;base64,iVBORw0=');
    expect(out).not.toContain('cid:logo123');
  });

  it('replaces multiple distinct cid: refs in the same body', () => {
    const html = '<img src="cid:a"><img src="cid:b">';
    const out = embedInlineImages(html, [
      { contentId: 'a', contentType: 'image/jpeg', contentBytes: 'AAA=' },
      { contentId: 'b', contentType: 'image/gif', contentBytes: 'BBB=' },
    ]);
    expect(out).toContain('data:image/jpeg;base64,AAA=');
    expect(out).toContain('data:image/gif;base64,BBB=');
  });

  it('skips attachments whose contentType is not image/* (Hardening #1: prevent data:text/html injection)', () => {
    const html = '<img src="cid:evil">';
    const out = embedInlineImages(html, [{ contentId: 'evil', contentType: 'text/html', contentBytes: 'PHNjcmlwdD4=' }]);
    expect(out).toContain('cid:evil');
    expect(out).not.toContain('data:text/html');
    expect(out).not.toContain('PHNjcmlwdD4=');
  });

  it('leaves the body unchanged when no attachments match any cid', () => {
    const html = '<img src="cid:nope">';
    const out = embedInlineImages(html, [{ contentId: 'logo123', contentType: 'image/png', contentBytes: 'AAA=' }]);
    expect(out).toBe(html);
  });

  it('leaves the body unchanged when there are no inline attachments at all', () => {
    const html = '<img src="cid:logo">';
    expect(embedInlineImages(html, [])).toBe(html);
  });

  it('skips attachments without a contentId', () => {
    const html = '<img src="cid:logo">';
    const out = embedInlineImages(html, [{ contentId: '', contentType: 'image/png', contentBytes: 'AAA=' }]);
    expect(out).toBe(html);
  });

  it('escapes regex metacharacters in the contentId so they do not break the replace', () => {
    const html = '<img src="cid:foo.bar+baz">';
    const out = embedInlineImages(html, [{ contentId: 'foo.bar+baz', contentType: 'image/png', contentBytes: 'XXX=' }]);
    expect(out).toContain('data:image/png;base64,XXX=');
    expect(out).not.toContain('cid:foo');
  });

  it('does NOT touch non-cid img sources (regular https URLs)', () => {
    const html = '<img src="https://example.com/logo.png">';
    const out = embedInlineImages(html, [{ contentId: 'logo', contentType: 'image/png', contentBytes: 'AAA=' }]);
    expect(out).toBe(html);
  });

  it('treats contentType case-insensitively for the image/ prefix check', () => {
    const html = '<img src="cid:x">';
    const out = embedInlineImages(html, [{ contentId: 'x', contentType: 'IMAGE/PNG', contentBytes: 'YYY=' }]);
    expect(out).toContain('data:IMAGE/PNG;base64,YYY=');
  });
});
