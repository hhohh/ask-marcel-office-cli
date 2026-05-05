/**
 * Email body HTML often references inline attachments via `cid:<id>`
 * URLs (e.g. `<img src="cid:logo123">`). Those `cid:` refs are
 * meaningless outside the original mail viewer. To make markdown /
 * HTML output self-contained we replace each `cid:<id>` with a base64
 * `data:` URI sourced from the matching inline attachment.
 *
 * Hardening #1: only `image/*` content-types are embedded. Anything
 * else (`text/html`, `application/javascript`, ...) would let an
 * attacker turn an attachment into an executable payload via the
 * data URI; skip and leave the original `cid:` ref in place.
 */

type InlineAttachment = {
  readonly contentId: string;
  readonly contentType: string;
  readonly contentBytes: string;
};

const embedInlineImages = (html: string, attachments: ReadonlyArray<InlineAttachment>): string => {
  let out = html;
  for (const a of attachments) {
    if (a.contentId === '') continue;
    if (!a.contentType.toLowerCase().startsWith('image/')) continue;
    const dataUri = `data:${a.contentType};base64,${a.contentBytes}`;
    // String.replaceAll(searchValue: string, replaceValue) does a literal
    // global replace — no regex, so contentId metacharacters (`.`, `+`,
    // etc.) are treated literally with no escape step needed.
    out = out.replaceAll(`cid:${a.contentId}`, dataUri);
  }
  return out;
};

export { embedInlineImages };
export type { InlineAttachment };
