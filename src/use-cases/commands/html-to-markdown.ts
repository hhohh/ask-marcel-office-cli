import TurndownService from 'turndown';

/**
 * Turn arbitrary HTML — Graph-converted Office docs, OneNote pages,
 * Outlook email bodies, SharePoint pages — into clean markdown.
 *
 * Wrapped in a function-only adapter so the rest of the codebase stays
 * class-free per atelier rule 1. Each call constructs a fresh service
 * (turndown is cheap to instantiate; no mutable shared state).
 *
 * Turndown's default HTML parser drops comment nodes during parsing,
 * so we don't need a custom comment-stripping rule. <script> and
 * <style> aren't dropped by default — those need explicit removal so
 * their text content doesn't leak into the markdown output.
 */

const buildService = (): TurndownService => {
  const td = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-',
    emDelimiter: '_',
    strongDelimiter: '**',
  });
  td.remove(['script', 'style']);
  return td;
};

const htmlToMarkdown = (html: string): string => buildService().turndown(html);

export { htmlToMarkdown };
