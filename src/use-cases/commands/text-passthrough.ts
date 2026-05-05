/**
 * Microsoft Graph's `?format=pdf` / `?format=html` conversion accepts a
 * fixed set of source extensions (csv, doc, docx, odp, ods, odt, pot,
 * potm, potx, pps, ppsx, ppsxm, ppt, pptm, pptx, rtf, xls, xlsx). For
 * anything outside that whitelist Graph returns 4xx, which is a
 * cryptic failure mode for an LLM agent.
 *
 * The conversion commands short-circuit on the extensions below: they
 * return the raw bytes instead of attempting (and wasting) a Graph
 * round-trip. Hardening #1 of the format-conversion plan.
 */

const PLAIN_TEXT_EXTENSIONS: ReadonlySet<string> = new Set(['txt', 'md', 'markdown', 'html', 'htm', 'json', 'xml', 'log', 'yaml', 'yml', 'css', 'js', 'ts', 'sh', 'sql']);

const isPlainTextFilename = (name: string): boolean => {
  const dot = name.lastIndexOf('.');
  if (dot === -1) return false;
  return PLAIN_TEXT_EXTENSIONS.has(name.slice(dot + 1).toLowerCase());
};

export { isPlainTextFilename, PLAIN_TEXT_EXTENSIONS };
