import type { z } from 'zod';

/**
 * Render a Zod validation error as a single human-readable line.
 *
 * Default `error.message` is a stringified JSON array of every issue —
 * useful for debuggers, hostile to LLM agents and users at the CLI.
 * This helper picks each issue's `path` and `message` and joins them
 * with `; ` so the user sees `query: Too small: expected string to
 * have >=1 characters` instead of a 200-character JSON dump.
 */
const formatZodError = (error: z.ZodError): string => {
  if (error.issues.length === 0) return 'validation failed';
  return error.issues
    .map((issue) => {
      const path = issue.path.length > 0 ? issue.path.join('.') : '<root>';
      return `${path}: ${issue.message}`;
    })
    .join('; ');
};

export { formatZodError };
