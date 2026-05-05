import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { formatZodError } from './format-zod-error.ts';

describe('formatZodError', () => {
  it('renders a single missing-field issue as `<path>: <message>`', () => {
    const result = z.object({ query: z.string().min(1) }).safeParse({});
    expect(result.success).toBe(false);
    if (!result.success) {
      const formatted = formatZodError(result.error);
      expect(formatted).toBe('query: Invalid input: expected string, received undefined');
    }
  });

  it('renders a single too-small issue as `<path>: <zod message>`', () => {
    const result = z.object({ query: z.string().min(1) }).safeParse({ query: '' });
    expect(result.success).toBe(false);
    if (!result.success) {
      const formatted = formatZodError(result.error);
      expect(formatted).toBe('query: Too small: expected string to have >=1 characters');
    }
  });

  it('joins multiple issues with `; ` so a multi-field failure stays on one line', () => {
    const result = z.object({ a: z.string().min(1), b: z.string().min(1) }).safeParse({ a: '', b: '' });
    expect(result.success).toBe(false);
    if (!result.success) {
      const formatted = formatZodError(result.error);
      expect(formatted).toContain('a: ');
      expect(formatted).toContain('b: ');
      expect(formatted).toContain('; ');
    }
  });

  it('uses `<root>` as the path label for issues without a path (root-level refinements)', () => {
    const schema = z.string().refine(() => false, { message: 'always fails' });
    const result = schema.safeParse('anything');
    expect(result.success).toBe(false);
    if (!result.success) {
      const formatted = formatZodError(result.error);
      expect(formatted).toBe('<root>: always fails');
    }
  });
});
