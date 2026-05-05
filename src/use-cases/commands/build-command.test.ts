import { describe, expect, it } from 'bun:test';
import { z } from 'zod';
import { ok } from '../../domain/result.ts';
import type { GraphClient } from '../../infra/graph-client.ts';
import { buildCommand } from './build-command.ts';

const fakeGraph: GraphClient = {
  get: async () => ok({}),
  post: async () => ok({}),
  getBinary: async () => ok({}),
  fetchUrl: async () => ok({}),
  put: async () => ok({}),
  delete: async () => ok({}),
};

describe('buildCommand', () => {
  it('returns err({ type: "validation_error" }) with the zod message when schema validation fails', async () => {
    const cmd = buildCommand((p) => `/items/${p.id}`, z.object({ id: z.string().min(1) }));
    const result = await cmd.execute(fakeGraph, {});
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.type).toBe('validation_error');
      if (result.error.type === 'validation_error') expect(result.error.message).toBe('id: Invalid input: expected string, received undefined');
    }
  });

  it('calls graph.get with the constructed path on valid params', async () => {
    let captured = '';
    const graph: GraphClient = {
      get: async (path: string) => {
        captured = path;
        return ok({});
      },
      post: async () => ok({}),
      getBinary: async () => ok({}),
      fetchUrl: async () => ok({}),
      put: async () => ok({}),
      delete: async () => ok({}),
    };
    const cmd = buildCommand((p) => `/items/${p.id}`, z.object({ id: z.string() }));
    const result = await cmd.execute(graph, { id: '42' });
    expect(result).toEqual(ok({}));
    expect(captured).toBe('/items/42');
  });
});
