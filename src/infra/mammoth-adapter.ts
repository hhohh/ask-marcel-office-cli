import mammoth from 'mammoth';
import type { Result } from '../domain/result.ts';
import { err, ok } from '../domain/result.ts';
import type { GraphError } from './graph-client.ts';

const mammothToHtml = async (bytes: Uint8Array): Promise<Result<string, GraphError>> => {
  try {
    const result = await mammoth.convertToHtml({ buffer: Buffer.from(bytes) });
    return ok(result.value);
  } catch (e) {
    return err({ type: 'api_error', status: 500, message: `docx conversion failed: ${e instanceof Error ? e.message : String(e)}` });
  }
};

export { mammothToHtml };
