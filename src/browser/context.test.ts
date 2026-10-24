import { describe, expect, it, vi } from 'vitest';
import type { BrowserContext, Page } from 'playwright';

import { closeBrowser } from './context.js';

describe('closeBrowser', () => {
  it('closes the browser even when saving the session fails', async () => {
    const close = vi.fn(async () => {});
    const context = {
      storageState: async () => { throw new Error('net::ERR_ABORTED'); },
      close,
    } as unknown as BrowserContext;

    await expect(closeBrowser({ context, page: {} as Page })).rejects.toThrow('net::ERR_ABORTED');

    // A context left open keeps the CLI process alive after it reported the error.
    expect(close).toHaveBeenCalledOnce();
  });
});
