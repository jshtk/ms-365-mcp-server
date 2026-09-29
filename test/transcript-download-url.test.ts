import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { isMintableMeetingBytePath, UTILITY_TOOLS } from '../src/graph-tools.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';

describe('get-download-url: meeting transcript content is mintable', () => {
  it.each([
    '/me/onlineMeetings/m1/transcripts/t1/content',
    '/users/u1/onlineMeetings/m1/transcripts/t1/content',
    '/me/onlineMeetings/m1/recordings/r1/content',
    '/me/onlineMeetings/m1/recordings/r1',
    '/communications/calls/c1/recordings/r1',
  ])('accepts %s', (p) => expect(isMintableMeetingBytePath(p)).toBe(true));

  it.each([
    '/me/onlineMeetings/m1/transcripts',
    '/me/onlineMeetings/m1/transcripts/t1',
    '/me/onlineMeetings/m1/transcripts/t1/metadataContent',
    '/me/messages/x',
  ])('rejects %s', (p) => expect(isMintableMeetingBytePath(p)).toBe(false));
});

describe('get-download-url tool: transcript content target', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
  const TRANSCRIPT = '/me/onlineMeetings/m1/transcripts/t1/content';

  const ctx = {
    graphClient: {} as never,
    authManager: {
      isOAuthModeEnabled: () => false,
      isMultiAccount: async () => false,
      getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
    } as never,
    multiAccount: false,
    accountNames: [],
  };

  const parse = (r: unknown) =>
    JSON.parse((r as { content: Array<{ text: string }> }).content[0].text);

  afterEach(() => resetAttachmentMinting());

  describe('with --enable-attachment-urls', () => {
    beforeEach(() => {
      configureAttachmentMinting({
        store: new AttachmentTicketStore(120),
        config: { base: 'http://m365-mcp:3001', key: 'k', keyId: '1', ttlSeconds: 120 },
      });
    });

    it('mints a single-use server URL', async () => {
      const result = await tool.execute({ target: TRANSCRIPT }, ctx);
      const body = parse(result);
      expect((result as { isError?: boolean }).isError).toBeFalsy();
      expect(body.downloadUrl).toMatch(/^http:\/\/m365-mcp:3001\/attachment\?/);
      expect(body.singleUse).toBe(true);
    });
  });

  describe('without --enable-attachment-urls', () => {
    it('refuses with a message naming transcripts', async () => {
      const result = await tool.execute({ target: TRANSCRIPT }, ctx);
      expect((result as { isError?: boolean }).isError).toBe(true);
      expect(parse(result).error).toMatch(/Meeting recordings and transcripts/);
    });
  });
});
