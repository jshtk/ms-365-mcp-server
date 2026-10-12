import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { Writable } from 'node:stream';
import { UTILITY_TOOLS, isRequestIdentityMintable, mintAcceptFor } from '../src/graph-tools.js';
import { requestContext } from '../src/request-context.js';
import { AttachmentTicketStore, TICKET_PARAM } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import { createAttachmentHandler } from '../src/attachment-route.js';

/**
 * HTK policy on top of upstream request-identity minting (#692): under the caller's
 * token only meeting recordings/transcripts and driveItem-less $value endpoints get a
 * single-use URL; mail, event and group attachments and the driveItem fallback keep
 * the refusal. Transcript tickets redeem with Accept: text/vtt (pilot 29.09.2026).
 */
describe('HTK: request-identity minting policy', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;
  const TRANSCRIPT = '/me/onlineMeetings/MMM/transcripts/TTT/content';
  const RECORDING = '/me/onlineMeetings/MMM/recordings/RRR/content';
  const PHOTO = '/me/photo/$value';
  let store: AttachmentTicketStore;

  function ctx() {
    return {
      graphClient: {} as never,
      authManager: {
        isOAuthModeEnabled: () => false,
        isMultiAccount: async () => false,
        getToken: async () => 'SERVER_OWN_TOKEN',
        getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
      } as never,
      multiAccount: false,
      accountNames: [],
    };
  }

  const parse = (result: unknown) =>
    JSON.parse((result as { content: Array<{ text: string }> }).content[0].text);
  const ticketIdOf = (result: unknown) =>
    new URL(parse(result).downloadUrl).searchParams.get(TICKET_PARAM)!;
  const asCaller = <T>(fn: () => Promise<T>) =>
    requestContext.run({ accessToken: 'CALLER_TOKEN' }, fn);

  beforeEach(() => {
    store = new AttachmentTicketStore(120);
    configureAttachmentMinting({
      store,
      config: { base: 'http://m365-mcp:3001', key: 'k', keyId: '1', ttlSeconds: 120 },
    });
  });

  afterEach(() => resetAttachmentMinting());

  it.each([
    ['/me/messages/A/attachments/B/$value', false],
    ['/users/u/events/A/attachments/B/$value', false],
    ['/groups/g/messages/A/attachments/B/$value', false],
    ['/groups/g/events/A/attachments/B/$value', false],
    ['/drives/d/items/i/content', false],
    [TRANSCRIPT, true],
    [RECORDING, true],
    [PHOTO, true],
  ])('isRequestIdentityMintable(%s) === %s', (target, expected) => {
    expect(isRequestIdentityMintable(target)).toBe(expected);
  });

  it('mints a transcript URL under the caller token, carrying Accept: text/vtt', async () => {
    const result = await asCaller(() => tool.execute({ target: TRANSCRIPT }, ctx()));
    expect(result.isError).toBeFalsy();
    expect(parse(result).downloadUrl).toMatch(/^http:\/\/m365-mcp:3001\/attachment\?t=/);
    expect(store.redeem(ticketIdOf(result))).toMatchObject({
      kind: 'request-token',
      accessToken: 'CALLER_TOKEN',
      target: TRANSCRIPT,
      accept: 'text/vtt',
    });
  });

  it('mints recordings without an Accept override', async () => {
    const result = await asCaller(() => tool.execute({ target: RECORDING }, ctx()));
    const ticket = store.redeem(ticketIdOf(result))!;
    expect(ticket.accept).toBeUndefined();
  });

  it('refuses mail attachments under the caller token', async () => {
    const result = await asCaller(() =>
      tool.execute({ target: '/me/messages/A/attachments/B/$value' }, ctx())
    );
    expect(result.isError).toBe(true);
    expect(parse(result).error).toMatch(/download-bytes/);
    expect(store.size()).toBe(0);
  });

  it('still mints mail attachments under the server identity (upstream unchanged)', async () => {
    const result = await tool.execute({ target: '/me/messages/A/attachments/B/$value' }, ctx());
    expect(store.redeem(ticketIdOf(result))).toMatchObject({ kind: 'server-account' });
  });

  it('mintAcceptFor only matches transcript content', () => {
    expect(mintAcceptFor(TRANSCRIPT)).toBe('text/vtt');
    expect(mintAcceptFor('/users/u/onlineMeetings/M/transcripts/T/content')).toBe('text/vtt');
    expect(mintAcceptFor(RECORDING)).toBeUndefined();
    expect(mintAcceptFor('/me/onlineMeetings/M/transcripts/T')).toBeUndefined();
  });

  it('redemption sends Accept: text/vtt with the caller token', async () => {
    const result = await asCaller(() => tool.execute({ target: TRANSCRIPT }, ctx()));
    const downloadStream = vi.fn(async () => ({
      body: new ReadableStream({
        start(c) {
          c.enqueue(new TextEncoder().encode('WEBVTT\n'));
          c.close();
        },
      }),
      contentType: 'text/vtt',
      contentLength: 7,
      contentDisposition: null,
    }));
    const getToken = vi.fn(async () => 'SERVER_OWN_TOKEN');
    const getTokenForAccount = vi.fn(async () => 'SERVER_OWN_TOKEN');
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => ({ downloadStream }) as never,
      authManager: { isOAuthModeEnabled: () => false, getToken, getTokenForAccount } as never,
    });

    let status: number | undefined;
    const res = new Writable({
      write(_chunk, _enc, cb) {
        cb();
      },
    }) as Writable & Record<string, unknown>;
    res.status = (code: number) => {
      status = code;
      return res;
    };
    res.type = () => res;
    res.send = () => res;
    res.setHeader = () => {};

    // Redemption runs outside any request context, as nlm-mcp's fetch does.
    await handler(
      { method: 'GET', query: { [TICKET_PARAM]: ticketIdOf(result) } } as never,
      res as never,
      (() => {}) as never
    );

    expect(status).toBe(200);
    expect(downloadStream).toHaveBeenCalledWith(TRANSCRIPT, {
      accessToken: 'CALLER_TOKEN',
      headers: { Accept: 'text/vtt' },
    });
    expect(getToken).not.toHaveBeenCalled();
    expect(getTokenForAccount).not.toHaveBeenCalled();
  });
});
