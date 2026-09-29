/**
 * HTK: request-identity-bound minting (`--mint-with-request-identity`).
 *
 * Upstream refuses to mint whenever Graph identity comes from the request, because a
 * ticket redeemed later under the server's own token would read as a different identity.
 * The opt-in keeps that guarantee a different way: the ticket carries the caller's own
 * access token (memory only), and redemption fetches with exactly that token -- never
 * with the server's token cache.
 */

import { Writable } from 'node:stream';
import { inspect } from 'node:util';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
    verbose: vi.fn(),
  },
  enableConsoleLogging: vi.fn(),
}));

import logger from '../src/logger.js';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import { requestContext } from '../src/request-context.js';
import { AttachmentTicketStore, TICKET_PARAM } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import { createAttachmentHandler } from '../src/attachment-route.js';
import GraphClient from '../src/graph-client.js';

const TRANSCRIPT = '/me/onlineMeetings/m1/transcripts/t1/content';
const TOKEN_ALICE = 'eyJ.ALICE-ACCESS-TOKEN.sig';
const TOKEN_BOB = 'eyJ.BOB-ACCESS-TOKEN.sig';
const CONFIG = { base: 'http://m365-mcp:3001', key: 'k', keyId: '1', ttlSeconds: 120 };

const tool = UTILITY_TOOLS.find((t) => t.name === 'get-download-url')!;

/** The server's own token cache must never be touched on this path. */
const serverAuth = {
  isOAuthModeEnabled: () => false,
  isMultiAccount: async () => false,
  getToken: async () => {
    throw new Error('server token cache must not be used');
  },
  getTokenForAccount: async () => {
    throw new Error('server token cache must not be used');
  },
};

function toolCtx(graphClient: unknown = {}) {
  return {
    graphClient: graphClient as never,
    authManager: serverAuth as never,
    multiAccount: false,
    accountNames: [],
  };
}

function parse(result: unknown) {
  return JSON.parse((result as { content: Array<{ text: string }> }).content[0].text);
}

async function mintAs(token: string, target = TRANSCRIPT) {
  return requestContext.run({ accessToken: token }, () => tool.execute({ target }, toolCtx()));
}

function ticketOf(downloadUrl: string): string {
  return new URL(downloadUrl).searchParams.get(TICKET_PARAM)!;
}

function mockRes(sent: { status?: number; body?: unknown }) {
  const res = new Writable({
    write(_chunk, _enc, cb) {
      cb();
    },
  }) as Writable & Record<string, unknown>;
  res.status = (code: number) => {
    sent.status = code;
    return res;
  };
  res.type = () => res;
  res.send = (body: unknown) => {
    sent.body = body;
    return res;
  };
  res.setHeader = () => {};
  return res;
}

describe('request-identity-bound minting', () => {
  let store: AttachmentTicketStore;
  let authHeaders: string[];
  const originalFetch = global.fetch;

  const graphClient = new GraphClient(
    serverAuth as never,
    { clientId: 'x', tenantId: 'common', cloudType: 'global' } as never,
    'json'
  );

  async function redeem(ticket: string) {
    const sent: { status?: number; body?: unknown } = {};
    const handler = createAttachmentHandler({
      store,
      getGraphClient: () => graphClient,
      authManager: serverAuth as never,
    });
    await handler(
      { method: 'GET', query: { [TICKET_PARAM]: ticket } } as never,
      mockRes(sent) as never,
      (() => {}) as never
    );
    return sent;
  }

  function allLogText(): string {
    const l = logger as unknown as Record<string, { mock: { calls: unknown[][] } }>;
    return ['info', 'warn', 'error', 'debug', 'verbose']
      .flatMap((level) => l[level].mock.calls)
      .map((call) => call.map((a) => (typeof a === 'string' ? a : inspect(a))).join(' '))
      .join('\n');
  }

  beforeEach(() => {
    vi.clearAllMocks();
    store = new AttachmentTicketStore(120);
    authHeaders = [];
    global.fetch = (async (_url: string, init: RequestInit) => {
      authHeaders.push((init.headers as Record<string, string>).Authorization);
      return new Response('WEBVTT\n\n00:00.000 --> 00:01.000\n<v A>hi</v>\n', {
        status: 200,
        headers: { 'content-type': 'text/vtt' },
      });
    }) as typeof fetch;
  });

  afterEach(() => {
    global.fetch = originalFetch;
    resetAttachmentMinting();
  });

  describe('with --mint-with-request-identity', () => {
    beforeEach(() => {
      configureAttachmentMinting({ store, config: CONFIG, mintWithRequestIdentity: true });
    });

    it('mints under a request token and redeems with exactly that bearer', async () => {
      const result = await mintAs(TOKEN_ALICE);
      expect((result as { isError?: boolean }).isError).toBeFalsy();
      const { downloadUrl } = parse(result);
      expect(downloadUrl).toMatch(/^http:\/\/m365-mcp:3001\/attachment\?/);

      const sent = await redeem(ticketOf(downloadUrl));
      expect(sent.status).toBe(200);
      expect(authHeaders).toEqual([`Bearer ${TOKEN_ALICE}`]);
    });

    it('is single-use', async () => {
      const ticket = ticketOf(parse(await mintAs(TOKEN_ALICE)).downloadUrl);
      expect((await redeem(ticket)).status).toBe(200);
      expect((await redeem(ticket)).status).toBe(404);
      expect(authHeaders).toHaveLength(1);
      expect(store.boundTokenCount()).toBe(0);
    });

    it('rejects an expired ticket and wipes its token', () => {
      const NOW = 1_780_000_000_000;
      const { id } = store.mint(TRANSCRIPT, undefined, NOW, { boundAccessToken: TOKEN_ALICE });
      expect(store.boundTokenCount(NOW)).toBe(1);
      expect(store.redeem(id, NOW + 121_000)).toBeUndefined();
      expect(store.boundTokenCount(NOW + 121_000)).toBe(0);
    });

    it('wipes an expired token on sweep even if never redeemed', () => {
      const NOW = 1_780_000_000_000;
      store.mint(TRANSCRIPT, undefined, NOW, { boundAccessToken: TOKEN_ALICE });
      expect(store.size(NOW + 120_000)).toBe(0);
      expect(store.boundTokenCount(NOW + 120_000)).toBe(0);
    });

    it('keeps the token out of the URL, the tool output, the logs and any serialization', async () => {
      const result = await mintAs(TOKEN_ALICE);
      const text = (result as { content: Array<{ text: string }> }).content[0].text;
      expect(text).not.toContain(TOKEN_ALICE);
      expect(text).not.toContain('ALICE-ACCESS-TOKEN');

      const ticket = ticketOf(parse(result).downloadUrl);
      const NOW = Date.now();
      const redeemed = store.redeem(ticket, NOW)!;
      expect(redeemed).toBeDefined();
      expect(JSON.stringify(redeemed)).not.toContain(TOKEN_ALICE);
      expect(inspect(redeemed, { depth: 5 })).not.toContain(TOKEN_ALICE);
      expect(Object.keys(redeemed)).not.toContain('boundAccessToken');

      // Full mint + redeem round trip, including a failing upstream (error path logs).
      const second = ticketOf(parse(await mintAs(TOKEN_ALICE)).downloadUrl);
      global.fetch = (async () =>
        new Response(`denied for ${TOKEN_ALICE}`, { status: 500 })) as typeof fetch;
      expect((await redeem(second)).status).toBe(502);
      expect(allLogText()).not.toContain('ALICE-ACCESS-TOKEN');
    });

    it("always fetches one caller's ticket with that caller's token", async () => {
      const alice = ticketOf(parse(await mintAs(TOKEN_ALICE)).downloadUrl);
      const bob = ticketOf(parse(await mintAs(TOKEN_BOB)).downloadUrl);

      // Redeem in reverse order, and from inside another caller's request context,
      // to show the bound token -- not ambient state -- decides the identity.
      await requestContext.run({ accessToken: TOKEN_ALICE }, () => redeem(bob));
      await requestContext.run({ accessToken: TOKEN_BOB }, () => redeem(alice));
      expect(authHeaders).toEqual([`Bearer ${TOKEN_BOB}`, `Bearer ${TOKEN_ALICE}`]);
    });

    it('does not extend to the drive-item fallback (only byte endpoints)', async () => {
      const driveClient = {
        graphRequest: async () => ({ content: [{ type: 'text', text: '{"id":"x"}' }] }),
      };
      const result = await requestContext.run({ accessToken: TOKEN_ALICE }, () =>
        tool.execute({ target: '/drives/d1/items/i1/content' }, toolCtx(driveClient))
      );
      expect((result as { isError?: boolean }).isError).toBe(true);
      expect(parse(result).error).toMatch(/Graph identity comes from the request/);
      expect(store.size()).toBe(0);
    });

    it('still uses the server identity path unchanged when no request token exists', async () => {
      const serverOwn = {
        ...serverAuth,
        getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
      };
      const result = await tool.execute(
        { target: TRANSCRIPT },
        { ...toolCtx(), authManager: serverOwn as never }
      );
      const ticket = ticketOf(parse(result).downloadUrl);
      expect(store.boundTokenCount()).toBe(0);
      const sent: { status?: number } = {};
      await createAttachmentHandler({
        store,
        getGraphClient: () => graphClient,
        authManager: serverOwn as never,
      })(
        { method: 'GET', query: { [TICKET_PARAM]: ticket } } as never,
        mockRes(sent) as never,
        (() => {}) as never
      );
      expect(sent.status).toBe(200);
      expect(authHeaders).toEqual(['Bearer SERVER_OWN_TOKEN']);
    });
  });

  describe('without the flag', () => {
    beforeEach(() => {
      configureAttachmentMinting({ store, config: CONFIG });
    });

    it('keeps the upstream refusal for request identity', async () => {
      const result = await mintAs(TOKEN_ALICE);
      expect((result as { isError?: boolean }).isError).toBe(true);
      expect(parse(result).error).toMatch(/Graph identity comes from the request/);
      expect(store.size()).toBe(0);
    });
  });
});
