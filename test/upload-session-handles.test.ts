/**
 * HTK `--upload-session-handles`: short one-time handles for Graph upload sessions.
 *
 * In the pilot the LLM had to copy the ~1400-character `uploadUrl` (with its `tempauth`
 * token) from `create-upload-session` into the next tool call, and roughly one upload in
 * three failed with SharePoint 401 `badArgument` because one character was corrupted on
 * the way. With the flag on, the model never sees the real URL: the tool result carries a
 * short handle on the attachment listener, and the consumer (nlm-mcp) redeems it once,
 * over the container network, for the real URL.
 */

import { Writable } from 'node:stream';
import { beforeEach, afterEach, describe, expect, it, vi } from 'vitest';
import { z } from 'zod';

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

vi.mock('../src/audit-log.js', () => ({
  auditLog: vi.fn(),
  getUserIdentityForAudit: vi.fn(() => 'user@example.invalid'),
}));

vi.mock('../src/generated/client-beta.js', () => ({ api: { endpoints: [] } }));
vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'create-upload-session',
        method: 'post',
        path: '/drives/:driveId/items/:driveItemId/createUploadSession',
        description: 'Create an upload session.',
        parameters: [
          { name: 'driveId', type: 'Path', schema: z.string() },
          { name: 'driveItemId', type: 'Path', schema: z.string() },
          { name: 'body', type: 'Body', schema: z.object({}).passthrough().optional() },
        ],
      },
      {
        alias: 'get-drive-item',
        method: 'get',
        path: '/drives/:driveId/items/:driveItemId',
        description: 'Get a drive item.',
        parameters: [
          { name: 'driveId', type: 'Path', schema: z.string() },
          { name: 'driveItemId', type: 'Path', schema: z.string() },
        ],
      },
    ],
  },
}));

import logger from '../src/logger.js';
import { auditLog } from '../src/audit-log.js';
import { registerGraphTools } from '../src/graph-tools.js';
import type { GraphClient } from '../src/graph-client.js';
import { AttachmentTicketStore } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import {
  UploadSessionHandleStore,
  UPLOAD_SESSION_MAX_TTL_MS,
  createUploadSessionHandler,
} from '../src/lib/upload-session-handles.js';

const TEMPAUTH = 'v1.eyJhbGciOiJub25lIn0.U0VDUkVULVBBWUxPQUQta3lBRUI.c2lnbmF0dXJlLXZhbHVl';
const REAL_URL =
  'https://htkgmbh.sharepoint.com/sites/Kunden/_api/v2.0/drives/b!abc/items/01XYZ:/deck.pdf:/uploadSession' +
  `?guid=11111111-2222-3333-4444-555555555555&overwrite=True&rename=False&dc=0&tempauth=${TEMPAUTH}`;
const CONFIG = { base: 'http://m365-mcp:3001', key: 'k', keyId: '1', ttlSeconds: 120 };
const HANDLE_RE = /^http:\/\/m365-mcp:3001\/upload-session\/([A-Za-z0-9_-]{43,})$/;

function graphBody(expiration: string) {
  return {
    '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#microsoft.graph.uploadSession',
    expirationDateTime: expiration,
    nextExpectedRanges: ['0-'],
    uploadUrl: REAL_URL,
  };
}

function mockRes(sent: {
  status?: number;
  body?: unknown;
  type?: string;
  headers: Record<string, string>;
}) {
  const res = new Writable({
    write(_chunk, _enc, cb) {
      cb();
    },
  }) as Writable & Record<string, unknown>;
  res.status = (code: number) => {
    sent.status = code;
    return res;
  };
  res.type = (t: string) => {
    sent.type = t;
    return res;
  };
  res.send = (body: unknown) => {
    sent.body = body;
    return res;
  };
  res.setHeader = (name: string, value: string) => {
    sent.headers[name.toLowerCase()] = value;
  };
  return res;
}

async function redeem(store: UploadSessionHandleStore, id: string, method = 'GET', nowMs?: number) {
  const sent: { status?: number; body?: unknown; type?: string; headers: Record<string, string> } =
    { headers: {} };
  const handler = createUploadSessionHandler({
    store,
    now: nowMs === undefined ? undefined : () => nowMs,
  });
  await handler({ method, params: { id } } as never, mockRes(sent) as never, (() => {}) as never);
  return sent;
}

function allLoggedText(): string {
  const calls = [
    ...vi.mocked(logger.info).mock.calls,
    ...vi.mocked(logger.warn).mock.calls,
    ...vi.mocked(logger.error).mock.calls,
    ...vi.mocked(logger.debug).mock.calls,
  ];
  return calls
    .map((c) => c.map((a) => (typeof a === 'string' ? a : JSON.stringify(a))).join(' '))
    .join('\n');
}

function allAuditText(): string {
  return vi
    .mocked(auditLog)
    .mock.calls.map((c) => JSON.stringify(c))
    .join('\n');
}

describe('UploadSessionHandleStore', () => {
  it('mints a URL-safe id of at least 256 bits and redeems it exactly once', () => {
    const store = new UploadSessionHandleStore();
    const now = Date.parse('2026-09-29T10:00:00Z');
    const { id } = store.mint(REAL_URL, '2026-09-29T10:15:00Z', now);
    expect(id).toMatch(/^[A-Za-z0-9_-]{43,}$/);
    expect(Buffer.from(id, 'base64url').length).toBeGreaterThanOrEqual(32);

    expect(store.redeem(id, now + 1000)).toEqual({
      uploadUrl: REAL_URL,
      expirationDateTime: '2026-09-29T10:15:00Z',
    });
    expect(store.redeem(id, now + 2000)).toBeUndefined();
    expect(store.size(now + 2000)).toBe(0);
  });

  it('uses TTL = min(15 minutes, Graph expirationDateTime)', () => {
    const store = new UploadSessionHandleStore();
    const now = Date.parse('2026-09-29T10:00:00Z');
    // Graph expires sooner than 15 minutes: Graph wins.
    const early = store.mint(REAL_URL, '2026-09-29T10:05:00Z', now);
    expect(early.expiresAtMs).toBe(now + 5 * 60_000);
    // Graph expires later than 15 minutes: the cap wins.
    const late = store.mint(REAL_URL, '2026-09-30T10:00:00Z', now);
    expect(late.expiresAtMs).toBe(now + UPLOAD_SESSION_MAX_TTL_MS);
    expect(UPLOAD_SESSION_MAX_TTL_MS).toBe(15 * 60_000);
    // Missing or unparsable expiration: the cap.
    expect(store.mint(REAL_URL, undefined, now).expiresAtMs).toBe(now + UPLOAD_SESSION_MAX_TTL_MS);
    expect(store.mint(REAL_URL, 'not a date', now).expiresAtMs).toBe(
      now + UPLOAD_SESSION_MAX_TTL_MS
    );
  });

  it('refuses to mint a session that has already expired', () => {
    const store = new UploadSessionHandleStore();
    const now = Date.parse('2026-09-29T10:00:00Z');
    expect(() => store.mint(REAL_URL, '2026-09-29T09:59:00Z', now)).toThrow(/already expired/);
    expect(store.size(now)).toBe(0);
  });

  it('wipes an expired ticket and refuses it', () => {
    const store = new UploadSessionHandleStore();
    const now = Date.parse('2026-09-29T10:00:00Z');
    const { id, expiresAtMs } = store.mint(REAL_URL, '2026-09-29T10:15:00Z', now);
    expect(store.size(now)).toBe(1);
    expect(store.redeem(id, expiresAtMs)).toBeUndefined();
    expect(store.size(expiresAtMs)).toBe(0);
  });

  it('wipes an expired ticket on its own timer, without any further access', () => {
    vi.useFakeTimers();
    try {
      const store = new UploadSessionHandleStore();
      const now = Date.now();
      store.mint(REAL_URL, new Date(now + 60_000).toISOString(), now);
      expect(store.liveCountUnswept()).toBe(1);
      vi.advanceTimersByTime(60_001);
      expect(store.liveCountUnswept()).toBe(0);
    } finally {
      vi.useRealTimers();
    }
  });

  it('never exposes the real URL through inspection of the store', () => {
    const store = new UploadSessionHandleStore();
    store.mint(REAL_URL, undefined);
    expect(JSON.stringify(store)).not.toContain('tempauth');
    expect(String(store)).not.toContain('tempauth');
  });
});

describe('redemption route', () => {
  it('returns the real URL exactly once, with no-store; a second redeem is 404', async () => {
    const store = new UploadSessionHandleStore();
    const { id } = store.mint(REAL_URL, '2099-01-01T00:00:00Z');

    const first = await redeem(store, id);
    expect(first.status).toBe(200);
    expect(first.headers['cache-control']).toBe('no-store');
    expect(first.type).toMatch(/json/);
    expect(JSON.parse(String(first.body))).toEqual({
      uploadUrl: REAL_URL,
      expirationDateTime: '2099-01-01T00:00:00Z',
    });

    const second = await redeem(store, id);
    expect(second.status).toBe(404);
    expect(second.body).toBe('Not found');
  });

  it('gives an unknown id the same fixed 404 body', async () => {
    const store = new UploadSessionHandleStore();
    const sent = await redeem(store, 'doesnotexist');
    expect(sent.status).toBe(404);
    expect(sent.body).toBe('Not found');
  });

  it('gives an expired id the same 404 and wipes it', async () => {
    const store = new UploadSessionHandleStore();
    const now = Date.parse('2026-09-29T10:00:00Z');
    const { id, expiresAtMs } = store.mint(REAL_URL, '2026-09-29T10:15:00Z', now);
    const sent = await redeem(store, id, 'GET', expiresAtMs + 1);
    expect(sent.status).toBe(404);
    expect(sent.body).toBe('Not found');
    expect(store.size(expiresAtMs + 1)).toBe(0);
  });

  it('refuses HEAD without burning the ticket', async () => {
    const store = new UploadSessionHandleStore();
    const { id } = store.mint(REAL_URL, undefined);
    const head = await redeem(store, id, 'HEAD');
    expect(head.status).toBe(405);
    expect(store.size()).toBe(1);
  });

  it('logs nothing that carries the real URL or the handle id', async () => {
    vi.clearAllMocks();
    const store = new UploadSessionHandleStore();
    const { id } = store.mint(REAL_URL, undefined);
    await redeem(store, id);
    await redeem(store, id);
    const logged = allLoggedText();
    expect(logged).not.toContain(TEMPAUTH);
    expect(logged).not.toContain('tempauth');
    expect(logged).not.toContain(id);
  });
});

describe('create-upload-session tool result', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn>; registerTool: ReturnType<typeof vi.fn> };
  let graphClient: GraphClient;
  let graphResponse: { content: Array<{ type: string; text: string }>; isError?: boolean };

  beforeEach(() => {
    vi.clearAllMocks();
    mockServer = { tool: vi.fn(), registerTool: vi.fn() };
    graphResponse = {
      content: [{ type: 'text', text: JSON.stringify(graphBody('2099-01-01T00:00:00Z')) }],
    };
    graphClient = {
      graphRequest: vi.fn(async () => structuredClone(graphResponse)),
      serialize: (data: unknown) => JSON.stringify(data),
    } as unknown as GraphClient;
  });

  afterEach(() => {
    resetAttachmentMinting();
  });

  function handler(name: string) {
    registerGraphTools(mockServer, graphClient, false, undefined, true);
    const call = mockServer.registerTool.mock.calls.find((c: unknown[]) => c[0] === name);
    expect(call).toBeDefined();
    return call![call!.length - 1] as (params: Record<string, unknown>) => Promise<{
      content: Array<{ text: string }>;
      isError?: boolean;
    }>;
  }

  function enable(): UploadSessionHandleStore {
    const uploadSessions = new UploadSessionHandleStore();
    configureAttachmentMinting({
      store: new AttachmentTicketStore(120),
      config: CONFIG,
      uploadSessions,
    });
    return uploadSessions;
  }

  const params = { driveId: 'b!abc', driveItemId: '01XYZ:/deck.pdf:', body: {} };

  it('with the flag on, replaces uploadUrl with a handle and keeps every other field', async () => {
    const store = enable();
    const result = await handler('create-upload-session')(params);
    expect(result.isError).toBeFalsy();
    const body = JSON.parse(result.content[0].text);

    expect(body.uploadUrl).toMatch(HANDLE_RE);
    expect(body.uploadUrlIsHandle).toBe(true);
    const expected = graphBody('2099-01-01T00:00:00Z') as Record<string, unknown>;
    for (const [key, value] of Object.entries(expected)) {
      if (key === 'uploadUrl') continue;
      expect(body[key]).toEqual(value);
    }
    expect(Object.keys(body).sort()).toEqual(
      [...Object.keys(expected), 'uploadUrlIsHandle'].sort()
    );

    // The handle redeems for the real URL.
    const id = HANDLE_RE.exec(body.uploadUrl)![1];
    const sent = await redeem(store, id);
    expect(JSON.parse(String(sent.body)).uploadUrl).toBe(REAL_URL);
  });

  it('keeps the real URL out of the tool result, the logs and the audit stream', async () => {
    enable();
    const result = await handler('create-upload-session')(params);
    const everything = JSON.stringify(result);
    expect(everything).not.toContain(TEMPAUTH);
    expect(everything).not.toContain('tempauth');
    expect(everything).not.toContain('sharepoint.com');

    expect(vi.mocked(auditLog)).toHaveBeenCalled();
    const audit = allAuditText();
    expect(audit).toContain('create-upload-session');
    expect(audit).not.toContain(TEMPAUTH);
    expect(audit).not.toContain('tempauth');

    const logged = allLoggedText();
    expect(logged).not.toContain(TEMPAUTH);
    expect(logged).not.toContain('tempauth');
  });

  it('fails closed, without the real URL, when no handle can be minted', async () => {
    const store = enable();
    vi.spyOn(store, 'mint').mockImplementation(() => {
      throw new Error('No upload-session handle slots available');
    });
    const result = await handler('create-upload-session')(params);
    expect(result.isError).toBe(true);
    expect(JSON.stringify(result)).not.toContain('tempauth');
  });

  it('leaves a Graph error untouched', async () => {
    enable();
    graphResponse = {
      content: [{ type: 'text', text: JSON.stringify({ error: 'Graph said no' }) }],
      isError: true,
    };
    const result = await handler('create-upload-session')(params);
    expect(result.isError).toBe(true);
    expect(JSON.parse(result.content[0].text)).toEqual({ error: 'Graph said no' });
  });

  it('does not touch other tools', async () => {
    enable();
    graphResponse = {
      content: [{ type: 'text', text: JSON.stringify({ id: 'x', uploadUrl: 'https://keep.me' }) }],
    };
    const result = await handler('get-drive-item')({ driveId: 'd', driveItemId: 'i' });
    expect(JSON.parse(result.content[0].text)).toEqual({ id: 'x', uploadUrl: 'https://keep.me' });
  });

  it('with the flag off (feature on), returns the upstream result unchanged', async () => {
    configureAttachmentMinting({ store: new AttachmentTicketStore(120), config: CONFIG });
    const result = await handler('create-upload-session')(params);
    expect(result.content[0].text).toBe(graphResponse.content[0].text);
  });

  it('with attachment URLs off entirely, returns the upstream result unchanged', async () => {
    resetAttachmentMinting();
    const result = await handler('create-upload-session')(params);
    expect(result.content[0].text).toBe(graphResponse.content[0].text);
    expect(JSON.parse(result.content[0].text).uploadUrl).toBe(REAL_URL);
  });
});
