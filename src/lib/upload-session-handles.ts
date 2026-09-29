/**
 * HTK `--upload-session-handles`: short, single-use handles for Graph upload sessions.
 *
 * Why this exists: `create-upload-session` returns a ~1400-character `uploadUrl` whose
 * `tempauth` query parameter is itself a bearer credential for the upload. An agent had
 * to copy that string verbatim into its next tool call, and in the pilot roughly one copy
 * in three came out with one character changed -- SharePoint then answers 401
 * `badArgument`. Asking a language model to reproduce a long random string exactly is the
 * bug; no prompt fixes it.
 *
 * With the flag on, the tool result carries `http://<attachment base>/upload-session/<id>`
 * instead, and the consumer (a sidecar on the container network) redeems that handle once
 * on the attachment listener for the real URL. The model never sees the real URL at all,
 * which also keeps the `tempauth` credential out of the model context and transcripts.
 *
 * Same shape as the attachment ticket store, deliberately: memory only, 256-bit CSPRNG
 * ids, single redemption, one undifferentiated 404 for unknown/used/expired, a cap on
 * live entries. One difference: each entry also has its own unref'd timer, so an expired
 * upload credential is dropped from memory at its expiry rather than whenever the store
 * is next touched.
 */

import { randomBytes } from 'node:crypto';
import type { Handler, Request, Response } from 'express';

/** Route the handle points at. Mounted only on the dedicated attachment listener. */
export const UPLOAD_SESSION_ROUTE = '/upload-session';

/** Upper bound on a handle's lifetime, whatever Graph says. */
export const UPLOAD_SESSION_MAX_TTL_MS = 15 * 60 * 1000;

/** 32 bytes of CSPRNG output: the handle id is the whole capability. */
const HANDLE_BYTES = 32;

/** Cap on live handles; refuses rather than evicting, like the attachment store. */
const MAX_LIVE_HANDLES = 256;

/** base64url of 32 bytes is 43 characters, no padding. Anything else is not ours. */
const HANDLE_ID_RE = /^[A-Za-z0-9_-]{43}$/;

/** Same body for every refusal, so a probe cannot tell unknown from used from expired. */
const NOT_FOUND_BODY = 'Not found';

export class UploadSessionHandleError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'UploadSessionHandleError';
  }
}

interface Entry {
  uploadUrl: string;
  expirationDateTime: string;
  expiresAtMs: number;
  timer: ReturnType<typeof setTimeout>;
}

export interface RedeemedUploadSession {
  uploadUrl: string;
  expirationDateTime: string;
}

export class UploadSessionHandleStore {
  // Private and non-enumerable in spirit: nothing here is ever serialised or logged.
  #entries = new Map<string, Entry>();

  private sweep(nowMs: number): void {
    for (const [id, entry] of this.#entries) {
      if (entry.expiresAtMs <= nowMs) this.forget(id);
    }
  }

  private forget(id: string): void {
    const entry = this.#entries.get(id);
    if (!entry) return;
    clearTimeout(entry.timer);
    this.#entries.delete(id);
  }

  /**
   * Store the real upload URL and return a handle id.
   * TTL = min(15 minutes, Graph's expirationDateTime); a missing or unparsable
   * expiration falls back to the cap.
   */
  mint(
    uploadUrl: string,
    expirationDateTime: string | undefined,
    nowMs: number = Date.now()
  ): { id: string; expiresAtMs: number } {
    this.sweep(nowMs);
    const graphExpiry =
      typeof expirationDateTime === 'string' ? Date.parse(expirationDateTime) : Number.NaN;
    const capped = nowMs + UPLOAD_SESSION_MAX_TTL_MS;
    const expiresAtMs = Number.isFinite(graphExpiry) ? Math.min(capped, graphExpiry) : capped;
    if (expiresAtMs <= nowMs) {
      throw new UploadSessionHandleError(
        'Graph returned an upload session that has already expired; create a new upload session.'
      );
    }
    if (this.#entries.size >= MAX_LIVE_HANDLES) {
      throw new UploadSessionHandleError(
        `No upload-session handle slots available (limit ${MAX_LIVE_HANDLES}); retry once outstanding handles are used or expire.`
      );
    }
    const id = randomBytes(HANDLE_BYTES).toString('base64url');
    const timer = setTimeout(() => this.forget(id), expiresAtMs - nowMs);
    timer.unref?.();
    this.#entries.set(id, {
      uploadUrl,
      expirationDateTime: Number.isFinite(graphExpiry)
        ? (expirationDateTime as string)
        : new Date(expiresAtMs).toISOString(),
      expiresAtMs,
      timer,
    });
    return { id, expiresAtMs };
  }

  /** Return the real URL and burn the handle, or undefined for unknown/used/expired. */
  redeem(id: string, nowMs: number = Date.now()): RedeemedUploadSession | undefined {
    this.sweep(nowMs);
    const entry = this.#entries.get(id);
    if (!entry) return undefined;
    this.forget(id);
    return { uploadUrl: entry.uploadUrl, expirationDateTime: entry.expirationDateTime };
  }

  /** Live handle count after sweeping, for tests and diagnostics. */
  size(nowMs: number = Date.now()): number {
    this.sweep(nowMs);
    return this.#entries.size;
  }

  /** Live handle count without sweeping: shows what the expiry timers alone removed. */
  liveCountUnswept(): number {
    return this.#entries.size;
  }

  clear(): void {
    for (const id of [...this.#entries.keys()]) this.forget(id);
  }

  toJSON(): Record<string, never> {
    return {};
  }

  toString(): string {
    return '[UploadSessionHandleStore]';
  }
}

/** `http://m365-mcp:3001/upload-session/<id>` from the configured attachment URL base. */
export function buildUploadSessionHandleUrl(base: string, id: string): string {
  return new URL(`${UPLOAD_SESSION_ROUTE}/${id}`, base).toString();
}

export type UploadSessionSwapResult =
  | { ok: true; body: unknown; realUrl?: string }
  | { ok: false; error: string };

/**
 * Replace `uploadUrl` in a create-upload-session result with a handle.
 * Every other field is kept; `uploadUrlIsHandle: true` is added. Fails closed: if the
 * body cannot be parsed, or no handle can be minted, the caller gets an error and never
 * the original text (which would carry the real URL).
 */
export function swapUploadUrlForHandle(
  text: string,
  store: UploadSessionHandleStore,
  base: string,
  nowMs: number = Date.now()
): UploadSessionSwapResult {
  let body: unknown;
  try {
    body = JSON.parse(text);
  } catch {
    return {
      ok: false,
      error: 'create-upload-session returned a body that could not be parsed; no handle issued.',
    };
  }
  if (typeof body !== 'object' || body === null || Array.isArray(body)) return { ok: true, body };
  const record = body as Record<string, unknown>;
  if (typeof record.uploadUrl !== 'string') return { ok: true, body };
  const realUrl = record.uploadUrl;
  try {
    const expiration =
      typeof record.expirationDateTime === 'string' ? record.expirationDateTime : undefined;
    const { id } = store.mint(realUrl, expiration, nowMs);
    return {
      ok: true,
      realUrl,
      body: {
        ...record,
        uploadUrl: buildUploadSessionHandleUrl(base, id),
        uploadUrlIsHandle: true,
      },
    };
  } catch (error) {
    const message =
      error instanceof UploadSessionHandleError
        ? error.message
        : 'Could not issue an upload-session handle.';
    return { ok: false, error: message };
  }
}

export interface UploadSessionRouteDeps {
  store: UploadSessionHandleStore;
  /** Clock override for tests. */
  now?: () => number;
}

function refuse(res: Response): void {
  res.status(404).type('text/plain').send(NOT_FOUND_BODY);
}

/**
 * `GET /upload-session/:id`: 200 `{uploadUrl, expirationDateTime}` once, then 404.
 * Logs nothing: neither the id (the capability) nor the URL (a credential).
 */
export function createUploadSessionHandler(deps: UploadSessionRouteDeps): Handler {
  return async (req: Request, res: Response): Promise<void> => {
    // Express routes HEAD to GET handlers; a probe must not burn the handle.
    if (req.method !== 'GET') {
      res.setHeader('allow', 'GET');
      res.status(405).type('text/plain').send('Method not allowed');
      return;
    }
    const id = (req.params as Record<string, unknown> | undefined)?.id;
    if (typeof id !== 'string' || !HANDLE_ID_RE.test(id)) {
      refuse(res);
      return;
    }
    const redeemed = deps.store.redeem(id, deps.now ? deps.now() : Date.now());
    if (!redeemed) {
      refuse(res);
      return;
    }
    res.setHeader('cache-control', 'no-store');
    res.setHeader('x-content-type-options', 'nosniff');
    res.status(200).type('application/json').send(JSON.stringify(redeemed));
  };
}
