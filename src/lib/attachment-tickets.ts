/**
 * Short-TTL, single-redemption tickets for Graph byte resources that expose no
 * pre-authenticated download URL of their own.
 *
 * A ticket is a capability and nothing else: it names one Graph path and one
 * account, it is redeemable once, and it expires. It carries no credential --
 * the redemption route authenticates to Graph with the server's own token, the
 * same way every tool does. Holding a ticket therefore authorises exactly one
 * authenticated GET of exactly one resource, which is the smallest grant that
 * makes an out-of-band fetch possible at all.
 *
 * HTK exception, opt-in via `--mint-with-request-identity`: a ticket minted while
 * Graph identity comes from the request carries that caller's access token, so
 * redemption reads as the caller rather than as the server. The token sits only in
 * this store's memory, is never part of the URL, is non-enumerable on the redeemed
 * ticket, and is dropped with the ticket on redemption or expiry.
 *
 * Memory-only, deliberately. Persisting tickets would mean a redeemable
 * capability surviving a restart, and re-reading it from disk is a second place
 * for it to leak; a ticket outliving the process it was minted in has no
 * legitimate use when the TTL is measured in minutes.
 */

import { randomBytes } from 'node:crypto';
import { ATTACHMENT_ROUTE, type AttachmentUrlConfig } from './attachment-url-config.js';
import { signUrl } from './url-signing.js';

/** Query parameter carrying the ticket id. */
export const TICKET_PARAM = 't';

/**
 * Build the signed, redeemable URL for a minted ticket.
 *
 * **The ticket travels in the query, never the path**, and that is a hard
 * requirement of the verifying sidecar rather than a style choice: docglean's
 * error messages keep a fetched URL's path (so an operator can tell which
 * document failed) and strip its query. A ticket in the path would be signed
 * just as correctly and would also land in every one of those messages.
 */
export function buildAttachmentUrl(
  config: AttachmentUrlConfig,
  ticketId: string,
  nowMs: number = Date.now()
): string {
  const url = new URL(ATTACHMENT_ROUTE, config.base);
  url.searchParams.set(TICKET_PARAM, ticketId);
  return signUrl(
    url.toString(),
    { key: config.key, keyId: config.keyId, ttlSeconds: config.ttlSeconds },
    nowMs
  );
}

/** Any origin and version prefix will do; only whether parsing changes the path matters. */
const PROBE_ORIGIN = 'https://graph.invalid';
const PROBE_PREFIX = '/v1.0';

/**
 * Whether a minted target survives URL parsing unchanged.
 *
 * The mint-site checks match on a suffix, which says nothing about what precedes it, and
 * `performRequest` later concatenates the target onto the Graph origin and lets WHATWG
 * parse the result. Anything the parser rewrites on the way makes the path that was
 * validated and the path that gets fetched two different things.
 *
 * This asks that question directly rather than enumerating the ways it can happen, because
 * enumerating them kept coming up short: dot segments resolve away, a fragment never
 * reaches the wire, and TAB, LF and CR are *deleted before* dot segments resolve, so a `..`
 * split by one of them survives a segment comparison and is reassembled by the parser.
 * Round-tripping catches that class whole, including the encoded and control-character
 * spellings, and stays correct if the parser grows another normalisation.
 *
 * `//` is rejected separately. With the version prefix in front it stays an ordinary path
 * segment, so the round trip accepts it, but it is only inert for as long as the caller
 * keeps prefixing something -- without that it parses as an authority and points off-origin.
 */
export function isPlainGraphPath(target: string): boolean {
  if (!target.startsWith('/')) return false;
  if (target.includes('//')) return false;
  let resolved: URL;
  try {
    resolved = new URL(PROBE_ORIGIN + PROBE_PREFIX + target);
  } catch {
    return false;
  }
  return (
    resolved.origin === PROBE_ORIGIN &&
    resolved.search === '' &&
    resolved.hash === '' &&
    resolved.pathname === PROBE_PREFIX + target
  );
}

export interface AttachmentTicket {
  /** Relative Graph path, exactly as the minting tool validated it. */
  readonly target: string;
  /** Account this ticket was minted for; undefined in single-account mode. */
  readonly accountName: string | undefined;
  /** Epoch milliseconds after which this ticket is dead. */
  readonly expiresAtMs: number;
  /**
   * HTK (`--mint-with-request-identity`): the caller's own Graph access token when the
   * ticket was minted under request identity. Redemption must fetch with exactly this
   * token and never fall back to the server's token cache.
   *
   * Present on a redeemed ticket as a **non-enumerable** property only, so JSON,
   * `util.inspect` and object spreads of a ticket never carry it. The store keeps it in
   * a separate map that is cleared together with the ticket (redeem or expiry).
   */
  readonly boundAccessToken?: string;
  /**
   * HTK: explicit Accept header for the redemption fetch. Graph refuses transcript
   * `/content` with fetch's default wildcard Accept ("Invalid format ... specified."),
   * so transcript tickets carry `text/vtt`. Unset for everything else (unchanged).
   */
  readonly accept?: string;
}

export interface MintOptions {
  /** Bind the caller's access token to this ticket (request-identity minting). */
  boundAccessToken?: string;
  /** Accept header to send when the ticket is redeemed. */
  accept?: string;
}

/**
 * Cap on live tickets. A ticket is ~200 bytes, so this bounds the store at a
 * few hundred KB -- but the reason for a cap is not memory, it is that an agent
 * in a retry loop should hit a refusal it can report rather than grow the
 * process without limit. Minting refuses when full, after sweeping; it never
 * evicts a live ticket, because evicting the oldest would let a caller minting
 * in a loop invalidate tickets someone else is about to redeem.
 */
const MAX_LIVE_TICKETS = 256;

/** 32 bytes of CSPRNG output -- the ticket id is the whole capability. */
const TICKET_BYTES = 32;

export class TicketStoreFullError extends Error {
  constructor(public readonly limit: number) {
    super(`No ticket slots available (limit ${limit}); retry once outstanding tickets expire.`);
    this.name = 'TicketStoreFullError';
  }
}

export class AttachmentTicketStore {
  private readonly tickets = new Map<string, AttachmentTicket>();
  /** HTK: bound caller tokens, keyed by ticket id. Never logged, never serialized. */
  private readonly boundTokens = new Map<string, string>();

  constructor(private readonly ttlSeconds: number) {}

  /** Drop every expired ticket. Called before each mint and each redemption. */
  private sweep(nowMs: number): void {
    for (const [id, ticket] of this.tickets) {
      if (ticket.expiresAtMs <= nowMs) this.forget(id);
    }
  }

  /** Drop a ticket and any token bound to it, together. */
  private forget(id: string): void {
    this.tickets.delete(id);
    this.boundTokens.delete(id);
  }

  mint(
    target: string,
    accountName: string | undefined,
    nowMs: number = Date.now(),
    options: MintOptions = {}
  ): { id: string; expiresAtMs: number } {
    this.sweep(nowMs);
    if (this.tickets.size >= MAX_LIVE_TICKETS) {
      throw new TicketStoreFullError(MAX_LIVE_TICKETS);
    }
    const id = randomBytes(TICKET_BYTES).toString('base64url');
    const expiresAtMs = nowMs + this.ttlSeconds * 1000;
    this.tickets.set(id, {
      target,
      accountName,
      expiresAtMs,
      ...(options.accept ? { accept: options.accept } : {}),
    });
    if (options.boundAccessToken) this.boundTokens.set(id, options.boundAccessToken);
    return { id, expiresAtMs };
  }

  /**
   * Return the ticket and burn it, or undefined.
   *
   * One `undefined` for every failure -- unknown id, already redeemed, expired.
   * The caller answers 404 to all three, so a probe cannot use the response to
   * tell "never existed" from "already used", which would confirm a guessed id.
   *
   * The delete happens before the value is returned rather than in the caller,
   * so an exception on the streaming path cannot leave a redeemed ticket live.
   */
  redeem(id: string, nowMs: number = Date.now()): AttachmentTicket | undefined {
    this.sweep(nowMs);
    const stored = this.tickets.get(id);
    if (!stored) return undefined;
    const boundAccessToken = this.boundTokens.get(id);
    this.forget(id);
    const ticket: AttachmentTicket = { ...stored };
    if (boundAccessToken !== undefined) {
      Object.defineProperty(ticket, 'boundAccessToken', {
        value: boundAccessToken,
        enumerable: false,
        writable: false,
        configurable: false,
      });
    }
    // No second expiry check here: `sweep` above ran against this same `nowMs`
    // and already removed anything at or past its expiry, so a surviving entry
    // is live by construction. A re-check would be unreachable code asserting a
    // guarantee the sweep already provides -- and because both use one captured
    // timestamp, there is no sweep/get race for it to cover.
    return ticket;
  }

  /** Live ticket count, for tests and diagnostics. Never logged with ids. */
  size(nowMs: number = Date.now()): number {
    this.sweep(nowMs);
    return this.tickets.size;
  }

  /** HTK: number of live tickets carrying a bound caller token, for tests. */
  boundTokenCount(nowMs: number = Date.now()): number {
    this.sweep(nowMs);
    return this.boundTokens.size;
  }

  clear(): void {
    this.tickets.clear();
    this.boundTokens.clear();
  }
}
