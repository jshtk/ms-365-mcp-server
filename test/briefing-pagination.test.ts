import { afterEach, describe, expect, it, vi } from 'vitest';
import { registerGraphTools } from '../src/graph-tools.js';

vi.mock('../src/logger.js', () => ({ default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() } }));
vi.mock('../src/generated/client-beta.js', () => ({ api: { endpoints: [] } }));
// A collection endpoint carries OData query params; since v0.156 upstream only
// offers skiptoken on such endpoints (isSkiptokenApplicable).
vi.mock('../src/generated/client.js', async () => {
  const { z } = await import('zod');
  return { api: { endpoints: [{
    alias: 'list-mail-messages', method: 'get', path: '/me/messages',
    parameters: [{ name: 'top', type: 'Query', schema: z.number().optional() }],
  }] } };
});

const page = (body: object) => ({ content: [{ type: 'text', text: JSON.stringify(body) }] });
function setup(responses: unknown[]) {
  const server = { registerTool: vi.fn(), tool: vi.fn() };
  const request = vi.fn();
  for (const response of responses) request.mockResolvedValueOnce(response);
  registerGraphTools(server as any, { graphRequest: request, serialize: JSON.stringify } as any, false);
  const registration = server.registerTool.mock.calls.find(c => c[0] === 'list-mail-messages')!;
  return { request, handler: registration.at(-1) as (args: object) => Promise<any>, schema: registration[1].inputSchema };
}
afterEach(() => vi.unstubAllEnvs());
describe('briefing continuation', () => {
  it('retains the next link and original total when a page cap is reached', async () => {
    vi.stubEnv('MS365_MCP_MAX_PAGES', '1');
    const link = 'https://graph.microsoft.com/v1.0/me/messages?$skiptoken=opaque';
    const { handler, request } = setup([page({ value: [{ id: 'a' }], '@odata.nextLink': link, '@odata.count': 30 })]);
    const result = JSON.parse((await handler({ fetchAllPages: true })).content[0].text);
    expect(result['@odata.nextLink']).toBe(link);
    expect(result['@odata.count']).toBe(30);
    expect(request).toHaveBeenCalledTimes(1);
  });
  it('forwards a continuation token as a query value without changing the resource', async () => {
    const { handler, request, schema } = setup([page({ value: [] })]);
    expect(schema.shape.skiptoken).toBeDefined();
    await handler({ skiptoken: 'opaque&$filter=anything' });
    expect(request.mock.calls[0][0]).toBe('/me/messages?$skiptoken=opaque%26%24filter%3Danything');
  });
  it('removes the continuation only after the last page', async () => {
    const { handler } = setup([
      page({ value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=1' }),
      page({ value: [{ id: 'b' }] }),
    ]);
    const result = JSON.parse((await handler({ fetchAllPages: true })).content[0].text);
    expect(result.value).toHaveLength(2);
    expect(result['@odata.nextLink']).toBeUndefined();
  });
  it('does not present a failed second page as a successful collection', async () => {
    const { handler } = setup([
      page({ value: [{ id: 'a' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/messages?$skip=1' }),
      { isError: true, content: [{ type: 'text', text: '{"error":"unavailable"}' }] },
    ]);
    expect((await handler({ fetchAllPages: true })).isError).toBe(true);
  });
});
