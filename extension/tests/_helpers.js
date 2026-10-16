export const goodSettings = () => ({
  apiBaseUrl: 'https://api.example.com/',
  apiKey: 'secret-key-123',
});
export const goodJob = (over = {}) => ({
  id: 'job_1', to: 'a@example.com', subject: 'Hello', body: 'Line 1\nLine 2', ...over,
});
export function fakeResponse(status, body, headers = {}) {
  const text = body === undefined ? '' : typeof body === 'string' ? body : JSON.stringify(body);
  return {
    ok: status >= 200 && status < 300, status,
    headers: { get: (k) => headers[k.toLowerCase()] ?? null },
    text: async () => text,
  };
}
