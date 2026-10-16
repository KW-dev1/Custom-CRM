// In-memory stand-in for a chrome.storage area.
export function memoryArea(initial = {}) {
  const data = structuredClone(initial);
  return {
    data,
    async get(k) {
      if (typeof k === 'string') return k in data ? { [k]: structuredClone(data[k]) } : {};
      return structuredClone(data);
    },
    async set(o) { Object.assign(data, structuredClone(o)); },
    async remove(k) { for (const x of [].concat(k)) delete data[x]; },
  };
}
