// Structured one-line JSON logs. Callers must never pass email addresses, subjects or bodies.
export function createLogger(write = (line) => process.stdout.write(`${line}\n`)) {
  return (level, event, fields = {}) => write(JSON.stringify({ t: new Date().toISOString(), level, event, ...fields }));
}
