import dns from "node:dns/promises";
import https from "node:https";
import net from "node:net";

const MAX_RESPONSE_BYTES = 512 * 1024;
const DEFAULT_TIMEOUT_MS = 15_000;
const SECRET_KEYS = /token|access[_-]?token|refresh[_-]?token|secret|client[_-]?secret|password|authorization|cookie|api[-_]?key/i;

function isPrivateHost(hostname) {
  const host = String(hostname || "").toLowerCase().replace(/\.$/, "");
  if (["localhost", "localhost.localdomain"].includes(host)) return true;
  const ipVersion = net.isIP(host);
  if (!ipVersion) return false;
  if (ipVersion === 4) {
    const octets = host.split(".").map(Number);
    const n = (((octets[0] * 256 + octets[1]) * 256 + octets[2]) * 256 + octets[3]) >>> 0;
    const inRange = (start, end) => n >= start && n <= end;
    return inRange(0x00000000, 0x00ffffff) || inRange(0x0a000000, 0x0affffff) || inRange(0x64400000, 0x647fffff) || inRange(0x7f000000, 0x7fffffff) || inRange(0xa9fe0000, 0xa9feffff) || inRange(0xac100000, 0xac1fffff) || inRange(0xc0000000, 0xc00000ff) || inRange(0xc0a80000, 0xc0a8ffff) || inRange(0xc6120000, 0xc613ffff) || inRange(0xc6336400, 0xc63364ff) || inRange(0xcb007100, 0xcb0071ff) || n >= 0xe0000000;
  }
  const normalized = host.replace(/^::ffff:/, "");
  if (net.isIP(normalized) === 4) return isPrivateHost(normalized);
  return host === "::" || host === "::1" || host.startsWith("fc") || host.startsWith("fd") || host.startsWith("fe8") || host.startsWith("fe9") || host.startsWith("fea") || host.startsWith("feb");
}

async function assertPublicResolution(hostname) {
  if (isPrivateHost(hostname)) throw new Error("Connector refuses private or local hosts.");
  const records = await dns.lookup(hostname, { all: true, verbatim: true });
  if (!records.length || records.some((record) => isPrivateHost(record.address))) throw new Error("Connector resolved to a private or local host.");
}

function sanitizeValue(value, depth = 0) {
  if (depth > 6) return "[depth limited]";
  if (typeof value === "string") {
    return value
      .replace(/([?&#\s](?:token|access[_-]?token|refresh[_-]?token|secret|client[_-]?secret|password|api[_-]?key)\s*[=:]\s*)[^&#\s,;]*/gi, "$1[REDACTED]")
      .replace(/(authorization\s*:\s*bearer\s+)[^\s,;]+/gi, "$1[REDACTED]")
      .replace(/(^|[\s,{])((?:token|access[_-]?token|refresh[_-]?token|secret|client[_-]?secret|password|api[_-]?key)\s*[=:]\s*)[^\s,;}]+/gi, "$1$2[REDACTED]");
  }
  if (Array.isArray(value)) return value.slice(0, 500).map((item) => sanitizeValue(item, depth + 1));
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.entries(value).slice(0, 500).map(([key, item]) => [key, SECRET_KEYS.test(key) ? "[REDACTED]" : sanitizeValue(item, depth + 1)]));
  }
  return value;
}

export function normalizeConnectorRequest(input = {}) {
  const parsed = new URL(String(input.url || ""));
  if (parsed.protocol !== "https:") throw new Error("Connector requires HTTPS.");
  if (parsed.username || parsed.password || parsed.search || parsed.hash) throw new Error("Connector URLs must not contain credentials, query secrets, or fragments.");
  if (isPrivateHost(parsed.hostname)) throw new Error("Connector refuses private or local hosts.");
  const operation = String(input.operation || "api_read").toLowerCase();
  if (!["read", "browser_read", "api_read"].includes(operation)) throw new Error("Connector permits read-only operations.");
  const credentialScope = String(input.credential_scope || "").trim();
  if (credentialScope.length > 160) throw new Error("Credential scope is too long.");
  return { url: parsed.toString(), origin: parsed.origin, operation, purpose: String(input.purpose || "").trim().slice(0, 240), credential_scope: credentialScope };
}

function requestHttps(url, headers, timeoutMs) {
  const parsed = new URL(url);
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (error, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(deadline);
      if (error) reject(error); else resolve(value);
    };
    const deadline = setTimeout(() => { req.destroy(); finish(new Error("Connector request timed out.")); }, timeoutMs);
    const req = https.request(parsed, {
      method: "GET", headers, timeout: timeoutMs, servername: parsed.hostname,
      lookup: (hostname, _options, callback) => dns.lookup(hostname, { all: true, verbatim: true }).then((records) => {
        const publicRecord = records.find((record) => !isPrivateHost(record.address));
        if (!publicRecord || records.some((record) => isPrivateHost(record.address))) throw new Error("Connector resolved to a private or local host.");
        callback(null, publicRecord.address, publicRecord.family);
      }).catch((error) => callback(error)),
    }, (response) => {
      const chunks = [];
      let total = 0;
      response.on("data", (chunk) => {
        total += chunk.length;
        if (total > MAX_RESPONSE_BYTES) { req.destroy(); finish(new Error("Connector response is too large.")); return; }
        chunks.push(chunk);
      });
      response.on("error", (error) => finish(error));
      response.on("aborted", () => finish(new Error("Connector response was aborted.")));
      response.on("end", () => finish(null, { ok: response.statusCode >= 200 && response.statusCode < 300, status: response.statusCode || 0, content_type: response.headers["content-type"] || "", body: Buffer.concat(chunks).toString("utf8") }));
    });
    req.on("timeout", () => req.destroy(new Error("Connector request timed out.")));
    req.on("error", (error) => finish(error));
    req.end();
  });
}

export class ExternalConnector {
  constructor({ enabled = false, allowedOrigins = [], credentialResolver = null, timeoutMs = DEFAULT_TIMEOUT_MS, requester = requestHttps } = {}) {
    this.enabled = enabled === true;
    this.allowedOrigins = new Set(allowedOrigins);
    this.credentialResolver = credentialResolver;
    this.requester = requester;
    this.timeoutMs = Math.max(1000, Math.min(Number(timeoutMs) || DEFAULT_TIMEOUT_MS, 60_000));
  }

  async request(input, { originApproved = false, credentialApproved = false } = {}) {
    const request = normalizeConnectorRequest(input);
    if (!this.enabled) throw new Error("External connector is disabled.");
    await assertPublicResolution(new URL(request.url).hostname);
    if (!originApproved || !this.allowedOrigins.has(request.origin)) throw new Error("External origin approval is required.");
    if (request.credential_scope && !credentialApproved) throw new Error("Separate credential approval is required.");
    const headers = { Accept: "application/json, text/plain;q=0.9" };
    if (request.credential_scope) {
      if (typeof this.credentialResolver !== "function") throw new Error("Credential resolver is unavailable.");
      const token = await this.credentialResolver(request.credential_scope);
      if (!token || typeof token !== "string") throw new Error("Credential is unavailable.");
      headers.ZakupayToken = token;
    }
    const response = await this.requester(request.url, headers, this.timeoutMs);
    let data;
    try { data = JSON.parse(response.body); } catch { data = String(response.body).slice(0, MAX_RESPONSE_BYTES); }
    return { ok: response.ok, status: response.status, content_type: response.content_type, data: sanitizeValue(data) };
  }
}

export { isPrivateHost, sanitizeValue };
