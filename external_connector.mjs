import net from "node:net";

const MAX_RESPONSE_BYTES = 512 * 1024;
const DEFAULT_TIMEOUT_MS = 15_000;
const SECRET_KEYS = /token|secret|password|authorization|api[-_]?key|cookie/i;

function isPrivateHost(hostname) {
  const host = String(hostname || "").toLowerCase().replace(/\.$/, "");
  if (["localhost", "localhost.localdomain"].includes(host)) return true;
  const ipVersion = net.isIP(host);
  if (!ipVersion) return false;
  if (ipVersion === 4) {
    const [a, b] = host.split(".").map(Number);
    return a === 10 || a === 127 || (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168) || a === 0 || a >= 224;
  }
  const normalized = host.replace(/^::ffff:/, "");
  if (net.isIP(normalized) === 4) return isPrivateHost(normalized);
  return host === "::" || host === "::1" || host.startsWith("fc") || host.startsWith("fd") || host.startsWith("fe8") || host.startsWith("fe9") || host.startsWith("fea") || host.startsWith("feb");
}

function sanitizeValue(value, depth = 0) {
  if (depth > 6) return "[depth limited]";
  if (typeof value === "string") {
    return value.replace(/([?&](?:token|secret|password|api[_-]?key|authorization)=[^&#\s]*)/gi, "$1".replace(/=[^&#\s]*/, "=[REDACTED]"));
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

export class ExternalConnector {
  constructor({ enabled = false, allowedOrigins = [], credentialResolver = null, timeoutMs = DEFAULT_TIMEOUT_MS } = {}) {
    this.enabled = enabled === true;
    this.allowedOrigins = new Set(allowedOrigins);
    this.credentialResolver = credentialResolver;
    this.timeoutMs = Math.max(1000, Math.min(Number(timeoutMs) || DEFAULT_TIMEOUT_MS, 60_000));
  }

  async request(input, { originApproved = false, credentialApproved = false } = {}) {
    const request = normalizeConnectorRequest(input);
    if (!this.enabled) throw new Error("External connector is disabled.");
    if (!originApproved || !this.allowedOrigins.has(request.origin)) throw new Error("External origin approval is required.");
    if (request.credential_scope && !credentialApproved) throw new Error("Separate credential approval is required.");
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    try {
      const headers = { Accept: "application/json, text/plain;q=0.9" };
      if (request.credential_scope) {
        if (typeof this.credentialResolver !== "function") throw new Error("Credential resolver is unavailable.");
        const token = await this.credentialResolver(request.credential_scope);
        if (!token || typeof token !== "string") throw new Error("Credential is unavailable.");
        headers.ZakupayToken = token;
      }
      const response = await fetch(request.url, { method: "GET", redirect: "manual", headers, signal: controller.signal });
      const body = await response.text();
      if (Buffer.byteLength(body, "utf8") > MAX_RESPONSE_BYTES) throw new Error("Connector response is too large.");
      let data;
      try { data = JSON.parse(body); } catch { data = body.slice(0, MAX_RESPONSE_BYTES); }
      return { ok: response.ok, status: response.status, content_type: response.headers.get("content-type") || "", data: sanitizeValue(data) };
    } finally {
      clearTimeout(timer);
    }
  }
}

export { isPrivateHost, sanitizeValue };
