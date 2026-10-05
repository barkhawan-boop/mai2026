let cachedAccessKeys = null;
let cachedAccessKeysExpiry = 0;

function jsonResponse(value, init = {}) {
  return new Response(JSON.stringify(value), {
    ...init,
    headers: {
      "content-type": "application/json; charset=utf-8",
      "cache-control": "no-store",
      ...(init.headers || {}),
    },
  });
}

function base64UrlToBytes(value) {
  const base64 = value.replace(/-/g, "+").replace(/_/g, "/");
  const padded = base64 + "=".repeat((4 - (base64.length % 4)) % 4);
  const binary = atob(padded);
  return Uint8Array.from(binary, (character) => character.charCodeAt(0));
}

function decodeJwtPart(value) {
  return JSON.parse(new TextDecoder().decode(base64UrlToBytes(value)));
}

async function accessIdentity(request, env) {
  const teamDomain = env.ACCESS_TEAM_DOMAIN?.replace(/\/$/, "");
  const audience = env.ACCESS_AUD;
  const token = request.headers.get("cf-access-jwt-assertion");
  if (!teamDomain || !audience || !token) return null;

  const [encodedHeader, encodedPayload, encodedSignature] = token.split(".");
  if (!encodedHeader || !encodedPayload || !encodedSignature) return null;

  let header;
  let payload;
  try {
    header = decodeJwtPart(encodedHeader);
    payload = decodeJwtPart(encodedPayload);
  } catch {
    return null;
  }

  const now = Math.floor(Date.now() / 1000);
  const audiences = Array.isArray(payload.aud) ? payload.aud : [payload.aud];
  if (
    header.alg !== "RS256" ||
    !header.kid ||
    payload.iss !== teamDomain ||
    !audiences.includes(audience) ||
    typeof payload.exp !== "number" ||
    payload.exp <= now ||
    (typeof payload.nbf === "number" && payload.nbf > now) ||
    typeof payload.email !== "string" ||
    !payload.email.trim()
  ) {
    return null;
  }

  if (!cachedAccessKeys || cachedAccessKeysExpiry <= now) {
    try {
      const response = await fetch(`${teamDomain}/cdn-cgi/access/certs`, {
        cf: { cacheTtl: 300, cacheEverything: true },
      });
      if (!response.ok) return null;
      const keySet = await response.json();
      cachedAccessKeys = Array.isArray(keySet.keys) ? keySet.keys : [];
      cachedAccessKeysExpiry = now + 300;
    } catch {
      return null;
    }
  }

  let jwk = cachedAccessKeys.find((key) => key.kid === header.kid);
  if (!jwk) return null;

  try {
    const key = await crypto.subtle.importKey(
      "jwk",
      jwk,
      { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
      false,
      ["verify"],
    );
    const signedContent = new TextEncoder().encode(`${encodedHeader}.${encodedPayload}`);
    const valid = await crypto.subtle.verify(
      "RSASSA-PKCS1-v1_5",
      key,
      base64UrlToBytes(encodedSignature),
      signedContent,
    );
    return valid ? payload : null;
  } catch {
    return null;
  }
}

async function userStorageKey(email) {
  const normalizedEmail = email.trim().toLowerCase();
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(normalizedEmail));
  const hash = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
  return `user:${hash}`;
}

async function handleDataRequest(request, env, userKey) {
  if (!env.DB) return jsonResponse({ error: "Missing D1 database binding" }, { status: 500 });

  if (request.method === "GET") {
    const row = await env.DB.prepare("SELECT payload FROM user_data WHERE user_id = ?1")
      .bind(userKey)
      .first();
    return new Response(row?.payload || "null", {
      headers: {
        "content-type": "application/json; charset=utf-8",
        "cache-control": "no-store",
      },
    });
  }

  if (request.method === "POST") {
    const body = await request.text();
    if (new TextEncoder().encode(body).byteLength > 1_000_000) {
      return jsonResponse({ error: "Data is too large" }, { status: 413 });
    }

    try {
      const parsed = JSON.parse(body);
      if (!parsed || !Array.isArray(parsed.inspectors) || !parsed.global || typeof parsed.global !== "object") {
        return jsonResponse({ error: "Invalid data shape" }, { status: 400 });
      }
    } catch {
      return jsonResponse({ error: "Invalid JSON" }, { status: 400 });
    }

    await env.DB.prepare(
      `INSERT INTO user_data (user_id, payload, updated_at)
       VALUES (?1, ?2, CURRENT_TIMESTAMP)
       ON CONFLICT(user_id) DO UPDATE SET
         payload = excluded.payload,
         updated_at = CURRENT_TIMESTAMP`,
    )
      .bind(userKey, body)
      .run();
    return jsonResponse({ ok: true });
  }

  return jsonResponse({ error: "Method not allowed" }, { status: 405 });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === "/api/data") {
      const identity = await accessIdentity(request, env);
      if (!identity?.email) return jsonResponse({ error: "Sign in required" }, { status: 401 });
      if (request.method === "GET" && url.searchParams.has("intent")) {
        const intent = url.searchParams.get("intent") === "save" ? "save" : "load";
        return Response.redirect(`${url.origin}/?cloud=${intent}`, 302);
      }
      return handleDataRequest(request, env, await userStorageKey(identity.email));
    }

    if (env.ASSETS) return env.ASSETS.fetch(request);
    return new Response("Not found", { status: 404 });
  },
};
