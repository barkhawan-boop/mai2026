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

async function handleDataRequest(request, env) {
  if (!env.DB) return jsonResponse({ error: "Missing D1 database binding" }, { status: 500 });

  if (request.method === "GET") {
    const row = await env.DB.prepare("SELECT payload FROM user_data WHERE user_id = 'public'").first();
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
       VALUES ('public', ?1, CURRENT_TIMESTAMP)
       ON CONFLICT(user_id) DO UPDATE SET
         payload = excluded.payload,
         updated_at = CURRENT_TIMESTAMP`,
    )
      .bind(body)
      .run();
    return jsonResponse({ ok: true });
  }

  return jsonResponse({ error: "Method not allowed" }, { status: 405 });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === "/api/data") {
      if (request.method === "GET" && url.searchParams.has("intent")) {
        const intent = url.searchParams.get("intent") === "save" ? "save" : "load";
        return Response.redirect(`${url.origin}/?cloud=${intent}`, 302);
      }
      return handleDataRequest(request, env);
    }

    if (env.ASSETS) return env.ASSETS.fetch(request);
    return new Response("Not found", { status: 404 });
  },
};
