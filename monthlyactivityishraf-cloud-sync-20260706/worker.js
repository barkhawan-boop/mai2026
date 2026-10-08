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

function parseReportPeriod(url) {
  const activityYear = Number(url.searchParams.get("year"));
  const activityMonth = Number(url.searchParams.get("month"));
  if (!Number.isInteger(activityYear) || activityYear < 2000 || activityYear > 2100 ||
      !Number.isInteger(activityMonth) || activityMonth < 1 || activityMonth > 12) return null;
  return { activityYear, activityMonth };
}

function parseReportAction(body) {
  const inspectorId = Number(body?.inspectorId);
  const activityYear = Number(body?.year);
  const activityMonth = Number(body?.month);
  if (!Number.isSafeInteger(inspectorId) || inspectorId < 1 ||
      !Number.isInteger(activityYear) || activityYear < 2000 || activityYear > 2100 ||
      !Number.isInteger(activityMonth) || activityMonth < 1 || activityMonth > 12) return null;
  return { inspectorId, activityYear, activityMonth };
}

async function handleReportStatusRequest(request, env, url) {
  if (!env.DB) return jsonResponse({ error: "Missing D1 database binding" }, { status: 500 });

  if (request.method === "GET") {
    const period = parseReportPeriod(url);
    if (!period) return jsonResponse({ error: "Invalid report period" }, { status: 400 });
    const result = await env.DB.prepare(
      "SELECT inspector_id AS inspectorId, sent_at AS sentAt, printed_at AS printedAt FROM report_status WHERE activity_year = ?1 AND activity_month = ?2 AND (printed_at IS NULL OR printed_at > datetime('now', '-7 days'))",
    ).bind(period.activityYear, period.activityMonth).all();
    return jsonResponse({ reports: result.results || [] });
  }

  if (request.method !== "POST") return jsonResponse({ error: "Method not allowed" }, { status: 405 });

  let action;
  try {
    action = parseReportAction(await request.json());
  } catch {
    return jsonResponse({ error: "Invalid JSON" }, { status: 400 });
  }
  if (!action) return jsonResponse({ error: "Invalid report status request" }, { status: 400 });

  if (url.pathname === "/api/report-status/send") {
    await env.DB.prepare(
      `INSERT INTO report_status (inspector_id, activity_year, activity_month, sent_at)
       VALUES (?1, ?2, ?3, CURRENT_TIMESTAMP)
       ON CONFLICT(inspector_id, activity_year, activity_month) DO UPDATE SET sent_at = CURRENT_TIMESTAMP`,
    ).bind(action.inspectorId, action.activityYear, action.activityMonth).run();
    return jsonResponse({ ok: true });
  }

  if (url.pathname === "/api/report-status/printed") {
    await env.DB.prepare(
      "UPDATE report_status SET printed_at = CURRENT_TIMESTAMP WHERE inspector_id = ?1 AND activity_year = ?2 AND activity_month = ?3",
    ).bind(action.inspectorId, action.activityYear, action.activityMonth).run();
    return jsonResponse({ ok: true });
  }

  return jsonResponse({ error: "Not found" }, { status: 404 });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === "/boss") {
      if (!env.ASSETS) return new Response("Not found", { status: 404 });
      return env.ASSETS.fetch(new Request(new URL("/", url), request));
    }
    if (url.pathname === "/api/report-status" || url.pathname.startsWith("/api/report-status/")) {
      return handleReportStatusRequest(request, env, url);
    }
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
