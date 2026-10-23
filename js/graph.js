/* Smart File - Microsoft Graph calls (mail folders, message history, move, create folder). */
(function (global) {
  "use strict";

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const FOLDER_SELECT = "$select=id,displayName,parentFolderId,childFolderCount,totalItemCount&$top=250";
  const MSG_SELECT = "$select=id,parentFolderId,conversationId,internetMessageId,subject";
  const BATCH_MAX = 20;       // Graph accepts at most 20 requests in one $batch
  const MAILBOX_SLOTS = 4;    // Exchange runs at most 4 requests per mailbox at a time; a $batch uses all 4
  const EQ_MAX = 120;         // longest value Graph accepts in an "eq" filter
  const RETRYABLE = { 429: true, 503: true, 504: true };
  const NOT_READABLE = { 400: true, 403: true, 404: true };   // this folder can't be read; asking again won't help

  class GraphError extends Error {
    constructor(status, code, message, body) {
      super(message || ("Microsoft Graph returned " + status));
      this.name = "GraphError";
      this.status = status;
      this.code = code || "";
      this.body = body;
    }
  }

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const q = (s) => s.replace(/'/g, "''"); // escape single quotes inside OData string literals

  /* ---------- concurrency gate ----------
   * Exchange answers anything beyond 4 concurrent requests per mailbox with 429, and every 429 costs a
   * multi-second wait. Every call goes through this gate: a single request takes 1 slot, a $batch takes all 4.
   * Waiters are served in order, so a 4-slot batch is never starved by a stream of single requests. A request
   * marked urgent (something the user is waiting on, like a move) takes free slots straight away and otherwise
   * goes to the front of the queue; while an urgent section is open, queued background work stays queued. */
  const gate = { free: MAILBOX_SLOTS, waiting: [] };
  let urgentDepth = 0;      // open urgent sections
  let urgentSleeping = 0;   // urgent requests currently waiting out a Retry-After; the slots are free meanwhile
  const FETCH_TIMEOUT_MS = 60 * 1000;
  const BATCH_TIMEOUT_MS = 120 * 1000;

  function acquire(n, urgent) {
    return new Promise((resolve) => {
      if (gate.free >= n && (urgent || !gate.waiting.length)) { gate.free -= n; resolve(); return; }
      const w = { n: n, urgent: !!urgent, resolve: resolve };
      if (urgent) gate.waiting.unshift(w); else gate.waiting.push(w);
    });
  }

  function release(n) {
    gate.free += n;
    const holding = urgentDepth > urgentSleeping;        // a user action is mid-flight (not just waiting out a 429)
    while (gate.waiting.length) {
      const w = gate.waiting[0];
      if (holding && !w.urgent) break;                   // hold background work back until the user's action is done
      if (gate.free < w.n) break;
      gate.waiting.shift();
      gate.free -= w.n;
      w.resolve();
    }
  }

  /** Run `fn` ahead of any queued background work. Every request `fn` makes must be urgent, or it could queue behind the hold. */
  function urgently(fn) {
    urgentDepth++;
    return Promise.resolve().then(fn).finally(() => { urgentDepth--; release(0); });
  }

  function fetchWithTimeout(url, init, ms) {
    if (typeof AbortController === "undefined") return fetch(url, init);
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), ms);
    return fetch(url, Object.assign({}, init, { signal: ctrl.signal })).finally(() => clearTimeout(timer));
  }

  function retryAfterMs(header) {
    const s = Number(header);
    return Math.min(10, s > 0 ? s : 2) * 1000;
  }

  function headerOf(headers, name) {
    if (!headers) return null;
    const key = Object.keys(headers).find((k) => k.toLowerCase() === name);
    return key ? headers[key] : null;
  }

  async function request(token, path, opts) {
    opts = opts || {};
    const method = opts.method || "GET";
    const slots = opts.slots || 1;
    const url = path.startsWith("http") ? path : GRAPH + path;
    const headers = Object.assign({ Authorization: "Bearer " + token, Accept: "application/json" }, opts.headers || {});
    if (opts.body) headers["Content-Type"] = "application/json";
    let retries = opts.retries == null ? 2 : opts.retries;

    for (;;) {
      await acquire(slots, opts.urgent);
      let res;
      try {
        res = await fetchWithTimeout(url, { method: method, headers: headers, body: opts.body ? JSON.stringify(opts.body) : undefined }, slots > 1 ? BATCH_TIMEOUT_MS : FETCH_TIMEOUT_MS);
      } catch (e) {
        release(slots);
        if (e && e.name === "AbortError") throw new GraphError(0, "timeout", "Microsoft 365 took too long to answer. Try again.");
        throw new GraphError(0, "network", "Can’t reach Microsoft Graph. Check your network connection.");
      }

      if (RETRYABLE[res.status] && retries > 0) {
        retries--;
        release(slots);                                    // don't hold the slot while waiting
        if (opts.urgent) { urgentSleeping++; release(0); } // and let queued work use it meanwhile
        try { await sleep(retryAfterMs(res.headers.get("Retry-After"))); }
        finally { if (opts.urgent) urgentSleeping--; }
        continue;
      }

      try {
        if (res.status === 204) return null;
        const text = await res.text();
        let data = null;
        if (text) { try { data = JSON.parse(text); } catch (e) { data = { raw: text }; } }
        if (!res.ok) {
          const err = (data && data.error) || {};
          throw new GraphError(res.status, err.code, err.message, data);
        }
        return data;
      } finally {
        release(slots);
      }
    }
  }

  const NO_ANSWER = { error: { code: "NoResponse", message: "Microsoft 365 did not answer part of the request. Try again." } };

  /**
   * POST /$batch. `requests` = [{ id, method?, url, body? }], at most 20.
   * Returns { id: { status, headers, body } }. The batch itself succeeding does not mean every item did:
   * check each item's status. Items that come back 429/503/504 are retried (twice) after their Retry-After.
   * Only the first round can throw; if a retry round fails outright, its items are reported with that failure
   * and everything that already succeeded is kept.
   */
  async function batch(token, requests) {
    if (requests.length > BATCH_MAX) throw new Error("A $batch can hold at most " + BATCH_MAX + " requests.");
    const out = {};
    let pending = requests.map((r) => {
      const item = { id: String(r.id), method: r.method || "GET", url: r.url };
      if (r.body) { item.body = r.body; item.headers = { "Content-Type": "application/json" }; }
      return item;
    });

    for (let attempt = 0; pending.length; attempt++) {
      let data;
      try {
        data = await request(token, "/$batch", { method: "POST", body: { requests: pending }, slots: MAILBOX_SLOTS });
      } catch (e) {
        if (attempt === 0 || e.status === 400 || e.status === 401 || e.status === 403) throw e;   // rejected outright: tell the caller
        // A retry round that died in transit: report the items as unanswered (status 0, so nobody mistakes it for "folder absent").
        pending.forEach((p) => { out[p.id] = { id: p.id, status: 0, body: { error: { code: e.code || "NoResponse", message: e.message, httpStatus: e.status } } }; });
        break;
      }
      const responses = (data && data.responses) || [];
      const retry = [];
      let wait = 0;
      responses.forEach((r) => {
        if (RETRYABLE[r.status] && attempt < 2) {
          wait = Math.max(wait, Number(headerOf(r.headers, "retry-after")) || 1);
          retry.push(String(r.id));
        } else {
          out[String(r.id)] = r;
        }
      });
      pending.forEach((p) => { if (!out[p.id] && !retry.includes(p.id)) out[p.id] = { id: p.id, status: 0, body: NO_ANSWER }; });
      pending = pending.filter((p) => retry.includes(p.id));
      if (pending.length) await sleep(Math.min(10, wait) * 1000);
    }
    return out;
  }

  async function listAll(token, path) {
    let out = [];
    let next = path;
    while (next) {
      const page = await request(token, next);
      out = out.concat(page.value || []);
      next = page["@odata.nextLink"];
    }
    return out;
  }

  function withPath(f, parent) {
    return Object.assign({}, f, {
      depth: parent ? parent.depth + 1 : 0,
      pathParts: parent ? parent.pathParts.concat([f.displayName]) : [f.displayName]
    });
  }

  /**
   * Walk the whole folder tree (breadth-first) and return { folders, complete, problems, gaps }.
   * Each level's child lookups go out 20 to a $batch, one batch at a time.
   * `problems` are folders whose children could not be read for a passing reason (throttling, outage):
   * `complete` is false and the caller should try again later. `gaps` are folders Graph refuses to open at all;
   * they are simply skipped and do not count against `complete`. Only the top-level listing (or a request that
   * is rejected outright, 400/401/403) throws; everything else degrades.
   */
  async function listAllFolders(token, maxDepth) {
    const out = [];
    const problems = [];
    const gaps = [];
    let level = (await listAll(token, "/me/mailFolders?" + FOLDER_SELECT)).map((f) => withPath(f, null));

    while (level.length) {
      out.push.apply(out, level);
      const parents = level.filter((f) => f.childFolderCount > 0 && f.depth < maxDepth);
      const next = [];
      for (let i = 0; i < parents.length; i += BATCH_MAX) {
        const slice = parents.slice(i, i + BATCH_MAX);
        let responses;
        try {
          responses = await batch(token, slice.map((p, idx) => ({
            id: String(idx), url: "/me/mailFolders/" + encodeURIComponent(p.id) + "/childFolders?" + FOLDER_SELECT
          })));
        } catch (e) {
          if (e.status === 400 || e.status === 401 || e.status === 403) throw e;   // not going to get better by itself
          slice.forEach((p) => problems.push(p.displayName));
          continue;
        }
        for (let idx = 0; idx < slice.length; idx++) {
          const p = slice[idx];
          const r = responses[String(idx)];
          if (!r || r.status !== 200 || !r.body) {
            (r && NOT_READABLE[r.status] ? gaps : problems).push(p.displayName);
            continue;
          }
          let children = r.body.value || [];
          if (r.body["@odata.nextLink"]) {                 // more than 250 subfolders: rare, fetch the rest
            try { children = children.concat(await listAll(token, r.body["@odata.nextLink"])); }
            catch (e) { problems.push(p.displayName); }
          }
          children.forEach((c) => next.push(withPath(c, p)));
        }
      }
      level = next;
    }

    const folders = out.map((f) => ({
      id: f.id,
      name: f.displayName,
      parentId: f.parentFolderId,
      depth: f.depth,
      pathParts: f.pathParts,
      path: f.pathParts.join(" / "),
      totalItemCount: f.totalItemCount || 0,
      childFolderCount: f.childFolderCount || 0
    }));
    return { folders: folders, complete: problems.length === 0, problems: problems, gaps: gaps };
  }

  /**
   * Resolve well-known folder names (inbox, sentitems, ...) to IDs. Returns { ids, complete, unresolved }.
   * A 404 (or 400/403) means the mailbox simply has no such folder for us; a passing failure lands the name in
   * `unresolved` and makes `complete` false, so the caller does not trust a map that is missing, say, Deleted Items.
   */
  async function wellKnownFolderIds(token, names) {
    const unique = Array.from(new Set(names.map((n) => String(n).toLowerCase())));
    const ids = {};
    const unresolved = [];
    for (let i = 0; i < unique.length; i += BATCH_MAX) {
      const slice = unique.slice(i, i + BATCH_MAX);
      const responses = await batch(token, slice.map((n) => ({ id: n, url: "/me/mailFolders/" + n + "?$select=id,displayName" })));
      slice.forEach((n) => {
        const r = responses[n];
        if (r && r.status === 200 && r.body && r.body.id) ids[n] = r.body.id;
        else if (!r || !NOT_READABLE[r.status]) unresolved.push(n);
      });
    }
    return { ids: ids, complete: unresolved.length === 0, unresolved: unresolved };
  }

  function getMessage(token, id, opts) {
    return request(token, "/me/messages/" + encodeURIComponent(id) + "?" + MSG_SELECT, opts);
  }

  /** True if the folder still exists, false on 404; other failures throw. */
  async function folderExists(token, id) {
    try {
      await request(token, "/me/mailFolders/" + encodeURIComponent(id) + "?$select=id", { retries: 0, urgent: true });
      return true;
    } catch (e) {
      if (e.status === 404) return false;
      throw e;
    }
  }

  async function findByInternetMessageId(token, internetMessageId) {
    const filter = "internetMessageId eq '" + q(internetMessageId) + "'";
    const d = await request(token, "/me/messages?$filter=" + encodeURIComponent(filter) + "&" + MSG_SELECT + "&$top=5");
    return (d.value || [])[0] || null;
  }

  async function messagesInConversation(token, conversationId, top) {
    const filter = "conversationId eq '" + q(conversationId) + "'";
    const d = await request(token, "/me/messages?$filter=" + encodeURIComponent(filter) + "&$select=id,parentFolderId&$top=" + top);
    return d.value || [];
  }

  /**
   * One round trip for everything we need about the open email: the message itself (for its Graph id and folder)
   * plus the three history lookups. p = { restId, sender, conversationId, domain, top }.
   * Returns { msg, msgError, sender, conversation, domain, failed }. A history lookup that fails yields [].
   */
  async function messageAndHistory(token, p) {
    const hist = "$select=id,parentFolderId&$top=" + p.top;
    const requests = [{ id: "msg", url: "/me/messages/" + encodeURIComponent(p.restId) + "?" + MSG_SELECT }];
    if (p.sender && p.sender.length <= EQ_MAX) {
      requests.push({ id: "sender", url: "/me/messages?$filter=" + encodeURIComponent("from/emailAddress/address eq '" + q(p.sender) + "'") + "&" + hist });
    }
    if (p.conversationId && p.conversationId.length <= EQ_MAX) {
      requests.push({ id: "conv", url: "/me/messages?$filter=" + encodeURIComponent("conversationId eq '" + q(p.conversationId) + "'") + "&" + hist });
    }
    if (p.domain) {
      requests.push({ id: "domain", url: "/me/messages?$search=" + encodeURIComponent("\"from:" + p.domain + "\"") + "&" + hist });
    }

    const r = await batch(token, requests);
    const list = (id) => (r[id] && r[id].status === 200 && r[id].body && r[id].body.value) || [];
    const m = r.msg || { status: 0, body: NO_ANSWER };
    const err = (m.body && m.body.error) || {};
    return {
      msg: m.status === 200 ? m.body : null,
      msgError: m.status === 200 ? null : new GraphError(m.status, err.code, err.message, m.body),
      sender: list("sender"),
      conversation: list("conv"),
      domain: list("domain"),
      failed: ["sender", "conv", "domain"].filter((id) => r[id] && r[id].status !== 200)
    };
  }

  function moveMessage(token, id, destinationId) {
    return request(token, "/me/messages/" + encodeURIComponent(id) + "/move", { method: "POST", body: { destinationId: destinationId }, urgent: true });
  }

  function createFolder(token, displayName, parentId) {
    const path = parentId ? "/me/mailFolders/" + encodeURIComponent(parentId) + "/childFolders" : "/me/mailFolders";
    return request(token, path, { method: "POST", body: { displayName: displayName }, urgent: true });
  }

  async function findChildFolderByName(token, parentId, name) {
    const path = parentId ? "/me/mailFolders/" + encodeURIComponent(parentId) + "/childFolders" : "/me/mailFolders";
    const filter = "displayName eq '" + q(name) + "'";
    const d = await request(token, path + "?$filter=" + encodeURIComponent(filter) + "&$select=id,displayName,parentFolderId,childFolderCount,totalItemCount", { urgent: true });
    return (d.value || [])[0] || null;
  }

  global.SmartFileGraph = {
    GraphError, request, batch, urgently, listAllFolders, wellKnownFolderIds, getMessage, folderExists, findByInternetMessageId,
    messagesInConversation, messageAndHistory, moveMessage, createFolder, findChildFolderByName
  };
})(typeof window !== "undefined" ? window : globalThis);
