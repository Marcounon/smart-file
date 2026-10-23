/* Smart File - Microsoft Graph calls (mail folders, message history, move, create folder). */
(function (global) {
  "use strict";

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const FOLDER_SELECT = "$select=id,displayName,parentFolderId,childFolderCount,totalItemCount&$top=250";

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

  async function request(token, path, opts) {
    opts = opts || {};
    const method = opts.method || "GET";
    const retries = opts.retries == null ? 2 : opts.retries;
    const url = path.startsWith("http") ? path : GRAPH + path;
    const headers = Object.assign({ Authorization: "Bearer " + token, Accept: "application/json" }, opts.headers || {});
    if (opts.body) headers["Content-Type"] = "application/json";

    let res;
    try {
      res = await fetch(url, { method: method, headers: headers, body: opts.body ? JSON.stringify(opts.body) : undefined });
    } catch (e) {
      throw new GraphError(0, "network", "Can\u2019t reach Microsoft Graph. Check your network connection.");
    }

    if ((res.status === 429 || res.status === 503 || res.status === 504) && retries > 0) {
      const wait = Math.min(10, Number(res.headers.get("Retry-After")) || 2) * 1000;
      await sleep(wait);
      return request(token, path, Object.assign({}, opts, { retries: retries - 1 }));
    }
    if (res.status === 204) return null;

    const text = await res.text();
    let data = null;
    if (text) { try { data = JSON.parse(text); } catch (e) { data = { raw: text }; } }
    if (!res.ok) {
      const err = (data && data.error) || {};
      throw new GraphError(res.status, err.code, err.message, data);
    }
    return data;
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

  /** Walk the whole folder tree (breadth-first) and return a flat list with paths. */
  async function listAllFolders(token, maxDepth) {
    const out = [];
    let level = (await listAll(token, "/me/mailFolders?" + FOLDER_SELECT))
      .map((f) => Object.assign({}, f, { depth: 0, pathParts: [f.displayName] }));

    while (level.length) {
      out.push.apply(out, level);
      const parents = level.filter((f) => f.childFolderCount > 0 && f.depth < maxDepth);
      const next = [];
      for (let i = 0; i < parents.length; i += 4) {           // 4 requests at a time, avoids throttling
        const batch = parents.slice(i, i + 4);
        const children = await Promise.all(batch.map((p) =>
          listAll(token, "/me/mailFolders/" + encodeURIComponent(p.id) + "/childFolders?" + FOLDER_SELECT).catch(() => [])
        ));
        children.forEach((list, idx) => {
          const p = batch[idx];
          list.forEach((c) => next.push(Object.assign({}, c, { depth: p.depth + 1, pathParts: p.pathParts.concat([c.displayName]) })));
        });
      }
      level = next;
    }

    return out.map((f) => ({
      id: f.id,
      name: f.displayName,
      parentId: f.parentFolderId,
      depth: f.depth,
      pathParts: f.pathParts,
      path: f.pathParts.join(" / "),
      totalItemCount: f.totalItemCount || 0,
      childFolderCount: f.childFolderCount || 0
    }));
  }

  /** Resolve well-known folder names (inbox, sentitems, ...) to IDs with one $batch call. */
  async function wellKnownFolderIds(token, names) {
    const out = {};
    for (let i = 0; i < names.length; i += 20) {
      const requests = names.slice(i, i + 20).map((n) => ({ id: n, method: "GET", url: "/me/mailFolders/" + n + "?$select=id,displayName" }));
      const data = await request(token, "/$batch", { method: "POST", body: { requests: requests } });
      (data.responses || []).forEach((r) => {
        if (r.status === 200 && r.body && r.body.id) out[r.id] = r.body.id;
      });
    }
    return out;
  }

  const MSG_SELECT = "$select=id,parentFolderId,conversationId,internetMessageId,subject";

  function getMessage(token, id) {
    return request(token, "/me/messages/" + encodeURIComponent(id) + "?" + MSG_SELECT);
  }

  async function findByInternetMessageId(token, internetMessageId) {
    const filter = "internetMessageId eq '" + q(internetMessageId) + "'";
    const d = await request(token, "/me/messages?$filter=" + encodeURIComponent(filter) + "&" + MSG_SELECT + "&$top=5");
    return (d.value || [])[0] || null;
  }

  async function messagesFromSender(token, address, top) {
    const filter = "from/emailAddress/address eq '" + q(address) + "'";
    const d = await request(token, "/me/messages?$filter=" + encodeURIComponent(filter) + "&$select=id,parentFolderId&$top=" + top);
    return d.value || [];
  }

  async function messagesInConversation(token, conversationId, top) {
    const filter = "conversationId eq '" + q(conversationId) + "'";
    const d = await request(token, "/me/messages?$filter=" + encodeURIComponent(filter) + "&$select=id,parentFolderId&$top=" + top);
    return d.value || [];
  }

  async function messagesFromDomain(token, domain, top) {
    const search = encodeURIComponent("\"from:" + domain + "\"");
    const d = await request(token, "/me/messages?$search=" + search + "&$select=id,parentFolderId&$top=" + top);
    return d.value || [];
  }

  function moveMessage(token, id, destinationId) {
    return request(token, "/me/messages/" + encodeURIComponent(id) + "/move", { method: "POST", body: { destinationId: destinationId } });
  }

  function createFolder(token, displayName, parentId) {
    const path = parentId ? "/me/mailFolders/" + encodeURIComponent(parentId) + "/childFolders" : "/me/mailFolders";
    return request(token, path, { method: "POST", body: { displayName: displayName } });
  }

  async function findChildFolderByName(token, parentId, name) {
    const path = parentId ? "/me/mailFolders/" + encodeURIComponent(parentId) + "/childFolders" : "/me/mailFolders";
    const filter = "displayName eq '" + q(name) + "'";
    const d = await request(token, path + "?$filter=" + encodeURIComponent(filter) + "&$select=id,displayName,parentFolderId,childFolderCount,totalItemCount");
    return (d.value || [])[0] || null;
  }

  global.SmartFileGraph = {
    GraphError, request, listAllFolders, wellKnownFolderIds, getMessage, findByInternetMessageId,
    messagesFromSender, messagesInConversation, messagesFromDomain, moveMessage, createFolder, findChildFolderByName
  };
})(typeof window !== "undefined" ? window : globalThis);
