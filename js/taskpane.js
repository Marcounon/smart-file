/* Smart File - task pane flow.
 *   File button  ->  "File this email in <folder>?"  ->  Yes: move it
 *                                                    ->  No:  3 other folders, create a new folder, or browse all
 */
(function () {
  "use strict";

  const cfg = window.SmartFileConfig;
  const Graph = window.SmartFileGraph;
  const Suggest = window.SmartFileSuggest;
  const Auth = window.SmartFileAuth;

  const FOLDER_CACHE_MS = 10 * 60 * 1000;

  const state = {
    runId: 0,
    folders: null,        // prepared candidate list (all folders, not yet filtered)
    foldersAt: 0,
    wellKnown: null,
    session: null,        // { ctx, msg, ranking, candidates }
    lastMove: null        // { newId, originalFolderId, folder, subject, sender }
  };

  /* ---------- tiny DOM helpers ---------- */

  function h(tag, attrs) {
    const el = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v == null) return;
      if (k === "class") el.className = v;
      else if (k === "onclick") el.addEventListener("click", v);
      else if (k === "onkeydown") el.addEventListener("keydown", v);
      else if (k === "oninput") el.addEventListener("input", v);
      else if (k === "text") el.textContent = v;
      else el.setAttribute(k, v);
    });
    for (let i = 2; i < arguments.length; i++) {
      const c = arguments[i];
      if (c == null) continue;
      if (Array.isArray(c)) c.forEach((x) => { if (x != null) el.append(x.nodeType ? x : document.createTextNode(String(x))); });
      else el.append(c.nodeType ? c : document.createTextNode(String(c)));
    }
    return el;
  }

  function setView() {
    const root = document.getElementById("view");
    root.replaceChildren.apply(root, Array.from(arguments).filter(Boolean));
    renderFooter();
  }

  function button(label, onClick, kind) {
    return h("button", { class: "btn " + (kind || "secondary"), type: "button", onclick: onClick }, label);
  }

  function emailCard(ctx) {
    return h("div", { class: "email" },
      h("div", { class: "email-subject", text: ctx.subject || "(no subject)" }),
      h("div", { class: "email-from", text: ctx.fromName ? ctx.fromName + " \u2039" + ctx.fromAddress + "\u203a" : ctx.fromAddress })
    );
  }

  function folderBlock(folder, reasons, big) {
    const parents = folder.pathParts.slice(0, -1).join(" / ");
    return h("span", { class: "folder" + (big ? " folder-big" : "") },
      h("span", { class: "folder-name", text: folder.name }),
      parents ? h("span", { class: "folder-path", text: parents }) : null,
      reasons && reasons.length ? h("span", { class: "reasons" }, reasons.slice(0, 2).map((r) => h("span", { class: "reason", text: r }))) : null
    );
  }

  /* ---------- Office bootstrap ---------- */

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) {
      renderError("Open Smart File from the File button in Outlook.");
      return;
    }
    try {
      // Fires when the pane is pinned and the user selects another email.
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => run());
    } catch (e) { /* older clients without pinning */ }
    run();
  });

  /* ---------- main flow ---------- */

  async function run() {
    const runId = ++state.runId;
    const item = Office.context.mailbox.item;
    if (!item) { renderEmpty(); return; }

    renderLoading("Reading this email\u2026");
    try {
      const ctx = await readItem(item);
      if (runId !== state.runId) return;

      renderLoading("Checking your folders\u2026");
      const token = (await Auth.getToken()).accessToken;
      if (runId !== state.runId) return;

      const [folders, wellKnown, msg] = await Promise.all([
        getFolders(token),
        getWellKnown(token),
        resolveMessage(token, ctx)
      ]);
      if (runId !== state.runId) return;

      const candidates = candidateFolders(folders, wellKnown, msg);
      const session = { ctx: ctx, msg: msg, candidates: candidates, ranking: null };
      state.session = session;

      if (!candidates.length) { renderNoFolders(session); return; }

      renderLoading("Looking for the best match\u2026");
      const history = await gatherHistory(token, ctx, msg, candidates);
      if (runId !== state.runId) return;

      session.ranking = Suggest.rankFolders(candidates, ctx, history, Rules.load());
      renderSuggestion(session);
    } catch (e) {
      if (runId !== state.runId) return;
      if (e instanceof Auth.AuthRequired) renderSignIn(e);
      else renderError(friendlyError(e), run);
    }
  }

  function readItem(item) {
    return new Promise((resolve) => {
      const from = item.from || item.sender || {};
      const me = (Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
      const mapAddr = (list) => (list || []).map((r) => ({ name: r.displayName || "", address: (r.emailAddress || "").toLowerCase() }));
      const recipients = mapAddr(item.to).concat(mapAddr(item.cc));
      const fromAddress = (from.emailAddress || "").toLowerCase();

      // The "counterparty" is who this email is really about: the sender, unless I sent it, then the first outside recipient.
      let counterparty = { name: from.displayName || "", address: fromAddress };
      if (fromAddress && fromAddress === me) {
        const myDomain = Suggest.domainOf(me);
        counterparty = recipients.find((r) => r.address && Suggest.domainOf(r.address) !== myDomain) || recipients[0] || counterparty;
      }

      const ctx = {
        subject: item.normalizedSubject || item.subject || "",
        fromName: from.displayName || "",
        fromAddress: fromAddress,
        recipients: recipients,
        userAddress: me,
        counterparty: counterparty,
        conversationId: item.conversationId || "",
        internetMessageId: item.internetMessageId || "",
        itemId: item.itemId,
        restId: toRestId(item.itemId),
        bodyText: ""
      };

      try {
        item.body.getAsync(Office.CoercionType.Text, (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) ctx.bodyText = (res.value || "").slice(0, cfg.bodySampleChars);
          resolve(ctx);
        });
      } catch (e) { resolve(ctx); }
    });
  }

  function toRestId(itemId) {
    try { return Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0); }
    catch (e) { return itemId; }
  }

  async function getFolders(token) {
    if (state.folders && Date.now() - state.foldersAt < FOLDER_CACHE_MS) return state.folders;
    const list = await Graph.listAllFolders(token, cfg.maxFolderDepth);
    state.folders = list.map(Suggest.prepareFolder);
    state.foldersAt = Date.now();
    return state.folders;
  }

  async function getWellKnown(token) {
    if (state.wellKnown) return state.wellKnown;
    state.wellKnown = await Graph.wellKnownFolderIds(token, cfg.excludedWellKnownFolders);
    return state.wellKnown;
  }

  async function resolveMessage(token, ctx) {
    try {
      return await Graph.getMessage(token, ctx.restId);
    } catch (e) {
      if ((e.status === 404 || e.status === 400) && ctx.internetMessageId) {
        const m = await Graph.findByInternetMessageId(token, ctx.internetMessageId);
        if (m) return m;
      }
      throw e;
    }
  }

  function candidateFolders(folders, wellKnown, msg) {
    const excludedIds = new Set(Object.values(wellKnown));
    excludedIds.add(msg.parentFolderId); // never suggest the folder the email is already in
    const excludedNames = new Set(cfg.excludedFolderNames.map((n) => n.toLowerCase()));
    return folders.filter((f) => !excludedIds.has(f.id) && !(f.depth === 0 && excludedNames.has(f.name.toLowerCase())));
  }

  async function gatherHistory(token, ctx, msg, candidates) {
    const candidateIds = new Set(candidates.map((f) => f.id));
    const n = cfg.historySampleSize;
    const count = (list) => {
      const byFolder = {};
      let total = 0;
      (list || []).forEach((m) => {
        if (m.id !== msg.id && candidateIds.has(m.parentFolderId)) { byFolder[m.parentFolderId] = (byFolder[m.parentFolderId] || 0) + 1; total++; }
      });
      return { byFolder: byFolder, total: total };
    };

    const cp = ctx.counterparty.address;
    const domain = Suggest.domainOf(cp);
    const myDomain = Suggest.domainOf(ctx.userAddress);
    const useDomain = domain && !Suggest.isPublicDomain(domain) && domain !== myDomain;

    const [sender, conv, dom] = await Promise.all([
      cp ? Graph.messagesFromSender(token, cp, n).catch(() => []) : [],
      msg.conversationId ? Graph.messagesInConversation(token, msg.conversationId, n).catch(() => []) : [],
      useDomain ? Graph.messagesFromDomain(token, domain, n).catch(() => []) : []
    ]);

    const s = count(sender), c = count(conv), d = count(dom);
    return { sender: s.byFolder, senderTotal: s.total, conversation: c.byFolder, domain: d.byFolder };
  }

  /* ---------- actions ---------- */

  async function fileTo(session, folder) {
    renderLoading("Moving to " + folder.name + "\u2026");
    try {
      const token = (await ensureToken()).accessToken;
      const moved = await Graph.moveMessage(token, session.msg.id, folder.id);
      state.lastMove = {
        newId: (moved && moved.id) || session.msg.id,
        originalFolderId: session.msg.parentFolderId,
        folder: folder,
        subject: session.ctx.subject,
        sender: session.ctx.counterparty.address
      };
      Rules.remember(session.ctx.counterparty.address, folder.id);
      renderFiled(folder);
    } catch (e) {
      if (e instanceof Auth.AuthRequired) renderSignIn(e);
      else renderError(friendlyError(e), () => renderOptions(session));
    }
  }

  async function createAndFile(session, name, parentId) {
    name = (name || "").trim();
    if (!name) return;
    renderLoading("Creating \u201c" + name + "\u201d\u2026");
    try {
      const token = (await ensureToken()).accessToken;
      let created;
      try {
        created = await Graph.createFolder(token, name, parentId);
      } catch (e) {
        if (e.status === 409 || /exist/i.test(e.message || "")) created = await Graph.findChildFolderByName(token, parentId, name);
        if (!created) throw e;
      }
      const parent = parentId ? state.folders.find((f) => f.id === parentId) : null;
      const folder = Suggest.prepareFolder({
        id: created.id,
        name: created.displayName,
        parentId: parentId || null,
        depth: parent ? parent.depth + 1 : 0,
        pathParts: parent ? parent.pathParts.concat([created.displayName]) : [created.displayName],
        path: parent ? parent.path + " / " + created.displayName : created.displayName,
        totalItemCount: created.totalItemCount || 0,
        childFolderCount: 0
      });
      if (!state.folders.some((f) => f.id === folder.id)) state.folders.push(folder);
      session.candidates.push(folder);
      await fileTo(session, folder);
    } catch (e) {
      if (e instanceof Auth.AuthRequired) renderSignIn(e);
      else renderError(friendlyError(e), () => renderOptions(session));
    }
  }

  async function undoLastMove() {
    const lm = state.lastMove;
    if (!lm) return;
    renderLoading("Moving it back\u2026");
    try {
      const token = (await ensureToken()).accessToken;
      await Graph.moveMessage(token, lm.newId, lm.originalFolderId);
      Rules.forget(lm.sender, lm.folder.id);
      state.lastMove = null;
      setView(h("div", { class: "status" }, h("p", { text: "Moved back to where it was." }), button("Suggest a folder again", run, "secondary")));
    } catch (e) {
      renderError(friendlyError(e));
    }
  }

  async function ensureToken() {
    return Auth.getToken();
  }

  async function signIn() {
    renderLoading("Opening sign-in\u2026");
    try {
      await Auth.getTokenInteractive();
      run();
    } catch (e) {
      renderSignIn(e, Auth.describeError(e));
    }
  }

  /* ---------- learned rules (per mailbox, synced by Outlook) ---------- */

  const Rules = {
    key: "smartfile.rules.v1",
    load() {
      try {
        const raw = Office.context.roamingSettings.get(this.key);
        const data = raw ? JSON.parse(raw) : {};
        return { bySender: data.bySender || {} };
      } catch (e) { return { bySender: {} }; }
    },
    remember(sender, folderId) {
      if (!sender || !folderId) return;
      const data = this.load();
      data.bySender[sender] = { f: folderId, t: Date.now() };
      const entries = Object.entries(data.bySender);
      if (entries.length > cfg.maxLearnedRules) {
        entries.sort((a, b) => a[1].t - b[1].t);
        entries.slice(0, entries.length - cfg.maxLearnedRules).forEach(([k]) => { delete data.bySender[k]; });
      }
      this.save(data);
    },
    forget(sender, folderId) {
      const data = this.load();
      if (data.bySender[sender] && data.bySender[sender].f === folderId) { delete data.bySender[sender]; this.save(data); }
    },
    save(data) {
      try {
        Office.context.roamingSettings.set(this.key, JSON.stringify(data));
        Office.context.roamingSettings.saveAsync(() => {});
      } catch (e) { /* best effort */ }
    }
  };

  /* ---------- views ---------- */

  function renderLoading(text) {
    setView(h("div", { class: "loading" }, h("div", { class: "spinner", "aria-hidden": "true" }), h("p", { text: text, role: "status" })));
  }

  function renderEmpty() {
    state.session = null;
    setView(h("div", { class: "status" }, h("p", { text: "Select an email to file it." })));
  }

  function renderError(message, retry) {
    setView(h("div", { class: "status error" },
      h("p", { text: message }),
      retry ? button("Try again", retry, "secondary") : null
    ));
  }

  function renderSignIn(err, detail) {
    setView(h("div", { class: "status" },
      h("p", { text: "Smart File needs permission to read your folders and move email. You only have to do this once." }),
      detail ? h("p", { class: "muted", text: detail }) : null,
      button("Sign in with Microsoft 365", signIn, "primary")
    ));
  }

  function renderNoFolders(session) {
    setView(
      emailCard(session.ctx),
      h("div", { class: "status" }, h("p", { text: "There are no folders to file into yet. Create the first one:" })),
      newFolderForm(session, session.ctx.fromName, null)
    );
  }

  function renderSuggestion(session) {
    const r = session.ranking;
    const top = r.top;
    const yes = button("Yes, file it", () => fileTo(session, top.folder), "primary");
    setView(
      emailCard(session.ctx),
      h("p", { class: "question", text: r.confident ? "File this email in" : "No strong match. Best guess:" }),
      folderBlock(top.folder, top.reasons, true),
      h("div", { class: "actions" },
        yes,
        button("No, show other options", () => renderOptions(session), "secondary")
      )
    );
    yes.focus();
  }

  function renderOptions(session) {
    const r = session.ranking;
    const rejected = r.top ? r.top.folder : null;
    const alternatives = r.alternatives.filter((a) => a.folder !== rejected);

    const parentDefault = defaultParentFor(session);
    const list = h("div", { class: "option-list" }, alternatives.map((alt) =>
      h("button", { class: "option", type: "button", onclick: () => fileTo(session, alt.folder) }, folderBlock(alt.folder, alt.reasons, false))
    ));

    setView(
      emailCard(session.ctx),
      h("p", { class: "question", text: alternatives.length ? "Other likely folders" : "No other likely folders" }),
      list,
      h("p", { class: "question", text: "Or create a new folder" }),
      newFolderForm(session, r.suggestedNewName, parentDefault),
      browseAll(session)
    );
  }

  function defaultParentFor(session) {
    if (cfg.newFolderDefaultParent) {
      const named = state.folders.find((f) => f.name.toLowerCase() === cfg.newFolderDefaultParent.toLowerCase());
      if (named) return named.id;
    }
    const top = session.ranking && session.ranking.top && session.ranking.top.folder;
    if (top && top.parentId && state.folders.some((f) => f.id === top.parentId)) return top.parentId;
    return "";
  }

  function newFolderForm(session, suggestedName, parentDefault) {
    const name = h("input", { class: "input", type: "text", placeholder: "Folder name", value: suggestedName || "", "aria-label": "New folder name", maxlength: "255" });
    const parent = h("select", { class: "input", "aria-label": "Create inside" },
      h("option", { value: "", text: "Top level (next to Inbox)" }),
      state.folders.slice().sort((a, b) => a.path.localeCompare(b.path)).map((f) => h("option", { value: f.id, text: f.path }))
    );
    parent.value = parentDefault || "";
    const create = () => createAndFile(session, name.value, parent.value || null);
    name.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); create(); } });
    return h("div", { class: "new-folder" },
      name,
      h("label", { class: "field" }, h("span", { class: "muted", text: "Create inside" }), parent),
      button("Create folder and file it", create, "primary")
    );
  }

  function browseAll(session) {
    const results = h("div", { class: "option-list" });
    const search = h("input", { class: "input", type: "search", placeholder: "Search all folders", "aria-label": "Search all folders" });
    const refresh = () => {
      const needle = search.value.trim().toLowerCase();
      results.replaceChildren();
      if (!needle) return;
      const matches = session.candidates.filter((f) => f.path.toLowerCase().includes(needle)).slice(0, 8);
      if (!matches.length) { results.append(h("p", { class: "muted", text: "No folder matches." })); return; }
      matches.forEach((f) => results.append(h("button", { class: "option", type: "button", onclick: () => fileTo(session, f) }, folderBlock(f, null, false))));
    };
    search.addEventListener("input", refresh);
    return h("details", { class: "browse" },
      h("summary", { text: "Browse all folders" }),
      search,
      results
    );
  }

  function renderFiled(folder) {
    setView(h("div", { class: "done" },
      h("div", { class: "check", "aria-hidden": "true" }, "\u2713"),
      h("p", { class: "done-text" }, "Filed in ", h("strong", { text: folder.name })),
      folder.pathParts.length > 1 ? h("p", { class: "muted", text: folder.pathParts.slice(0, -1).join(" / ") }) : null,
      h("div", { class: "actions" }, button("Undo", undoLastMove, "secondary")),
      h("p", { class: "muted hint", text: "Pin this pane to file the next email without reopening it." })
    ));
  }

  function renderFooter() {
    const footer = document.getElementById("footer");
    const lm = state.lastMove;
    const showing = document.querySelector("#view .done");
    if (!lm || showing) { footer.replaceChildren(); footer.hidden = true; return; }
    footer.hidden = false;
    footer.replaceChildren(
      h("span", { class: "footer-text" }, "Filed \u201c", h("span", { class: "footer-subject", text: lm.subject || "(no subject)" }), "\u201d in ", h("strong", { text: lm.folder.name })),
      h("button", { class: "link", type: "button", onclick: undoLastMove }, "Undo")
    );
  }

  function friendlyError(e) {
    if (e && e.name === "GraphError") {
      if (e.status === 401 || e.status === 403) return "Microsoft 365 refused the request (" + (e.code || e.status) + "). Check that the Entra app has the Mail.ReadWrite permission and that consent was granted.";
      if (e.status === 404) return "This email can\u2019t be found any more. It may already have been moved or deleted.";
      if (e.status === 0) return e.message;
      return "Microsoft 365 returned an error" + (e.code ? " (" + e.code + ")" : "") + ": " + e.message;
    }
    return (e && e.message) || String(e);
  }
})();
