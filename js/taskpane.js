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

  const PATCHY_RETRY_MS = 10 * 60 * 1000;   // how often a folder list with known gaps is re-walked
  const WALK_SPACING_MS = 2 * 60 * 1000;    // never two walks closer together than this

  const state = {
    runId: 0,
    activeItemId: null,   // item the pane is currently showing or loading; stops duplicate ItemChanged events re-running everything
    folders: null,        // prepared candidate list (all folders, not yet filtered)
    foldersAt: 0,         // when the folder list was last fetched from Graph (0 = refresh it at the next chance)
    foldersComplete: false, // false while the list is known to be missing subtrees or well-known folders
    wellKnown: null,      // well-known folder name -> id
    wellKnownUnresolved: [], // well-known names whose id could not be fetched; excluded by display name instead
    refreshing: null,     // in-flight folder refresh, so two callers share one walk
    walkedAt: 0,          // when the tree was last walked, complete or not
    lastWalkComplete: true, // false after a walk that came back with gaps (or failed); such walks are retried less often
    createdDuringRefresh: [], // folders created from the pane while a walk was running; merged into its result
    session: null,        // { ctx, msg, ranking, candidates, rejected }
    moveSeq: 0,           // counts moves in click order, so a slow older move cannot become "the last move"
    lastMove: null,       // { seq, newId, originalFolderId, folder, subject, sender }
    notice: null          // a failure that happened after the pane had moved on to another email; shown in the footer
  };

  // Without at least these, ranking is unsafe: Sent Items or Deleted Items could win on history alone.
  const CRITICAL_WELL_KNOWN = ["inbox", "sentitems", "deleteditems", "junkemail", "drafts"];

  // English display names of the well-known folders. Only a safety net: when an id could not be fetched, a
  // top-level folder with this name is left out of the suggestions rather than offered.
  const WELL_KNOWN_DISPLAY = {
    inbox: "Inbox", drafts: "Drafts", sentitems: "Sent Items", deleteditems: "Deleted Items", junkemail: "Junk Email",
    outbox: "Outbox", archive: "Archive", conversationhistory: "Conversation History", syncissues: "Sync Issues",
    recoverableitemsdeletions: "Recoverable Items", clutter: "Clutter", scheduled: "Scheduled", searchfolders: "Search Folders",
    serverfailures: "Server Failures", localfailures: "Local Failures", conflicts: "Conflicts"
  };

  /* ---------- timing (shown under the suggestion while cfg.showTimings is on) ---------- */

  const now = () => (typeof performance !== "undefined" && performance.now) ? performance.now() : Date.now();

  // One timer per run, so a run that is abandoned half-way cannot write into the next one's line.
  function makeTiming() {
    const t0 = now();
    const parts = [];
    let last = t0;
    const timing = {
      step(label) { const t = now(); parts.push([label, t - last]); last = t; },
      record(label, ms) { parts.push([label, ms]); last = now(); },
      span(label, promise) {
        const t = now();
        return promise.then((v) => { timing.record(label, now() - t); return v; }, (e) => { timing.record(label, now() - t); throw e; });
      },
      summary() {
        const s = (ms) => (ms / 1000).toFixed(1) + "s";
        return parts.map(([l, ms]) => l + " " + s(ms)).join(" · ") + " · total " + s(now() - t0);
      }
    };
    return timing;
  }

  /* ---------- folder list remembered between opens ----------
   * Walking 200+ folders is the slowest thing the add-in does, and the pane is reloaded every time the button is
   * pressed, so the list is kept in localStorage (per mailbox) and refreshed in the background when it gets old. */

  const FolderCache = {
    key() {
      const me = ((Office.context.mailbox.userProfile && Office.context.mailbox.userProfile.emailAddress) || "").toLowerCase();
      return "smartfile.folders.v1:" + (Office.context.partitionKey || "") + ":" + me;
    },
    load() {
      try {
        const data = JSON.parse(localStorage.getItem(this.key()) || "null");
        return data && Array.isArray(data.folders) && data.wellKnown ? data : null;
      } catch (e) { return null; }
    },
    save(data) {
      try { localStorage.setItem(this.key(), JSON.stringify(data)); } catch (e) { /* storage blocked or full: fetch again next time */ }
    }
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
      h("div", { class: "email-from", text: ctx.fromName ? ctx.fromName + " ‹" + ctx.fromAddress + "›" : ctx.fromAddress })
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

  /**
   * Suggest a folder for the selected email. Without `force`, a repeat call for the email already on screen (or
   * being loaded) is ignored, because ItemChanged can fire more than once for one selection. Buttons that the
   * user presses on purpose ("Try again", sign-in, refresh) use rerun(), which always starts over.
   */
  async function run(force) {
    const item = Office.context.mailbox.item;
    if (!item) { state.runId++; state.activeItemId = null; renderEmpty(); return; }   // bumping runId cancels a run in flight
    if (!force && item.itemId && item.itemId === state.activeItemId) return;
    const runId = ++state.runId;
    state.activeItemId = item.itemId || null;

    const timing = makeTiming();
    renderLoading("Reading this email…");
    try {
      const tokenPromise = Auth.getToken();              // overlaps with Outlook handing us the body text
      tokenPromise.catch(() => {});                      // errors are handled below, once awaited
      const ctx = await readItem(item);
      if (runId !== state.runId) return;
      timing.step("read");
      const token = (await tokenPromise).accessToken;
      if (runId !== state.runId) return;
      timing.step("token");

      renderLoading("Checking your folders…");
      const [tree, lookup] = await Promise.all([
        getFolderTree(token, timing),
        timing.span("history", Graph.messageAndHistory(token, {
          restId: ctx.restId,
          sender: ctx.counterparty.address,
          conversationId: ctx.conversationId,
          domain: historyDomain(ctx),
          top: cfg.historySampleSize
        }))
      ]);
      if (runId !== state.runId) return;

      const msg = lookup.msg || await resolveMessageFallback(token, ctx, lookup.msgError);
      // Outlook and Graph normally agree on the conversation id (once both are in REST form). If they don't,
      // the batched lookup searched for the wrong conversation, so look it up again with Graph's id.
      let conversation = lookup.conversation;
      if (msg.conversationId && urlSafeId(msg.conversationId) !== ctx.conversationId) {
        conversation = await timing.span("conversation", Graph.messagesInConversation(token, msg.conversationId, cfg.historySampleSize).catch(() => []));
      }
      if (runId !== state.runId) return;

      const candidates = candidateFolders(tree.folders, tree.wellKnown, tree.unresolved, msg);
      const session = { ctx: ctx, msg: msg, candidates: candidates, ranking: null, rejected: null };
      state.session = session;
      if (!candidates.length) { renderNoFolders(session); refreshFoldersIfOld(token); return; }

      const history = countHistory({ sender: lookup.sender, conversation: conversation, domain: lookup.domain }, msg, candidates);
      session.ranking = Suggest.rankFolders(candidates, ctx, history, Rules.load());
      timing.step("ranked");
      renderSuggestion(session, timing);
      console.log("Smart File timing: " + timing.summary());
      refreshFoldersIfOld(token);
    } catch (e) {
      if (runId !== state.runId) return;
      if (e instanceof Auth.AuthRequired) renderSignIn(e);
      else renderError(friendlyError(e), rerun);
    }
  }

  const rerun = () => run(true);

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
        conversationId: urlSafeId(item.conversationId || ""),
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

  // Outlook hands out ids in Exchange form; Graph uses the URL-safe form of the same id ("/" -> "-", "+" -> "_").
  const urlSafeId = (id) => (id || "").replace(/\//g, "-").replace(/\+/g, "_");

  function toRestId(itemId) {
    try { return Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0); }
    catch (e) { return urlSafeId(itemId); }
  }

  /* ---------- folder list ---------- */

  const wellKnownNames = () => cfg.excludedWellKnownFolders.join(",");

  /** Keep a list in memory. An incomplete list gets fetchedAt 0, so it is refreshed in the background. */
  function adoptTree(folders, wellKnown, unresolved, fetchedAt, complete) {
    state.folders = folders.map(Suggest.prepareFolder);
    state.wellKnown = wellKnown;
    state.wellKnownUnresolved = unresolved || [];
    state.foldersComplete = !!complete;
    state.foldersAt = complete ? fetchedAt : 0;
  }

  // The fields worth remembering; prepareFolder adds the rest back when the list is loaded.
  function bareFolder(f) {
    return { id: f.id, name: f.name, parentId: f.parentId, depth: f.depth, pathParts: f.pathParts, path: f.path, totalItemCount: f.totalItemCount, childFolderCount: f.childFolderCount };
  }

  function saveTree() {
    if (!state.folders || !state.wellKnown) return;
    FolderCache.save({
      savedAt: state.foldersAt, walkedAt: state.walkedAt, complete: state.foldersComplete, lastWalkComplete: state.lastWalkComplete,
      wellKnownNames: wellKnownNames(), maxFolderDepth: cfg.maxFolderDepth,
      wellKnown: state.wellKnown, wellKnownUnresolved: state.wellKnownUnresolved, folders: state.folders.map(bareFolder)
    });
  }

  /** Folders from memory, then from the list remembered on this PC, and only then from Graph. */
  async function getFolderTree(token, timing) {
    const t = now();
    let source = "in memory";
    if (!state.folders) {
      const cached = FolderCache.load();
      if (cached && cached.wellKnownNames === wellKnownNames() && cached.maxFolderDepth === cfg.maxFolderDepth) {
        adoptTree(cached.folders, cached.wellKnown, cached.wellKnownUnresolved, cached.savedAt || 0, cached.complete !== false);
        state.walkedAt = cached.walkedAt || 0;
        state.lastWalkComplete = cached.lastWalkComplete !== false;
        source = "remembered";
      } else {
        await refreshFolderTree(token);
        source = "fetched";
      }
    }
    timing.record("folders " + source, now() - t);
    return { folders: state.folders, wellKnown: state.wellKnown, unresolved: state.wellKnownUnresolved };
  }

  /**
   * Walk the folder tree on Graph and remember it. Callers that overlap share one walk.
   * Resolves to { updated, complete, problems }. A walk that came back with gaps never replaces a complete list
   * we already hold; it is kept only when it is all we have, and then re-walked later.
   */
  function refreshFolderTree(token) {
    if (state.refreshing) return state.refreshing;
    state.refreshing = (async () => {
      let walkFinished = false;
      try {
        state.createdDuringRefresh = [];
        const tree = await Graph.listAllFolders(token, cfg.maxFolderDepth);
        const wk = await resolveWellKnown(token);
        state.walkedAt = Date.now();
        walkFinished = true;
        const complete = tree.complete && wk.complete;
        state.lastWalkComplete = complete;
        const problems = tree.problems.concat(wk.unresolved.map((n) => WELL_KNOWN_DISPLAY[n] || n));
        if (tree.gaps.length) console.warn("Smart File: these folders can’t be read and are skipped: " + tree.gaps.join(", "));
        if (!complete) console.warn("Smart File: folder list incomplete; could not read: " + problems.join(", "));
        if (!complete && state.folders && state.foldersComplete) {   // keep the good list, retry later
          state.foldersAt = 0;
          saveTree();                                      // same list; records when we last tried
          return { updated: false, complete: false, problems: problems };
        }
        // Folders created from the pane while the walk ran are not in its result yet.
        const known = new Set(tree.folders.map((f) => f.id));
        const folders = tree.folders.concat(state.createdDuringRefresh.filter((f) => !known.has(f.id)).map(bareFolder));
        const ids = wk.complete ? wk.ids : Object.assign({}, state.wellKnown || {}, wk.ids);
        adoptTree(folders, ids, wk.unresolved.filter((n) => !ids[n]), Date.now(), complete);
        saveTree();
        return { updated: true, complete: complete, problems: problems };
      } finally {
        if (!walkFinished) { state.walkedAt = Date.now(); state.lastWalkComplete = false; }
        state.createdDuringRefresh = [];
        state.refreshing = null;
      }
    })();
    return state.refreshing;
  }

  /**
   * Resolve the well-known folder ids. A failed lookup falls back to the ids we already know; if that still leaves
   * none of the system folders identified, one more attempt is made and then the walk is abandoned rather than
   * ranking with Sent Items or Deleted Items in the running.
   */
  async function resolveWellKnown(token) {
    const known = state.wellKnown || {};
    const fallback = () => ({ ids: Object.assign({}, known), complete: false, unresolved: cfg.excludedWellKnownFolders.filter((n) => !known[n]) });
    let wk;
    try { wk = await Graph.wellKnownFolderIds(token, cfg.excludedWellKnownFolders); }
    catch (e) { wk = fallback(); }
    const anyCritical = (ids) => CRITICAL_WELL_KNOWN.some((n) => ids[n]);
    if (wk.complete || anyCritical(wk.ids) || anyCritical(known)) return wk;
    try { wk = await Graph.wellKnownFolderIds(token, cfg.excludedWellKnownFolders); }
    catch (e) { wk = fallback(); }
    if (wk.complete || anyCritical(wk.ids)) return wk;
    throw new Error("Couldn’t identify your system folders (Inbox, Sent Items, Deleted Items). Try again in a moment.");
  }

  function refreshFoldersIfOld(token) {
    const minutes = cfg.folderRefreshMinutes == null ? 30 : cfg.folderRefreshMinutes;
    const sinceWalk = Date.now() - state.walkedAt;
    if (!state.lastWalkComplete && sinceWalk < PATCHY_RETRY_MS) return;          // last try had gaps: don't hammer
    if (state.foldersComplete && state.foldersAt && Date.now() - state.foldersAt < minutes * 60 * 1000) return;
    if (sinceWalk < WALK_SPACING_MS) return;
    refreshFolderTree(token).catch(() => { /* next open will try again */ });
  }

  /** A folder turned out not to exist any more: stop offering it and refresh the list at the next chance. */
  function dropFolder(folderId, session) {
    if (state.folders) state.folders = state.folders.filter((f) => f.id !== folderId);
    if (session) {
      session.candidates = session.candidates.filter((f) => f.id !== folderId);
      const r = session.ranking;
      if (r) {
        r.ranked = r.ranked.filter((x) => x.folder.id !== folderId);
        r.alternatives = r.alternatives.filter((x) => x.folder.id !== folderId);
        if (r.top && r.top.folder.id === folderId) r.top = r.ranked[0] || null;
      }
    }
    state.foldersAt = 0;
    saveTree();
  }

  /* ---------- message and history ---------- */

  /** The batched message lookup failed: a throttled or unanswered item gets one direct try, a 404/400 falls back to the internet message id. */
  async function resolveMessageFallback(token, ctx, err) {
    const status = err ? err.status : 0;
    if (status === 0 || status === 429 || status === 503 || status === 504) {
      try { return await Graph.getMessage(token, ctx.restId, { retries: 0, urgent: true }); }
      catch (e) { err = e; }
    }
    if (err && (err.status === 404 || err.status === 400) && ctx.internetMessageId) {
      const m = await Graph.findByInternetMessageId(token, ctx.internetMessageId);
      if (m) return m;
    }
    throw err || new Error("Microsoft 365 did not return this email.");
  }

  function historyDomain(ctx) {
    const domain = Suggest.domainOf(ctx.counterparty.address);
    const myDomain = Suggest.domainOf(ctx.userAddress);
    return domain && !Suggest.isPublicDomain(domain) && domain !== myDomain ? domain : "";
  }

  function candidateFolders(folders, wellKnown, unresolved, msg) {
    const excludedIds = new Set(Object.values(wellKnown));
    excludedIds.add(msg.parentFolderId); // never suggest the folder the email is already in
    const excludedNames = new Set(cfg.excludedFolderNames.map((n) => n.toLowerCase()));
    (unresolved || []).forEach((n) => { if (WELL_KNOWN_DISPLAY[n]) excludedNames.add(WELL_KNOWN_DISPLAY[n].toLowerCase()); });
    return folders.filter((f) => !excludedIds.has(f.id) && !(f.depth === 0 && excludedNames.has(f.name.toLowerCase())));
  }

  /** Turn the raw history lists into per-folder counts, ignoring this email and folders we would never suggest. */
  function countHistory(lists, msg, candidates) {
    const candidateIds = new Set(candidates.map((f) => f.id));
    const count = (list) => {
      const byFolder = {};
      let total = 0;
      (list || []).forEach((m) => {
        if (m.id !== msg.id && candidateIds.has(m.parentFolderId)) { byFolder[m.parentFolderId] = (byFolder[m.parentFolderId] || 0) + 1; total++; }
      });
      return { byFolder: byFolder, total: total };
    };
    const s = count(lists.sender), c = count(lists.conversation), d = count(lists.domain);
    return { sender: s.byFolder, senderTotal: s.total, conversation: c.byFolder, domain: d.byFolder };
  }

  /* ---------- actions ----------
   * Each action remembers which run owned the pane when it started. If the user has moved on to another email by
   * the time it finishes, the result goes to the footer instead of replacing that email's view. */

  async function fileTo(session, folder, runId) {
    if (runId == null) runId = state.runId;
    const seq = ++state.moveSeq;
    state.notice = null;
    if (state.runId === runId) renderLoading("Moving to " + folder.name + "…");
    try {
      const token = (await ensureToken()).accessToken;
      const moved = await Graph.urgently(() => Graph.moveMessage(token, session.msg.id, folder.id));
      if (!state.lastMove || seq > state.lastMove.seq) {   // a slower, older move must not become "the last move"
        state.lastMove = {
          seq: seq,
          newId: (moved && moved.id) || session.msg.id,
          originalFolderId: session.msg.parentFolderId,
          folder: folder,
          subject: session.ctx.subject,
          sender: session.ctx.counterparty.address
        };
      }
      Rules.remember(session.ctx.counterparty.address, folder.id);
      if (state.activeItemId === session.ctx.itemId) {   // unless the pane has already moved on to another email
        state.session = null;                             // this email is gone from view; the next selection starts fresh
        state.activeItemId = null;
      }
      if (state.runId === runId) renderFiled(folder); else renderFooter();
    } catch (e) {
      if (e.status === 404) await handleMoveNotFound(session, folder);
      failed(runId, "Couldn’t move “" + (session.ctx.subject || "(no subject)") + "”", e, () => backTo(session));
    }
  }

  /** A move came back 404: either the folder was deleted since the list was remembered, or the email itself is gone. */
  async function handleMoveNotFound(session, folder) {
    let folderGone = false;
    try {
      const token = (await ensureToken()).accessToken;
      folderGone = !(await Graph.folderExists(token, folder.id));
    } catch (e) { /* can't tell; assume the email is the problem and keep the folder */ }
    if (folderGone) dropFolder(folder.id, session);
    else if (state.activeItemId === session.ctx.itemId) { state.session = null; state.activeItemId = null; }   // let a re-selection re-read it
  }

  async function createAndFile(session, name, parentId) {
    name = (name || "").trim();
    if (!name) return;
    const runId = state.runId;
    renderLoading("Creating “" + name + "”…");
    try {
      const token = (await ensureToken()).accessToken;
      const created = await Graph.urgently(async () => {
        try {
          return await Graph.createFolder(token, name, parentId);
        } catch (e) {
          const existing = (e.status === 409 || /exist/i.test(e.message || "")) ? await Graph.findChildFolderByName(token, parentId, name) : null;
          if (!existing) throw e;
          return existing;
        }
      });
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
      if (!state.folders.some((f) => f.id === folder.id)) {
        state.folders.push(folder);
        if (state.refreshing) state.createdDuringRefresh.push(folder);   // a walk in progress started before this folder existed
        saveTree();
      }
      session.candidates.push(folder);
      await fileTo(session, folder, runId);
    } catch (e) {
      failed(runId, "Couldn’t create “" + name + "”", e, () => backTo(session));
    }
  }

  async function undoLastMove() {
    const lm = state.lastMove;
    if (!lm) return;
    const runId = state.runId;
    state.notice = null;
    renderLoading("Moving it back…");
    try {
      const token = (await ensureToken()).accessToken;
      await Graph.urgently(() => Graph.moveMessage(token, lm.newId, lm.originalFolderId));
      Rules.forget(lm.sender, lm.folder.id);
      state.lastMove = null;
      if (state.runId !== runId) { renderFooter(); return; }
      state.session = null;
      state.activeItemId = null;
      setView(h("div", { class: "status" }, h("p", { text: "Moved back to where it was." }), button("Suggest a folder again", rerun, "secondary")));
    } catch (e) {
      failed(runId, "Couldn’t move “" + (lm.subject || "(no subject)") + "” back", e, undoLastMove);
    }
  }

  /** Show an action's failure: in place when its run still owns the pane, otherwise as a footer notice. */
  function failed(runId, what, e, retry) {
    const current = state.runId === runId;
    if (e instanceof Auth.AuthRequired) {
      if (current) renderSignIn(e); else notice(what + ": sign-in is needed.");
      return;
    }
    if (current) renderError(friendlyError(e), retry);
    else notice(what + ": " + friendlyError(e));
  }

  function notice(text) {
    state.notice = text;
    console.warn("Smart File: " + text);
    renderFooter();
  }

  /** The view to return to after a failed action: the options list, or the create form when there were no folders. */
  function backTo(session) {
    if (session.ranking) renderOptions(session); else renderNoFolders(session);
  }

  async function ensureToken() {
    return Auth.getToken();
  }

  async function signIn() {
    renderLoading("Opening sign-in…");
    try {
      await Auth.getTokenInteractive();
      rerun();
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
      newFolderForm(session, session.ctx.fromName, null),
      refreshLink()
    );
  }

  function renderSuggestion(session, timing) {
    const r = session.ranking;
    const top = r.top;
    if (!top) { renderOptions(session); return; }
    const yes = button("Yes, file it", () => fileTo(session, top.folder), "primary");
    setView(
      emailCard(session.ctx),
      h("p", { class: "question", text: r.confident ? "File this email in" : "No strong match. Best guess:" }),
      folderBlock(top.folder, top.reasons, true),
      h("div", { class: "actions" },
        yes,
        button("No, show other options", () => { session.rejected = top.folder; renderOptions(session); }, "secondary")
      ),
      cfg.showTimings && timing ? h("p", { class: "muted timing", text: timing.summary() }) : null
    );
    yes.focus();
  }

  function renderOptions(session) {
    const r = session.ranking;
    // Everything still ranked except the folder the user said no to (dropFolder may have promoted a new top).
    const shown = r.ranked.filter((x) => x.folder !== session.rejected).slice(0, 3);

    const parentDefault = defaultParentFor(session);
    const list = h("div", { class: "option-list" }, shown.map((alt) =>
      h("button", { class: "option", type: "button", onclick: () => fileTo(session, alt.folder) }, folderBlock(alt.folder, alt.reasons, false))
    ));

    setView(
      emailCard(session.ctx),
      h("p", { class: "question", text: shown.length ? "Other likely folders" : "No other likely folders" }),
      list,
      h("p", { class: "question", text: "Or create a new folder" }),
      newFolderForm(session, r.suggestedNewName, parentDefault),
      browseAll(session),
      refreshLink()
    );
  }

  async function refreshFoldersNow() {
    const runId = state.runId;
    renderLoading("Refreshing your folders…");
    try {
      const token = (await ensureToken()).accessToken;
      const result = await refreshFolderTree(token);
      const names = result && result.problems ? result.problems.slice(0, 5).join(", ") : "";
      if (state.runId !== runId) {                        // the user moved on to another email meanwhile
        if (result && !result.complete) notice("Folder list refreshed with gaps" + (names ? " (" + names + ")" : "") + ".");
        return;
      }
      if (result && !result.complete) {
        setView(h("div", { class: "status" },
          h("p", { text: "Some folders couldn’t be read just now" + (names ? " (" + names + ")" : "") + ", so the list may still be missing something." }),
          h("div", { class: "actions" }, button("Try again", refreshFoldersNow, "primary"), button("Continue anyway", rerun, "secondary"))
        ));
        return;
      }
      rerun();
    } catch (e) {
      if (state.runId !== runId) { notice("Couldn’t refresh the folder list: " + friendlyError(e)); return; }
      if (e instanceof Auth.AuthRequired) renderSignIn(e);
      else renderError(friendlyError(e), refreshFoldersNow);
    }
  }

  function refreshLink() {
    return h("p", { class: "muted hint" }, "Missing a folder you made recently? ", h("button", { class: "link", type: "button", onclick: refreshFoldersNow }, "Refresh the folder list"));
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
      h("div", { class: "check", "aria-hidden": "true" }, "✓"),
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
    const parts = [];
    if (state.notice) {
      parts.push(
        h("span", { class: "footer-text", text: state.notice }),
        h("button", { class: "link", type: "button", onclick: () => { state.notice = null; renderFooter(); } }, "Dismiss")
      );
    } else if (lm && !showing) {
      parts.push(
        h("span", { class: "footer-text" }, "Filed “", h("span", { class: "footer-subject", text: lm.subject || "(no subject)" }), "” in ", h("strong", { text: lm.folder.name })),
        h("button", { class: "link", type: "button", onclick: undoLastMove }, "Undo")
      );
    }
    if (!parts.length) { footer.replaceChildren(); footer.hidden = true; return; }
    footer.hidden = false;
    footer.replaceChildren.apply(footer, parts);
  }

  function friendlyError(e) {
    if (e && e.name === "GraphError") {
      if (e.status === 401 || e.status === 403) return "Microsoft 365 refused the request (" + (e.code || e.status) + "). Check that the Entra app has the Mail.ReadWrite permission and that consent was granted.";
      if (e.status === 404) return "This email or the folder can’t be found any more. One of them may have been moved or deleted.";
      if (e.status === 0) return e.message;
      return "Microsoft 365 returned an error" + (e.code ? " (" + e.code + ")" : "") + ": " + e.message;
    }
    return (e && e.message) || String(e);
  }
})();
