/* Smart File - folder suggestion engine.
 * Pure functions, no Office or Graph calls, so it can be unit-tested in Node.
 */
(function (global) {
  "use strict";

  const cfg = () => global.SmartFileConfig;

  const STOPWORDS = new Set([
    // generic
    "the", "and", "for", "are", "but", "not", "you", "all", "any", "can", "had", "her", "was", "one",
    "our", "out", "get", "has", "him", "his", "how", "its", "let", "may", "new", "now", "old", "see",
    "she", "too", "use", "way", "who", "with", "from", "this", "that", "have", "will", "your", "they",
    "been", "were", "what", "when", "which", "their", "there", "these", "those", "then", "than", "into",
    "about", "also", "just", "like", "more", "some", "such", "only", "over", "very", "would", "could",
    "should", "here", "where", "does", "did", "done", "each", "both", "same", "other", "after", "before",
    // email noise
    "please", "thanks", "thank", "regards", "best", "kind", "hello", "dear", "sincerely", "cheers",
    "email", "mail", "message", "sent", "subject", "re", "fw", "fwd", "aw", "wg", "ext", "external",
    "automatic", "reply", "attached", "attachment", "attachments", "below", "above", "following",
    "per", "via", "regarding", "question", "questions", "need", "needs", "know", "today", "tomorrow",
    "week", "monday", "tuesday", "wednesday", "thursday", "friday", "morning", "afternoon", "com",
    "www", "http", "https", "mailto", "phone", "mobile", "cell", "fax", "tel", "inc", "llc", "ltd",
    "corp", "corporation", "company", "sender", "caution", "originated", "outside", "organization",
    "organisation", "links", "click", "confidential", "intended", "recipient", "disclaimer"
  ]);

  const SUBDOMAIN_NOISE = new Set([
    "mail", "email", "e", "em", "smtp", "mx", "corp", "us", "na", "eu", "uk", "ca", "global", "group",
    "notifications", "notify", "noreply", "no-reply", "alerts", "alert", "info", "news", "marketing",
    "support", "service", "services", "sales", "go", "app", "apps", "my", "portal", "secure", "web",
    "www", "m", "mta", "bounce", "bounces", "reply", "replies", "send", "mailer", "hello"
  ]);

  const SECOND_LEVEL_SUFFIXES = new Set(["co", "com", "org", "net", "gov", "edu", "ac", "ltd", "plc", "gob", "or", "ne", "go"]);

  function normalize(s) {
    return (s || "").toLowerCase().replace(/[^a-z0-9]+/g, "");
  }

  function stem(t) {
    if (t.length > 3 && t.endsWith("s") && !t.endsWith("ss")) return t.slice(0, -1);
    return t;
  }

  function tokenize(s) {
    const seen = new Set();
    const out = [];
    (s || "").toLowerCase()
      .replace(/[\u2018\u2019']/g, "")
      .split(/[^a-z0-9]+/)
      .forEach((raw) => {
        if (raw.length < 3 || STOPWORDS.has(raw) || /^\d{1,2}$/.test(raw)) return;
        const t = stem(raw);
        if (!seen.has(t)) { seen.add(t); out.push(t); }
      });
    return out;
  }

  function bigrams(s) {
    const out = [];
    for (let i = 0; i < s.length - 1; i++) out.push(s.slice(i, i + 2));
    return out;
  }

  // Dice coefficient on character bigrams - cheap fuzzy match for near-identical names.
  function similarity(a, b) {
    a = normalize(a); b = normalize(b);
    if (!a || !b) return 0;
    if (a === b) return 1;
    const ba = bigrams(a), bb = bigrams(b);
    if (!ba.length || !bb.length) return 0;
    const counts = new Map();
    ba.forEach((g) => counts.set(g, (counts.get(g) || 0) + 1));
    let hits = 0;
    bb.forEach((g) => { const c = counts.get(g); if (c) { hits++; counts.set(g, c - 1); } });
    return (2 * hits) / (ba.length + bb.length);
  }

  function titleCase(s) {
    return (s || "").split(/\s+/).filter(Boolean)
      .map((w) => w[0].toUpperCase() + w.slice(1)).join(" ");
  }

  function domainOf(address) {
    const i = (address || "").lastIndexOf("@");
    return i < 0 ? "" : address.slice(i + 1).toLowerCase().trim();
  }

  function isPublicDomain(domain) {
    return !!domain && cfg().publicEmailDomains.includes(domain);
  }

  // "mail.rockwellautomation.com" -> { label: "rockwellautomation", tokens: [...], display: "Rockwellautomation" }
  function orgFromDomain(domain) {
    if (!domain || isPublicDomain(domain)) return null;
    const aliases = cfg().domainAliases || {};
    if (aliases[domain]) {
      const name = aliases[domain];
      return { label: normalize(name), tokens: tokenize(name), display: name, aliased: true };
    }
    const parts = domain.split(".").filter(Boolean);
    if (parts.length < 2) return null;
    let labels = parts.slice(0, -1);
    const last = labels[labels.length - 1];
    if (labels.length >= 2 && SECOND_LEVEL_SUFFIXES.has(last)) labels = labels.slice(0, -1);
    while (labels.length > 1 && SUBDOMAIN_NOISE.has(labels[0])) labels.shift();
    const label = labels[labels.length - 1];
    if (!label) return null;
    const words = label.split(/[-_]+/).filter(Boolean);
    return {
      label: normalize(label),
      tokens: words.map(stem).filter((w) => w.length >= 3),
      display: titleCase(words.join(" ")),
      aliased: false
    };
  }

  // Count how many words of `tokens` appear in `folderTokens`.
  function overlap(tokens, folderTokens) {
    const set = new Set(folderTokens);
    return tokens.filter((t) => set.has(t));
  }

  function orgMatch(org, folderNorm) {
    if (!org || !org.label) return 0;
    const label = org.label;
    if (label.length >= 4 && (folderNorm === label || folderNorm.includes(label) || (folderNorm.length >= 4 && label.includes(folderNorm)))) return 2;
    if (label.length === 3 && folderNorm.startsWith(label)) return 1;
    return 0;
  }

  /**
   * Prepare a folder for scoring. `folder` is {id, name, path, pathParts, totalItemCount, ...}.
   */
  function prepareFolder(folder) {
    const parents = (folder.pathParts || []).slice(0, -1);
    return Object.assign({}, folder, {
      norm: normalize(folder.name),
      tokens: tokenize(folder.name),
      pathTokens: tokenize(parents.join(" "))
    });
  }

  /**
   * Rank candidate folders for an email.
   * ctx:     { subject, fromName, fromAddress, recipients: [{name,address}], userAddress, bodyText, counterparty: {address,name} }
   * history: { sender: {folderId: n}, senderTotal, conversation: {folderId: n}, domain: {folderId: n} }
   * rules:   { bySender: { address: { f: folderId, t: timestamp } } }
   */
  function rankFolders(folders, ctx, history, rules) {
    const w = cfg().weights;
    history = history || { sender: {}, senderTotal: 0, conversation: {}, domain: {} };
    rules = rules || { bySender: {} };

    const userDomain = domainOf(ctx.userAddress);
    const cp = ctx.counterparty || { address: ctx.fromAddress, name: ctx.fromName };
    const cpAddress = (cp.address || "").toLowerCase();
    const cpDomain = domainOf(cpAddress);
    const org = orgFromDomain(cpDomain);
    const cpNameTokens = tokenize(cp.name);
    const subjectTokens = tokenize(ctx.subject);
    const bodyTokens = new Set(tokenize(ctx.bodyText));

    // Other external companies on the thread (customer CC'd, vendor rep, etc.)
    const otherOrgs = [];
    const seenDomains = new Set([cpDomain]);
    (ctx.recipients || []).concat([{ address: ctx.fromAddress }]).forEach((r) => {
      const d = domainOf((r.address || "").toLowerCase());
      if (!d || seenDomains.has(d) || d === userDomain) return;
      seenDomains.add(d);
      const o = orgFromDomain(d);
      if (o) otherOrgs.push(o);
    });

    // Learned rules
    const ruleEntry = rules.bySender[cpAddress];
    const learnedFolder = ruleEntry ? ruleEntry.f : null;
    const domainRuleCounts = {};
    if (cpDomain && !isPublicDomain(cpDomain)) {
      Object.entries(rules.bySender).forEach(([addr, entry]) => {
        if (domainOf(addr) === cpDomain && entry && entry.f) domainRuleCounts[entry.f] = (domainRuleCounts[entry.f] || 0) + 1;
      });
    }

    const results = folders.map((raw) => {
      const f = raw.tokens ? raw : prepareFolder(raw);
      let score = 0;
      const reasons = [];

      if (learnedFolder && learnedFolder === f.id) {
        score += w.learnedSender;
        reasons.push("You filed email from " + cpAddress + " here before");
      } else if (domainRuleCounts[f.id]) {
        score += w.learnedDomain;
        reasons.push("You filed email from @" + cpDomain + " here before");
      }

      const conv = history.conversation[f.id] || 0;
      if (conv) {
        score += w.conversation;
        reasons.push(conv === 1 ? "Another email in this conversation is here" : conv + " other emails in this conversation are here");
      }

      const sh = history.sender[f.id] || 0;
      if (sh) {
        const share = history.senderTotal ? sh / history.senderTotal : 1;
        score += Math.round(w.senderHistory * (0.4 + 0.6 * share));
        reasons.push(sh + " earlier email" + (sh > 1 ? "s" : "") + " from " + cpAddress + (sh > 1 ? " are" : " is") + " here");
      }

      const dh = history.domain[f.id] || 0;
      if (dh) {
        score += sh ? Math.round(w.domainHistory / 3) : w.domainHistory;
        if (!sh) reasons.push(dh + " email" + (dh > 1 ? "s" : "") + " from @" + cpDomain + (dh > 1 ? " are" : " is") + " here");
      }

      if (org) {
        const m = orgMatch(org, f.norm);
        if (m === 2) {
          score += w.orgName;
          reasons.push("Folder name matches " + (org.aliased ? org.display : cpDomain));
        } else if (m === 1) {
          score += w.orgPrefix;
          reasons.push("Folder name starts like " + cpDomain);
        } else {
          const hits = overlap(org.tokens, f.tokens);
          if (hits.length) { score += Math.min(40, 20 * hits.length); reasons.push("Folder name shares \u201c" + hits[0] + "\u201d with " + cpDomain); }
        }
      }

      const nameHits = overlap(cpNameTokens, f.tokens);
      if (nameHits.length) {
        score += Math.min(40, w.senderName * nameHits.length);
        reasons.push("Folder name matches \u201c" + nameHits[0] + "\u201d in the sender\u2019s name");
      }

      let recipientScore = 0;
      otherOrgs.forEach((o) => {
        if (orgMatch(o, f.norm) === 2 || overlap(o.tokens, f.tokens).length) recipientScore += w.recipientOrg;
      });
      if (recipientScore) { score += Math.min(35, recipientScore); reasons.push("Matches another company on this thread"); }

      const subjectHits = overlap(subjectTokens, f.tokens);
      let subjectScore = subjectHits.length * w.subjectToken;
      if (!subjectHits.length) {
        // fuzzy: "caterpiller" vs "caterpillar"
        const fuzzy = subjectTokens.some((t) => t.length >= 5 && f.tokens.some((ft) => ft.length >= 5 && similarity(t, ft) >= 0.8));
        if (fuzzy) subjectScore += 10;
      }
      if (subjectScore) { score += Math.min(w.subjectCap, subjectScore); reasons.push("Subject mentions \u201c" + (subjectHits[0] || "a similar word") + "\u201d"); }

      const bodyHits = f.tokens.filter((t) => bodyTokens.has(t));
      if (bodyHits.length) { score += Math.min(w.bodyCap, bodyHits.length * w.bodyToken); if (!subjectHits.length) reasons.push("Email text mentions \u201c" + bodyHits[0] + "\u201d"); }

      const pathHits = overlap(subjectTokens, f.pathTokens);
      if (pathHits.length) score += Math.min(w.pathCap, pathHits.length * w.pathToken);

      // Tiny prior: well-used folders win ties.
      score += Math.min(4, Math.log10((f.totalItemCount || 0) + 1));

      return { folder: f, score: Math.round(score * 10) / 10, reasons };
    });

    results.sort((a, b) => b.score - a.score || a.folder.path.localeCompare(b.folder.path));

    const top = results[0] || null;
    const suggestedNewName = org ? org.display : (cp.name ? titleCase(cp.name) : "");
    return {
      ranked: results,
      top: top,
      alternatives: results.slice(1, 4),
      confident: !!top && top.score >= w.confidentThreshold,
      suggestedNewName: suggestedNewName,
      counterparty: { address: cpAddress, domain: cpDomain, org: org }
    };
  }

  global.SmartFileSuggest = {
    rankFolders, prepareFolder, tokenize, normalize, similarity, domainOf, isPublicDomain, orgFromDomain, titleCase
  };
})(typeof window !== "undefined" ? window : globalThis);
