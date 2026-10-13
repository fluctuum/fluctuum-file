/* Fluctuum File: task pane
 * Files Outlook emails into FCD › Projects/<year>/<code name>/07 Comms and logs them
 * in the Correspondence list. In compose mode it tags the draft with the project so the
 * "FCD sent mail" flow files it after sending.
 */
(function () {
  "use strict";

  const CFG = window.FCD_CONFIG;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["User.Read", "Mail.ReadWrite", "Sites.ReadWrite.All", "MailboxSettings.ReadWrite"];
  const CACHE_KEY = "fcd.projects.v3";
  const CACHE_HOURS = 6;
  const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024 - 1024;
  const CHUNK = 320 * 1024 * 10; // 3.2 MB, multiple of 320 KiB as Graph requires

  let msalApp = null;
  let projects = [];          // {code, name, year, path}
  let mode = "read";          // read | compose
  let selection = [];         // [{itemId, subject, conversationId, internetMessageId}]
  let busy = false;
  let categoriesChecked = false;

  const $ = (id) => document.getElementById(id);

  /* ---------- small helpers ---------- */

  function show(el, on) { el.classList.toggle("hidden", !on); }

  function status(html, isError) {
    const el = $("status");
    el.innerHTML = html;
    el.classList.toggle("error", !!isError);
    show(el, !!html);
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  // Office.js may return conversation ids in EWS alphabet; Graph uses the URL safe one.
  function normConv(id) { return id ? String(id).replace(/\//g, "-").replace(/\+/g, "_") : ""; }

  function odataQuote(s) { return "'" + String(s).replace(/'/g, "''") + "'"; }

  function encPath(p) { return p.split("/").map(encodeURIComponent).join("/"); }

  function cleanName(s, max) {
    let t = String(s || "")
      .normalize("NFD").replace(/[̀-ͯ]/g, "")
      .replace(/^\s*((re|fw|fwd|rv|res|enc)\s*:\s*)+/i, "")
      .replace(/["*:<>?\/\\|#%~&{}\[\]\t\r\n]+/g, " ")
      .replace(/[^\x20-\x7E]+/g, "")
      .replace(/\s+/g, " ").trim();
    if (t.length > max) t = t.slice(0, max).trim();
    return t.replace(/[. ]+$/, "") || "email";
  }

  function yymmdd(iso) {
    const d = iso ? new Date(iso) : new Date();
    const p = (n) => String(n).padStart(2, "0");
    return p(d.getFullYear() % 100) + p(d.getMonth() + 1) + p(d.getDate());
  }

  function officeAsync(fn) {
    return new Promise((resolve, reject) => fn((r) => r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)));
  }

  function lsGet(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* ignore */ } }

  /* ---------- auth + Graph ---------- */

  async function getToken() {
    if (!msalApp) {
      if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
        throw new Error("This Outlook version does not support Nested App Authentication. Update Outlook.");
      }
      msalApp = await msal.createNestablePublicClientApplication({
        auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/" + CFG.tenantId },
        cache: { cacheLocation: "localStorage" }
      });
    }
    const req = { scopes: SCOPES };
    try {
      return (await msalApp.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError || /interaction|consent|login/i.test(e.message || "")) {
        return (await msalApp.acquireTokenPopup(req)).accessToken;
      }
      throw e;
    }
  }

  async function graph(path, opts) {
    opts = opts || {};
    const url = path.startsWith("http") ? path : GRAPH + path;
    for (let attempt = 0; attempt < 4; attempt++) {
      const token = await getToken();
      const headers = Object.assign({ Authorization: "Bearer " + token }, opts.headers || {});
      if (opts.json !== undefined) headers["Content-Type"] = "application/json";
      const res = await fetch(url, {
        method: opts.method || "GET",
        headers,
        body: opts.json !== undefined ? JSON.stringify(opts.json) : opts.body
      });
      if (res.status === 429 || res.status === 503) {
        const wait = Number(res.headers.get("Retry-After") || 2) * 1000;
        await new Promise((r) => setTimeout(r, Math.min(wait, 15000)));
        continue;
      }
      if (!res.ok) {
        let msg = res.status + " " + res.statusText;
        try { const j = await res.json(); msg = (j.error && j.error.message) || msg; } catch (e) { /* ignore */ }
        const err = new Error(msg); err.status = res.status; throw err;
      }
      if (opts.raw) return res;
      if (res.status === 204) return null;
      return res.json();
    }
    throw new Error("Microsoft Graph is busy, try again in a minute.");
  }

  async function graphAll(path) {
    let out = [], next = path;
    while (next) {
      const j = await graph(next);
      out = out.concat(j.value || []);
      next = j["@odata.nextLink"] || null;
    }
    return out;
  }

  /* ---------- projects ---------- */

  // Sources listed in the picker. Projects: <root>/<year>/<AAxxxx Name>. Proposals: Proposals/<year>/<PAAxxxx_Name>.
  const SOURCES = [
    { kind: "project", root: CFG.projectsRoot, re: /^(\d{6})\s*[-_ ]?\s*(.*)$/ },
    { kind: "proposal", root: CFG.proposalsRoot || "Proposals", re: /^(P\d{6})\s*[-_ ]?\s*(.*)$/i }
  ];

  async function listSource(src) {
    let years;
    try {
      years = (await graphAll(`/drives/${CFG.driveId}/root:/${encPath(src.root)}:/children?$select=name,folder&$top=999`))
        .filter((f) => f.folder && /^\d{4}$/.test(f.name)).map((f) => f.name);
    } catch (e) { if (e.status === 404) return []; throw e; }
    const perYear = await Promise.all(years.map((y) =>
      graphAll(`/drives/${CFG.driveId}/root:/${encPath(src.root + "/" + y)}:/children?$select=name,folder&$top=999`)
        .then((items) => items.filter((f) => f.folder).map((f) => {
          const m = src.re.exec(f.name);
          return m ? { code: m[1].toUpperCase(), name: m[2] || "", year: y, kind: src.kind, path: `${src.root}/${y}/${f.name}` } : null;
        }).filter(Boolean))
    ));
    return [].concat.apply([], perYear);
  }

  async function loadProjects(force) {
    const cached = lsGet(CACHE_KEY);
    if (!force && cached && Date.now() - cached.ts < CACHE_HOURS * 3600e3 && cached.list.length) {
      projects = cached.list;
      return;
    }
    $("listInfo").textContent = "Loading projects and proposals from FCD…";
    const all = [].concat.apply([], await Promise.all(SOURCES.map(listSource)));
    const seen = {};
    projects = all.filter((p) => (seen[p.code] ? false : (seen[p.code] = true)))
      .sort((a, b) => (b.year + b.code.replace(/^P/, "")).localeCompare(a.year + a.code.replace(/^P/, "")) || a.kind.localeCompare(b.kind));
    lsSet(CACHE_KEY, { ts: Date.now(), list: projects });
    saveCodesToRoaming();
  }

  // The send handler (launchevent.js) has no Graph access; it reads codes and known threads from roaming settings.
  function saveCodesToRoaming() {
    try {
      const rs = Office.context.roamingSettings;
      rs.set("codes", projects.map((p) => p.code).join(","));
      // Folder paths for recent projects and all proposals, so the send handler can tag the email with its folder
      const minYear = String(new Date().getFullYear() - 2);
      const paths = {};
      projects.filter((p) => p.kind === "proposal" || p.year >= minYear).forEach((p) => { paths[p.code] = p.path; });
      rs.set("paths", paths);
      rs.saveAsync(() => {});
    } catch (e) { /* ignore */ }
  }

  function rememberThread(convId, code) {
    if (!convId || !code) return;
    try {
      const rs = Office.context.roamingSettings;
      const key = normConv(convId).slice(-40);
      let map = rs.get("conv") || {};
      delete map[key];
      map[key] = code;
      const keys = Object.keys(map);
      if (keys.length > 150) keys.slice(0, keys.length - 150).forEach((k) => delete map[k]);
      rs.set("conv", map);
      rs.saveAsync(() => {});
    } catch (e) { /* ignore */ }
  }

  function findProject(code) { return projects.find((p) => p.code === code) || null; }

  let kindFilter = "all";
  function renderList() {
    const fold = (t) => t.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const q = fold($("search").value.trim());
    const ul = $("projectList");
    const hits = projects.filter((p) => (kindFilter === "all" || p.kind === kindFilter)
      && (!q || fold(p.code + " " + p.name).includes(q)));
    const shown = hits.slice(0, 80);
    let html = "", lastHead = null;
    shown.forEach((p) => {
      const head = p.year + (p.kind === "proposal" ? " · Proposals" : " · Projects");
      if (head !== lastHead) { html += `<li class="year">${esc(head)}</li>`; lastHead = head; }
      html += `<li><button class="project" data-code="${esc(p.code)}"><span class="code">${esc(p.code)}</span>${esc(p.name)}</button></li>`;
    });
    ul.innerHTML = html || `<li class="muted small">Nothing matches.</li>`;
    const nP = projects.filter((p) => p.kind === "project").length, nQ = projects.length - nP;
    $("listInfo").textContent = hits.length > shown.length
      ? `Showing ${shown.length} of ${hits.length}. Type to narrow down.`
      : `${nP} projects, ${nQ} proposals in FCD.`;
    document.querySelectorAll(".chip").forEach((c) => c.classList.toggle("on", c.dataset.kind === kindFilter));
  }

  /* ---------- Correspondence list ---------- */

  async function corrQuery(field, value, top) {
    const filter = encodeURIComponent(`fields/${field} eq ${odataQuote(value)}`);
    const j = await graph(`/sites/${CFG.siteId}/lists/${CFG.listId}/items?$filter=${filter}&$expand=fields($select=Project,ProjectName,InternetMessageId,FileUrl,Created)&$top=${top || 50}`,
      { headers: { Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly" } });
    return (j.value || []).map((i) => i.fields);
  }

  function mostCommon(rows) {
    const n = {};
    rows.forEach((r) => { if (r.Project && r.Project !== "NONE") n[r.Project] = (n[r.Project] || 0) + 1; });
    return Object.keys(n).sort((a, b) => n[b] - n[a])[0] || null;
  }

  /* ---------- suggestion ---------- */

  async function suggest(info) {
    // info: {conversationId, internetMessageId, subject, fromAddress}
    let filedIn = [], code = null, why = "";
    if (info.conversationId) {
      try {
        const rows = await corrQuery("ConversationId", normConv(info.conversationId));
        if (info.internetMessageId) {
          filedIn = rows.filter((r) => r.InternetMessageId === info.internetMessageId);
        }
        const c = mostCommon(rows);
        if (c) { code = c; why = "This thread is already filed there."; }
      } catch (e) { console.warn(e); }
    }
    if (!code && info.subject) {
      const m = (String(info.subject).match(/\bP?\d{6}\b/gi) || []).map((c) => c.toUpperCase());
      const hit = m.find((c) => findProject(c));
      if (hit) { code = hit; why = "Project code found in the subject."; }
    }
    if (!code && info.fromAddress) {
      const dom = info.fromAddress.split("@")[1];
      if (dom && dom.toLowerCase() !== CFG.internalDomain) {
        try {
          const c = mostCommon(await corrQuery("FromDomain", dom.toLowerCase(), 50));
          if (c) { code = c; why = `Most emails from ${dom} are filed there.`; }
        } catch (e) { console.warn(e); }
      }
    }
    return { code, why, filedIn };
  }

  function showSuggestion(s) {
    const p = s.code && findProject(s.code);
    show($("suggestBox"), !!p);
    if (p) {
      $("suggestBtn").dataset.code = p.code;
      $("suggestBtn").innerHTML = `<span class="code">${esc(p.code)}</span>${esc(p.name)}`;
      $("suggestWhy").textContent = s.why;
    }
    const filed = $("filedStatus");
    if (s.filedIn && s.filedIn.length) {
      filed.innerHTML = "Already filed in " + s.filedIn.map((r) => {
        const pr = findProject(r.Project);
        const label = esc(r.Project + (pr ? " " + pr.name : ""));
        return r.FileUrl ? `<a href="${esc(r.FileUrl)}" target="_blank">${label}</a>` : label;
      }).join(", ");
      show(filed, true);
    } else show(filed, false);
  }

  /* ---------- read mode: filing ---------- */

  async function ensureCategories() {
    if (categoriesChecked) return;
    categoriesChecked = true;
    try {
      const cats = (await graph("/me/outlook/masterCategories")).value.map((c) => c.displayName);
      if (!cats.includes(CFG.filedCategory)) {
        await graph("/me/outlook/masterCategories", { method: "POST", json: { displayName: CFG.filedCategory, color: "preset1" } });
      }
      if (!cats.includes(CFG.todoCategory)) {
        await graph("/me/outlook/masterCategories", { method: "POST", json: { displayName: CFG.todoCategory, color: "preset14" } });
      }
    } catch (e) { console.warn("Categories", e); }
  }

  async function uploadEml(path, blob) {
    if (blob.size <= SIMPLE_UPLOAD_MAX) {
      return graph(`/drives/${CFG.driveId}/root:/${encPath(path)}:/content?@microsoft.graph.conflictBehavior=rename`,
        { method: "PUT", body: blob, headers: { "Content-Type": "message/rfc822" } });
    }
    const session = await graph(`/drives/${CFG.driveId}/root:/${encPath(path)}:/createUploadSession`,
      { method: "POST", json: { item: { "@microsoft.graph.conflictBehavior": "rename" } } });
    let result = null;
    for (let start = 0; start < blob.size; start += CHUNK) {
      const end = Math.min(start + CHUNK, blob.size);
      const res = await fetch(session.uploadUrl, {
        method: "PUT",
        headers: { "Content-Range": `bytes ${start}-${end - 1}/${blob.size}` },
        body: blob.slice(start, end)
      });
      if (!res.ok && res.status !== 202) throw new Error("Upload failed (" + res.status + ")");
      if (res.status === 200 || res.status === 201) result = await res.json();
    }
    return result;
  }

  // Old projects use other names for the email folder ("Emails o correspondencia", "01 Comms"...).
  // Use the existing one; only create CFG.commsFolder when the project has none.
  const commsCache = {};
  async function commsFolderFor(project) {
    if (commsCache[project.path]) return commsCache[project.path];
    const def = project.kind === "proposal" ? (CFG.proposalCommsFolder || "04 Com") : CFG.commsFolder;
    let name = null;
    try {
      const kids = (await graphAll(`/drives/${CFG.driveId}/root:/${encPath(project.path)}:/children?$select=name,folder&$top=999`))
        .filter((f) => f.folder).map((f) => f.name);
      const norm = (n) => n.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
      name = kids.find((n) => norm(n) === norm(def))
        || kids.find((n) => /^\d{1,2}\s*[-_.]?\s*com(ms|unicaciones|munications?)?\b/i.test(n))
        || kids.find((n) => /(correspondencia|correspondence|e-?mails?|correos?|comms|communications?|comunicaciones)/i.test(norm(n)))
        || null;
    } catch (e) { console.warn("comms lookup", e); }
    const path = `${project.path}/${name || def}`;
    commsCache[project.path] = path;
    try {
      const rs = Office.context.roamingSettings;
      const cp = rs.get("cpath") || {};
      delete cp[project.code]; cp[project.code] = path;
      const ks = Object.keys(cp);
      if (ks.length > 80) ks.slice(0, ks.length - 80).forEach((k) => delete cp[k]);
      rs.set("cpath", cp); rs.saveAsync(() => {});
    } catch (e) { /* ignore */ }
    return path;
  }

  async function fileMessage(restId, project) {
    const me = (Office.context.mailbox.userProfile.emailAddress || "").toLowerCase();
    const msg = await graph(`/me/messages/${encodeURIComponent(restId)}?$select=subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,conversationId,internetMessageId,hasAttachments,categories`);
    const imid = (msg.internetMessageId || "").slice(0, 255);

    // Skip the upload if this exact email is already in this project
    let existing = [];
    if (imid) { try { existing = await corrQuery("InternetMessageId", imid, 20); } catch (e) { /* ignore */ } }
    const dup = existing.find((r) => r.Project === project.code);

    let webUrl = dup ? dup.FileUrl : null;
    if (!dup) {
      const fromAddr = ((msg.from && msg.from.emailAddress && msg.from.emailAddress.address) || "").toLowerCase();
      const out = fromAddr === me;
      const counterpart = out
        ? ((msg.toRecipients[0] && (msg.toRecipients[0].emailAddress.name || msg.toRecipients[0].emailAddress.address)) || "")
        : ((msg.from && (msg.from.emailAddress.name || msg.from.emailAddress.address)) || "");
      const when = out ? msg.sentDateTime : msg.receivedDateTime;
      const fileName = `${yymmdd(when)}_${cleanName(counterpart, 25)}_${cleanName(msg.subject, 60)}.eml`;

      const mimeRes = await graph(`/me/messages/${encodeURIComponent(restId)}/$value`, { raw: true });
      const blob = await mimeRes.blob();
      const item = await uploadEml(`${await commsFolderFor(project)}/${fileName}`, blob);
      webUrl = item && item.webUrl;

      const to = (msg.toRecipients || []).concat(msg.ccRecipients || [])
        .map((r) => r.emailAddress.address).join("; ");
      await graph(`/sites/${CFG.siteId}/lists/${CFG.listId}/items`, {
        method: "POST",
        json: {
          fields: {
            Title: (msg.subject || "(no subject)").slice(0, 255),
            Project: project.code,
            ProjectName: (project.name || "").slice(0, 255),
            Direction: out ? "Out" : "In",
            FromAddress: fromAddr.slice(0, 255),
            FromDomain: (fromAddr.split("@")[1] || "").slice(0, 255),
            ToAddresses: to,
            MessageDate: when,
            ConversationId: normConv(msg.conversationId).slice(0, 255),
            InternetMessageId: imid,
            FileUrl: webUrl || "",
            FiledBy: me,
            Method: "Manual",
            HasAttachments: !!msg.hasAttachments
          }
        }
      });
    }

    // Tag it in Outlook
    const cats = (msg.categories || []).filter((c) => c !== CFG.todoCategory);
    if (!cats.includes(CFG.filedCategory)) cats.push(CFG.filedCategory);
    try {
      await graph(`/me/messages/${encodeURIComponent(restId)}`, { method: "PATCH", json: { categories: cats } });
    } catch (e) { console.warn("Category", e); }

    rememberThread(msg.conversationId, project.code);
    return { webUrl, dup: !!dup };
  }

  async function fileSelection(code) {
    const project = findProject(code);
    if (!project || busy) return;
    if (!selection.length) { status("Select an email first.", true); return; }
    busy = true;
    await ensureCategories();
    let ok = 0, failed = 0, lastUrl = null, dups = 0;
    for (let i = 0; i < selection.length; i++) {
      status(`Filing ${i + 1} of ${selection.length} in <b>${esc(project.code)}</b>…`);
      try {
        const restId = Office.context.mailbox.convertToRestId(selection[i].itemId, Office.MailboxEnums.RestVersion.v2_0);
        const r = await fileMessage(restId, project);
        ok++; lastUrl = r.webUrl || lastUrl; if (r.dup) dups++;
      } catch (e) {
        console.error(e); failed++;
        status(`Error on "${esc(selection[i].subject || "")}": ${esc(e.message)}`, true);
      }
    }
    busy = false;
    if (!failed) {
      const folderUrl = lastUrl ? lastUrl.replace(/\/[^\/]+$/, "") : null;
      status(`Filed ${ok} email${ok === 1 ? "" : "s"} in <b>${esc(project.code)} ${esc(project.name)}</b>` +
        (dups ? ` (${dups} already there)` : "") +
        (folderUrl ? `<br><a href="${esc(folderUrl)}" target="_blank">Open ${esc(decodeURIComponent(folderUrl.split("/").pop()))}</a>` : ""));
      refreshContext(true).catch(console.error);
    } else {
      status(`Filed ${ok}, failed ${failed}. Check the error above and try again.`, true);
    }
  }

  /* ---------- compose mode ---------- */

  async function composeSet(code) {
    const item = Office.context.mailbox.item;
    const props = await officeAsync((cb) => item.loadCustomPropertiesAsync(cb));
    if (code === "NONE") {
      props.set("project", "NONE");
      props.remove("projectName"); props.remove("folder"); props.remove("comms");
    } else {
      const p = findProject(code);
      props.set("project", p.code);
      props.set("projectName", p.name);
      props.set("folder", p.path);
      props.set("comms", await commsFolderFor(p));
    }
    await officeAsync((cb) => props.saveAsync(cb));
    if (code !== "NONE" && item.conversationId) rememberThread(item.conversationId, code);
    const p = findProject(code);
    const where = code === "NONE" ? "" : (p.kind === "proposal" ? "04 Com" : "07 Comms");
    status(code === "NONE"
      ? "<b>Ready to send.</b> This email will go out without a project."
      : `<b>Ready to send.</b> Project ${esc(p.code)} ${esc(p.name)}. It will be filed in ${where} after sending.`);
    showComposeState(code);
    // Banner inside the email, visible even after closing the pane
    try {
      const msg = code === "NONE"
        ? "Fluctuum: no project. Ready to send."
        : `Fluctuum: ${p.code} ${p.name}. Ready to send, it will be filed after sending.`;
      item.notificationMessages.replaceAsync("fcdProject", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: msg.slice(0, 150), icon: "Icon.16", persistent: false
      }, () => {});
    } catch (e) { /* older client */ }
  }

  function showComposeState(code) {
    const filed = $("filedStatus");
    if (!code) { show(filed, false); return; }
    const p = findProject(code);
    filed.innerHTML = code === "NONE" ? "No project" : `Project: <b>${esc(code)}</b> ${esc(p ? p.name : "")}`;
    show(filed, true);
  }

  /* ---------- context ---------- */

  async function refreshContext(keepStatus) {
    const item = Office.context.mailbox.item;
    if (!keepStatus) status("");
    if (item && typeof item.subject === "object" && item.subject && typeof item.subject.getAsync === "function") {
      mode = "compose";
      $("modeLabel").textContent = "New email";
      show($("composeExtra"), true);
      const subject = await officeAsync((cb) => item.subject.getAsync(cb)).catch(() => "");
      $("itemSummary").textContent = subject || "(no subject yet)";
      const props = await officeAsync((cb) => item.loadCustomPropertiesAsync(cb)).catch(() => null);
      const current = props && props.get("project");
      showComposeState(current);
      if (!current) {
        const s = await suggest({ conversationId: item.conversationId, subject });
        showSuggestion({ code: s.code, why: s.why, filedIn: [] });
      } else show($("suggestBox"), false);
      return;
    }

    mode = "read";
    $("modeLabel").textContent = "Email";
    show($("composeExtra"), false);

    // Multi-select (Mailbox 1.13)
    let sel = [];
    if (Office.context.requirements.isSetSupported("Mailbox", "1.13") && Office.context.mailbox.getSelectedItemsAsync) {
      try { sel = await officeAsync((cb) => Office.context.mailbox.getSelectedItemsAsync(cb)); } catch (e) { sel = []; }
    }
    if (sel && sel.length > 1) {
      selection = sel.map((s) => ({ itemId: s.itemId, subject: s.subject, conversationId: s.conversationId, internetMessageId: s.internetMessageId }));
      $("itemSummary").textContent = `${selection.length} emails selected`;
      show($("filedStatus"), false);
      const s = await suggest({ conversationId: selection[0].conversationId, subject: selection[0].subject });
      showSuggestion({ code: s.code, why: s.why, filedIn: [] });
      return;
    }

    if (!item) {
      selection = [];
      $("itemSummary").textContent = "Select an email.";
      show($("suggestBox"), false); show($("filedStatus"), false);
      return;
    }
    selection = [{ itemId: item.itemId, subject: item.subject, conversationId: item.conversationId, internetMessageId: item.internetMessageId }];
    const from = item.from ? (item.from.displayName || item.from.emailAddress) : "";
    $("itemSummary").innerHTML = `${esc(item.subject || "(no subject)")}<div class="muted small">${esc(from)}</div>`;
    const s = await suggest({
      conversationId: item.conversationId,
      internetMessageId: item.internetMessageId,
      subject: item.subject,
      fromAddress: item.from && item.from.emailAddress
    });
    showSuggestion(s);
  }

  function onPick(code) {
    if (!code) return;
    if (mode === "compose") composeSet(code).catch((e) => status(esc(e.message), true));
    else fileSelection(code).catch((e) => { busy = false; status(esc(e.message), true); });
  }

  /* ---------- boot ---------- */

  Office.onReady(async () => {
    $("search").addEventListener("input", renderList);
    document.querySelectorAll(".chip").forEach((c) => c.addEventListener("click", () => { kindFilter = c.dataset.kind; renderList(); }));
    $("projectList").addEventListener("click", (ev) => {
      const b = ev.target.closest("button[data-code]");
      if (b) onPick(b.dataset.code);
    });
    $("suggestBtn").addEventListener("click", () => onPick($("suggestBtn").dataset.code));
    $("noProject").addEventListener("click", () => composeSet("NONE").catch((e) => status(esc(e.message), true)));
    $("refresh").addEventListener("click", async () => {
      try { await loadProjects(true); renderList(); refreshContext(); } catch (e) { status(esc(e.message), true); }
    });

    try {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => refreshContext().catch(console.error));
    } catch (e) { /* not pinnable here */ }
    try {
      if (Office.EventType.SelectedItemsChanged) {
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, () => refreshContext().catch(console.error));
      }
    } catch (e) { /* older client */ }

    try {
      await loadProjects(false);
      renderList();
      await refreshContext();
      $("search").focus();
    } catch (e) {
      console.error(e);
      status("Could not start: " + esc(e.message), true);
      $("itemSummary").textContent = "";
    }
  });
})();
