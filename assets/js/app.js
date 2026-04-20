/**
 * Dashboard behaviour: navigation, tables, filters, pricing calculator, mail draft.
 * Requires: index.html markup, data.js (MOCK_* globals), auth.js (optional for mail).
 */
(function () {
  const TITLES = {
    overview: "Overview",
    completed: "Completed",
    ongoing: "Ongoing",
    "fabric-quality": "Fabric quality selection",
    pricing: "Pricing",
    alerts: "Alerts",
    received: "Received",
    mail: "Mail",
  };

  let mailFolder = "inbox";
  let selectedMailId = null;
  let selectedDraftId = null;
  let selectedSentId = null;
  /** Session-only messages sent via Compose Send */
  let mailSentSession = [];
  /** Working copy of mock drafts; Save draft mutates this array */
  let mailDraftsRuntime = [];
  /** Expanded row id for Ongoing table (accordion) — stays on the row you opened. */
  let ongoingExpandedId = null;
  /** Which material line’s detail/strip selection is shown inside that drawer (same vendor as anchor). */
  let ongoingDrawerOrderId = null;
  /** Ongoing row id whose stage editor panel is open. */
  let ongoingEditOpenId = null;
  /** Fabric quality: warp / weft / review open only after that step’s Next is clicked (general → warp → weft → review). */
  let fqsRevealedWarp = false;
  let fqsRevealedWeft = false;
  let fqsRevealedReview = false;
  /** After General Next: general panel shows compact summary; warp is the active step. */
  let fqsGeneralCollapsed = false;

  const badgeClass = (progress) => {
    const p = (progress || "").toUpperCase();
    if (p.includes("ENQUIRY")) return "badge badge--enquiry";
    if (p.includes("SAMPLE")) return "badge badge--sample";
    if (p.includes("QUOTATION")) return "badge badge--quote";
    if (p.includes("PRICE CONFIRMATION")) return "badge badge--price";
    if (p.includes("ORDER CONFIRMATION")) return "badge badge--order";
    return "badge";
  };

  function parseISODate(s) {
    if (!s) return null;
    const d = new Date(s + "T12:00:00");
    return Number.isNaN(d.getTime()) ? null : d;
  }

  function inRange(dateStr, start, end) {
    const d = parseISODate(dateStr);
    if (!d) return false;
    return d >= start && d <= end;
  }

  function getGlobalRange() {
    const startEl = document.getElementById("global-start");
    const endEl = document.getElementById("global-end");
    const start = parseISODate(startEl.value) || new Date();
    const end = parseISODate(endEl.value) || new Date();
    if (start > end) return { start: end, end: start };
    const s = new Date(start);
    s.setHours(0, 0, 0, 0);
    const e = new Date(end);
    e.setHours(23, 59, 59, 999);
    return { start: s, end: e };
  }

  function isGlobalDateAll() {
    const root = document.getElementById("date-range-popover");
    return !!(root && root.dataset.dateMode === "all");
  }

  function setGlobalDateAll(on) {
    const root = document.getElementById("date-range-popover");
    if (!root) return;
    if (on) root.dataset.dateMode = "all";
    else delete root.dataset.dateMode;
  }

  /** When "All" is on, every row passes; otherwise same as inRange (missing dateStr still passes). */
  function rowMatchesGlobalDate(dateStr) {
    if (isGlobalDateAll()) return true;
    if (!dateStr) return true;
    const { start, end } = getGlobalRange();
    return inRange(dateStr, start, end);
  }

  function setDefaultDates() {
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - 29);
    document.getElementById("global-end").value = end.toISOString().slice(0, 10);
    document.getElementById("global-start").value = start.toISOString().slice(0, 10);
  }

  function updateDateLabel() {
    const el = document.getElementById("global-date-line");
    if (!el) return;
    if (isGlobalDateAll()) {
      el.textContent = "All dates";
      return;
    }
    const { start, end } = getGlobalRange();
    const fmt = (d) =>
      d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
    el.textContent = `${fmt(start)} → ${fmt(end)}`;
  }

  function toIsoDateLocal(d) {
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const day = d.getDate();
    return `${y}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  function applyDateRangePreset(preset) {
    const startEl = document.getElementById("global-start");
    const endEl = document.getElementById("global-end");
    if (!startEl || !endEl) return;
    if (preset === "all") {
      setGlobalDateAll(true);
      updateDateLabel();
      refreshAll();
      return;
    }
    setGlobalDateAll(false);
    const end = new Date();
    end.setHours(12, 0, 0, 0);
    let start = new Date(end);
    if (preset === "month") {
      start = new Date(end.getFullYear(), end.getMonth(), 1);
    } else {
      const days = Number(preset);
      if (!Number.isFinite(days) || days < 1) return;
      start.setDate(start.getDate() - (days - 1));
    }
    startEl.value = toIsoDateLocal(start);
    endEl.value = toIsoDateLocal(end);
    startEl.dispatchEvent(new Event("change", { bubbles: true }));
    endEl.dispatchEvent(new Event("change", { bubbles: true }));
  }

  /* --- Navigation --- */
  function showView(key) {
    if (key !== "mail") closeMailMenu();
    fabricQualityModalRestoreSectionToMain();
    document.querySelectorAll(".view").forEach((el) => {
      el.hidden = true;
      el.classList.remove("is-visible");
    });
    const section = document.getElementById(`view-${key}`);
    if (section) {
      section.hidden = false;
      section.classList.add("is-visible");
    }
    document.querySelectorAll(".nav-item").forEach((btn) => {
      btn.classList.toggle("is-active", btn.dataset.view === key);
    });
    const titleEl = document.getElementById("page-title");
    if (key === "mail") titleEl.textContent = mailFolderLabel(mailFolder);
    else titleEl.textContent = TITLES[key] || key;
    if (key === "pricing") prsPopulateFabricQualitySelect();
    updateDateLabel();
    const stageWrap = document.getElementById("topbar-stage-wrap");
    if (stageWrap) stageWrap.hidden = key !== "overview";
  }

  document.querySelectorAll(".nav-item").forEach((btn) => {
    if (btn.id === "nav-mail-trigger") return;
    btn.addEventListener("click", () => showView(btn.dataset.view));
  });

  /* --- Tables --- */
  function renderTableHead(table, columns) {
    const thead = table.querySelector("thead");
    thead.innerHTML = `<tr>${columns
      .map(
        (c) =>
          `<th${c.sortKey ? ` data-sort="${c.sortKey}"` : ""}>${c.label}</th>`
      )
      .join("")}</tr>`;
  }

  function renderTableBody(table, rows, columns, emptyMessage) {
    const tbody = table.querySelector("tbody");
    const emptyMsg = emptyMessage || "No rows in this range.";
    tbody.innerHTML = rows
      .map(
        (row) =>
          `<tr>${columns.map((c) => `<td>${c.render(row)}</td>`).join("")}</tr>`
      )
      .join("");
    if (rows.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${columns.length}" class="muted">${emptyMsg}</td></tr>`;
    }
  }

  const tableRowCache = {
    awaiting: [],
    completed: [],
    ongoing: [],
    received: [],
  };

  const tableColumns = {
    awaiting: null,
    completed: null,
    ongoing: null,
    received: null,
  };

  const AWAITING_COLS = [
    { label: "Order ID", sortKey: "id", render: (r) => r.id },
    { label: "Vendor name", sortKey: "vendor", render: (r) => r.vendor },
    { label: "Issue", sortKey: "issue", render: (r) => r.issue },
    {
      label: "Stage",
      sortKey: "stage",
      render: (r) => `<span class="${badgeClass(r.stage)}">${r.stage}</span>`,
    },
    {
      label: "Waiting on",
      sortKey: "waiting",
      render: (r) => (r.waiting === "elite" ? "Elite" : "Vendor"),
    },
    {
      label: "Actions",
      render: (r) =>
        `<button type="button" class="btn btn--ghost btn--small btn--cta" data-awaiting-detail="${escapeAttr(r.id)}">View details</button>`,
    },
  ];

  let overviewAwaitingBase = [];

  function sortRows(rows, key, dir) {
    return [...rows].sort((a, b) => {
      let va = a[key];
      let vb = b[key];
      if (typeof va === "string" && typeof vb === "string") {
        va = va.toLowerCase();
        vb = vb.toLowerCase();
      }
      if (va == null) va = "";
      if (vb == null) vb = "";
      let cmp = 0;
      if (typeof va === "number" && typeof vb === "number") cmp = va - vb;
      else cmp = String(va).localeCompare(String(vb));
      return dir === "asc" ? cmp : -cmp;
    });
  }

  function bindDelegatedSort(tableId, cacheKey) {
    const table = document.getElementById(tableId);
    table.addEventListener("click", (e) => {
      const th = e.target.closest("th[data-sort]");
      if (!th || !table.contains(th)) return;
      const key = th.dataset.sort;
      const current = table.dataset.sortKey;
      const dir = current === `${key}-asc` ? "desc" : "asc";
      table.dataset.sortKey = `${key}-${dir}`;
      table.querySelectorAll("th[data-sort]").forEach((h) => h.classList.remove("is-sorted"));
      th.classList.add("is-sorted");
      const cols = tableColumns[cacheKey];
      const sorted = sortRows(tableRowCache[cacheKey], key, dir);
      tableRowCache[cacheKey] = sorted;
      renderTableBody(table, sorted, cols);
    });
  }

  function fillOverviewFilterSelects() {
    const stageSel = document.getElementById("overview-filter-stage");
    const vendorSel = document.getElementById("overview-filter-vendor");
    if (!stageSel || !vendorSel || typeof PIPELINE_STAGES === "undefined") return;
    stageSel.innerHTML =
      `<option value="">All stages</option>` +
      PIPELINE_STAGES.map((s) => `<option value="${escapeAttr(s)}">${escapeHtml(s)}</option>`).join("");
    const vendors = [...new Set(MOCK_AWAITING_ELITE.map((a) => a.vendor))].sort();
    vendorSel.innerHTML =
      `<option value="">All vendors</option>` +
      vendors.map((v) => `<option value="${escapeAttr(v)}">${escapeHtml(v)}</option>`).join("");
  }

  function escapeAttr(s) {
    return String(s).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
  }

  const SIDEBAR_KEY = "elite-sidebar-collapsed";

  function closeMailMenu() {
    const panel = document.getElementById("nav-mail-menu");
    const trig = document.getElementById("nav-mail-trigger");
    const wrap = document.getElementById("nav-mail");
    if (!panel || !trig) return;
    panel.hidden = true;
    trig.setAttribute("aria-expanded", "false");
    wrap?.classList.remove("is-open");
  }

  function openMailMenu() {
    const panel = document.getElementById("nav-mail-menu");
    const trig = document.getElementById("nav-mail-trigger");
    const wrap = document.getElementById("nav-mail");
    if (!panel || !trig) return;
    panel.hidden = false;
    trig.setAttribute("aria-expanded", "true");
    wrap?.classList.add("is-open");
  }

  function mailFolderLabel(folder) {
    const labels = { inbox: "Inbox", compose: "Compose", sent: "Sent", draft: "Draft" };
    return `Mail · ${labels[folder] || folder}`;
  }

  function applySidebarCollapsed(collapsed) {
    const root = document.getElementById("app-root");
    const toggle = document.getElementById("sidebar-toggle");
    if (!root || !toggle) return;
    root.classList.toggle("app--sidebar-collapsed", collapsed);
    toggle.setAttribute("aria-expanded", collapsed ? "false" : "true");
    toggle.title = collapsed ? "Expand menu" : "Collapse menu";
    toggle.setAttribute(
      "aria-label",
      collapsed ? "Expand navigation menu" : "Collapse navigation menu"
    );
    closeMailMenu();
    try {
      localStorage.setItem(SIDEBAR_KEY, collapsed ? "1" : "0");
    } catch {
      /* ignore */
    }
  }

  function initDateRangePopover() {
    const root = document.getElementById("date-range-popover");
    const btn = document.getElementById("date-range-toggle");
    const panel = document.getElementById("date-range-panel");
    if (!root || !btn || !panel) return;

    function setOpen(open) {
      panel.hidden = !open;
      btn.setAttribute("aria-expanded", open ? "true" : "false");
      root.classList.toggle("is-open", open);
    }

    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      setOpen(panel.hidden);
    });

    document.addEventListener("click", () => setOpen(false));

    root.addEventListener("click", (e) => e.stopPropagation());

    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      if (panel.hidden) return;
      setOpen(false);
      btn.focus();
    });

    panel.querySelectorAll("[data-date-preset]").forEach((presetBtn) => {
      presetBtn.addEventListener("click", () => {
        applyDateRangePreset(presetBtn.getAttribute("data-date-preset"));
      });
    });
  }

  function initSidebarToggle() {
    const root = document.getElementById("app-root");
    const toggle = document.getElementById("sidebar-toggle");
    if (!root || !toggle) return;
    let collapsed = false;
    try {
      collapsed = localStorage.getItem(SIDEBAR_KEY) === "1";
    } catch {
      /* ignore */
    }
    const mq = window.matchMedia("(min-width: 769px)");
    const sync = () => {
      if (!mq.matches) {
        root.classList.remove("app--sidebar-collapsed");
        toggle.setAttribute("aria-expanded", "true");
        toggle.title = "Collapse menu";
        toggle.setAttribute("aria-label", "Collapse navigation menu");
        return;
      }
      applySidebarCollapsed(collapsed);
    };
    sync();
    mq.addEventListener("change", sync);
    toggle.addEventListener("click", () => {
      if (!mq.matches) return;
      collapsed = !root.classList.contains("app--sidebar-collapsed");
      applySidebarCollapsed(collapsed);
    });
  }

  function openAwaitingModal(row) {
    const modal = document.getElementById("awaiting-modal");
    if (!modal) return;
    document.getElementById("awaiting-modal-title").textContent = `${row.id} · ${row.vendor}`;
    document.getElementById("awaiting-modal-handler").innerHTML = `<strong>Handler</strong> · ${escapeHtml(row.handler || "—")}`;
    document.getElementById("awaiting-modal-issue").textContent = row.issue || "";

    const stages = typeof PIPELINE_STAGES !== "undefined" ? PIPELINE_STAGES : [];
    const currentIdx = stages.indexOf(row.stage);
    const safeIdx = currentIdx >= 0 ? currentIdx : 0;
    const dates = row.stepDates || {};
    const items = stages.map((name, i) => {
      let state = "upcoming";
      if (i < safeIdx) state = "done";
      else if (i === safeIdx) state = "current";
      const d = dates[name] ? parseISODate(dates[name]) : null;
      const dateStr =
        d && !Number.isNaN(d.getTime())
          ? d.toLocaleDateString(undefined, {
              month: "short",
              day: "numeric",
              year: "numeric",
            })
          : null;
      return { name, state, dateStr };
    });

    document.getElementById("awaiting-modal-timeline").innerHTML = items
      .map(
        (it) => `
      <li class="timeline__item timeline__item--${it.state}">
        <span class="timeline__dot" aria-hidden="true"></span>
        <div class="timeline__body">
          <span class="timeline__name">${escapeHtml(it.name)}</span>
          ${it.dateStr ? `<span class="timeline__date">${escapeHtml(it.dateStr)}</span>` : it.state === "current" ? `<span class="timeline__date timeline__date--now">In progress</span>` : ""}
        </div>
      </li>`
      )
      .join("");

    modal.hidden = false;
    modal.setAttribute("aria-hidden", "false");
    modal.classList.remove("is-open");
    requestAnimationFrame(() => {
      requestAnimationFrame(() => modal.classList.add("is-open"));
    });
    modal.querySelector(".modal__close")?.focus();
  }

  function closeAwaitingModal() {
    const modal = document.getElementById("awaiting-modal");
    if (!modal || modal.hidden) return;
    if (!modal.classList.contains("is-open")) {
      modal.hidden = true;
      modal.setAttribute("aria-hidden", "true");
      return;
    }
    modal.classList.remove("is-open");
    let finished = false;
    const done = () => {
      if (finished) return;
      finished = true;
      modal.hidden = true;
      modal.setAttribute("aria-hidden", "true");
    };
    modal.addEventListener(
      "transitionend",
      (e) => {
        if (e.target === modal && e.propertyName === "opacity") done();
      },
      { once: true }
    );
    window.setTimeout(done, 280);
  }

  function filterAwaitingBySearch(base) {
    const inp = document.getElementById("overview-awaiting-search");
    const q = inp && inp.value ? inp.value.trim().toLowerCase() : "";
    if (!q) return [...base];
    return base.filter((r) => {
      const blob = [r.id, r.vendor, r.issue, r.stage, r.handler || ""].join(" ").toLowerCase();
      return blob.includes(q);
    });
  }

  function parseAwaitingSortKey(sk) {
    if (!sk) return null;
    const last = sk.lastIndexOf("-");
    if (last <= 0) return null;
    const key = sk.slice(0, last);
    const dir = sk.slice(last + 1);
    if (dir !== "asc" && dir !== "desc") return null;
    return { key, dir };
  }

  function renderAwaitingTable(resetSort) {
    const t = document.getElementById("table-awaiting");
    if (!t) return;
    const cols = AWAITING_COLS;
    tableColumns.awaiting = cols;
    let rows = filterAwaitingBySearch(overviewAwaitingBase);
    if (resetSort) {
      t.dataset.sortKey = "";
    } else {
      const parsed = parseAwaitingSortKey(t.dataset.sortKey);
      if (parsed) rows = sortRows(rows, parsed.key, parsed.dir);
    }
    tableRowCache.awaiting = rows;
    const emptyAwaiting =
      rows.length === 0 && document.getElementById("overview-awaiting-search")?.value?.trim()
        ? "No orders match your search."
        : undefined;
    if (resetSort) {
      renderTableHead(t, cols);
      renderTableBody(t, rows, cols, emptyAwaiting);
      return;
    }
    renderTableBody(t, rows, cols, emptyAwaiting);
    const parsed = parseAwaitingSortKey(t.dataset.sortKey);
    t.querySelectorAll("th[data-sort]").forEach((h) => h.classList.remove("is-sorted"));
    if (parsed) t.querySelector(`th[data-sort="${parsed.key}"]`)?.classList.add("is-sorted");
  }

  /* --- Overview --- */
  function refreshOverview() {
    const awaitingFilter = document.getElementById("overview-awaiting").value;
    const stageFilter = document.getElementById("overview-filter-stage")?.value || "";
    const vendorFilter = document.getElementById("overview-filter-vendor")?.value || "";

    let awaiting = MOCK_AWAITING_ELITE.filter((a) => rowMatchesGlobalDate(a.opened));
    if (stageFilter) awaiting = awaiting.filter((a) => a.stage === stageFilter);
    if (vendorFilter) awaiting = awaiting.filter((a) => a.vendor === vendorFilter);
    if (awaitingFilter === "elite") awaiting = awaiting.filter((a) => a.waiting === "elite");
    else if (awaitingFilter === "vendor")
      awaiting = awaiting.filter((a) => a.waiting === "vendor");

    const waitingElite = awaiting.filter((a) => a.waiting === "elite").length;
    const waitingVendor = awaiting.filter((a) => a.waiting === "vendor").length;
    const vendorCount = new Set(awaiting.map((a) => a.vendor)).size;

    const metricsEl = document.getElementById("overview-metrics");
    metricsEl.innerHTML = [
      { label: "Waiting on Elite", value: waitingElite, accent: true },
      { label: "Waiting on Vendor", value: waitingVendor, accent: true },
      { label: "Orders open", value: awaiting.length, accent: false },
      { label: "Active vendors", value: vendorCount, accent: false },
    ]
      .map(
        (m) => `
      <article class="metric-card${m.accent ? " metric-card--accent" : ""}">
        <p class="metric-card__label">${m.label}</p>
        <p class="metric-card__value">${m.value}</p>
      </article>`
      )
      .join("");

    overviewAwaitingBase = awaiting;
    renderAwaitingTable(true);
  }

  const COMPLETED_COLS = [
    { label: "Order ID", sortKey: "id", render: (r) => r.id },
    { label: "Vendor name", sortKey: "vendor", render: (r) => r.vendor },
    { label: "Cloth type", sortKey: "cloth", render: (r) => r.cloth },
    { label: "Completion date", sortKey: "completed", render: (r) => r.completed },
    {
      label: "Time taken to complete (days)",
      sortKey: "days",
      render: (r) => r.days,
    },
  ];

  const RECEIVED_COLS = [
    { label: "Handler name", sortKey: "handler", render: (r) => r.handler },
    { label: "Vendor name", sortKey: "vendor", render: (r) => r.vendor },
    { label: "Order received date", sortKey: "received", render: (r) => r.received },
    { label: "Material type", sortKey: "material", render: (r) => r.material },
    {
      label: "Price quoted",
      sortKey: "quoted",
      render: (r) => (r.quoted == null ? "—" : `£${r.quoted.toFixed(2)}`),
    },
    { label: "Quantity", sortKey: "qty", render: (r) => `${r.qty.toLocaleString()} m` },
    { label: "Est. completion date", sortKey: "eta", render: (r) => r.eta },
  ];

  /* --- Completed --- */
  function fillSelect(id, values, allLabel) {
    const sel = document.getElementById(id);
    const opts = [`<option value="">${allLabel}</option>`].concat(
      values.map((v) => `<option value="${v}">${v}</option>`)
    );
    sel.innerHTML = opts.join("");
  }

  function refreshCompleted() {
    let rows = MOCK_COMPLETED.filter((r) => rowMatchesGlobalDate(r.completed));
    const v = document.getElementById("filter-completed-vendor").value;
    const m = document.getElementById("filter-completed-material").value;
    const o = document.getElementById("filter-completed-type").value;
    if (v) rows = rows.filter((r) => r.vendor === v);
    if (m) rows = rows.filter((r) => r.material === m);
    if (o) rows = rows.filter((r) => r.orderType === o);

    tableColumns.completed = COMPLETED_COLS;
    tableRowCache.completed = rows;
    document.getElementById("table-completed").dataset.sortKey = "";
    document.querySelectorAll("#table-completed th[data-sort]").forEach((h) => h.classList.remove("is-sorted"));
    const table = document.getElementById("table-completed");
    renderTableHead(table, COMPLETED_COLS);
    renderTableBody(table, rows, COMPLETED_COLS);
  }

  /* --- Ongoing --- */
  function refreshOngoing() {
    let filtered = MOCK_ONGOING.filter((r) => rowMatchesGlobalDate(r.received));
    const fv = document.getElementById("filter-ongoing-vendor")?.value || "";
    const fm = document.getElementById("filter-ongoing-material")?.value || "";
    const fp = document.getElementById("filter-ongoing-progress")?.value || "";
    const fi = document.getElementById("filter-ongoing-incharge")?.value || "";
    if (fv) filtered = filtered.filter((r) => r.vendor === fv);
    if (fm) filtered = filtered.filter((r) => r.material === fm);
    if (fp) filtered = filtered.filter((r) => r.progress === fp);
    if (fi) filtered = filtered.filter((r) => r.incharge === fi);
    const vendorStats = buildOngoingVendorStats(filtered);
    const rows = filtered.map((r) => {
      const s = vendorStats.get(r.vendor || "");
      const sortPct =
        s && s.count > 1 ? s.overallPct : ongoingProgressPercent(r);
      return { ...r, _ongoingSortProgress: sortPct };
    });
    if (ongoingExpandedId && !rows.some((r) => r.id === ongoingExpandedId)) {
      ongoingExpandedId = null;
      ongoingDrawerOrderId = null;
      ongoingEditOpenId = null;
    }
    tableColumns.ongoing = getOngoingCols(vendorStats);
    tableRowCache.ongoing = rows;
    document.getElementById("table-ongoing").dataset.sortKey = "";
    document.querySelectorAll("#table-ongoing th[data-sort]").forEach((h) => h.classList.remove("is-sorted"));
    const table = document.getElementById("table-ongoing");
    renderTableHead(table, tableColumns.ongoing);
    renderOngoingTableBody(table, rows);
  }

  /* --- Received --- */
  function refreshReceived() {
    const rows = MOCK_RECEIVED.filter((r) => rowMatchesGlobalDate(r.received));
    tableColumns.received = RECEIVED_COLS;
    tableRowCache.received = rows;
    document.getElementById("table-received").dataset.sortKey = "";
    document.querySelectorAll("#table-received th[data-sort]").forEach((h) => h.classList.remove("is-sorted"));
    const table = document.getElementById("table-received");
    renderTableHead(table, RECEIVED_COLS);
    renderTableBody(table, rows, RECEIVED_COLS);
  }

  /* --- Alerts --- */
  function refreshAlerts() {
    const el = document.getElementById("alerts-list");
    el.innerHTML = MOCK_ALERTS.map(
      (a) => `
      <article class="alert-item alert-item--${a.level}">
        <div>
          <h3 class="alert-item__title">${a.title}</h3>
          <p class="alert-item__meta">${a.meta}</p>
        </div>
      </article>`
    ).join("");
  }

  /* --- Pricing (stepped wireframe UI) --- */
  let lastQuote = null;
  let prsRevealedOverall = false;
  let prsRevealedWarp = false;
  let prsRevealedWeft = false;
  let prsRevealedFinal = false;
  /** Collapsed “compact + Edit” state for pricing steps (mirrors fabric quality general panel). */
  let prsAutoCollapsed = false;
  let prsOverallCollapsed = false;
  let prsWarpCollapsed = false;
  let prsWeftCollapsed = false;
  let prsFinalCollapsed = false;
  /** Session-only fabric specs created from Pricing → Create new (merged into dropdowns). */
  let prsRuntimeFabricQualityPresets = [];

  function prsNum(id) {
    const el = document.getElementById(id);
    if (!el) return NaN;
    const raw = el.value;
    if (raw === "" || (typeof raw === "string" && raw.trim() === "")) return NaN;
    const v = Number(raw);
    return Number.isFinite(v) ? v : NaN;
  }

  function prsFabricQualityPresets() {
    const base = typeof MOCK_FABRIC_QUALITY_PRESETS !== "undefined" ? MOCK_FABRIC_QUALITY_PRESETS : [];
    return [...base, ...prsRuntimeFabricQualityPresets];
  }

  function prsPopulateFabricQualitySelect() {
    const hidden = document.getElementById("prs-fabric-quality");
    const list = document.getElementById("prs-fabric-quality-listbox");
    if (!hidden || !list) return;
    const presets = prsFabricQualityPresets();
    const prev = hidden.value;
    const rows = [
      { value: "", label: "Select a saved specification…" },
      ...presets.map((p) => ({ value: p.id, label: p.label })),
    ];
    list.innerHTML = rows
      .map(
        (r) =>
          `<li role="presentation"><button type="button" role="option" class="fqs-custom-select__opt" data-value="${escapeAttr(r.value)}">${escapeHtml(r.label)}</button></li>`
      )
      .join("");
    hidden.value = prev && presets.some((x) => x.id === prev) ? prev : "";
    fqsRefreshCustomSelectDisplay(hidden.closest(".fqs-custom-select"));
  }

  let fqmSectionParent = null;
  let fqmSectionNext = null;

  function fabricQualityModalRestoreSectionToMain() {
    const modal = document.getElementById("fabric-quality-modal");
    const section = document.getElementById("view-fabric-quality");
    const modalBody = document.getElementById("fabric-quality-modal-body");
    if (!section || !modalBody || section.parentNode !== modalBody || !fqmSectionParent) return;
    modal?.classList.remove("is-open");
    if (fqmSectionNext && fqmSectionNext.parentNode === fqmSectionParent)
      fqmSectionParent.insertBefore(section, fqmSectionNext);
    else fqmSectionParent.appendChild(section);
    if (modal) {
      modal.hidden = true;
      modal.setAttribute("aria-hidden", "true");
    }
    document.body.style.overflow = "";
    const activeNav = document.querySelector(".nav-item.is-active");
    const key = activeNav?.dataset?.view || "";
    section.hidden = key !== "fabric-quality";
  }

  function closeFabricQualityModal() {
    const modal = document.getElementById("fabric-quality-modal");
    const section = document.getElementById("view-fabric-quality");
    const modalBody = document.getElementById("fabric-quality-modal-body");
    if (!modal || modal.hidden) return;

    const restore = () => {
      document.body.style.overflow = "";
      if (section && modalBody && section.parentNode === modalBody && fqmSectionParent) {
        if (fqmSectionNext && fqmSectionNext.parentNode === fqmSectionParent)
          fqmSectionParent.insertBefore(section, fqmSectionNext);
        else fqmSectionParent.appendChild(section);
        const activeNav = document.querySelector(".nav-item.is-active");
        const key = activeNav?.dataset?.view || "";
        section.hidden = key !== "fabric-quality";
      }
      modal.hidden = true;
      modal.setAttribute("aria-hidden", "true");
    };

    if (!modal.classList.contains("is-open")) {
      restore();
      return;
    }
    modal.classList.remove("is-open");
    let finished = false;
    const done = () => {
      if (finished) return;
      finished = true;
      restore();
    };
    modal.addEventListener(
      "transitionend",
      (e) => {
        if (e.target === modal && e.propertyName === "opacity") done();
      },
      { once: true }
    );
    window.setTimeout(done, 280);
  }

  function fqsResetFormToNew() {
    const cs = document.getElementById("fqs-customer-search");
    if (cs) cs.value = "";
    const em = document.getElementById("fqs-email");
    if (em) {
      em.value = "";
      const emWrap = em.closest(".fqs-custom-select");
      if (emWrap) fqsRefreshCustomSelectDisplay(emWrap);
    }
    const specH = document.getElementById("fqs-search-old");
    if (specH) {
      specH.value = "";
      const specWrap = specH.closest(".fqs-custom-select");
      if (specWrap) fqsRefreshCustomSelectDisplay(specWrap);
    }
    const mat = document.getElementById("fqs-material");
    if (mat) mat.value = "";
    const reed = document.getElementById("fqs-reed");
    if (reed) reed.value = "";
    const pick = document.getElementById("fqs-pick");
    if (pick) pick.value = "";
    const width = document.getElementById("fqs-width");
    if (width) width.value = "";
    const cert = document.getElementById("fqs-certificate");
    if (cert) cert.value = "";
    fqsSetAxisRows("warp", [{}]);
    fqsSetAxisRows("weft", [{}]);
    fqsRevealedWarp = false;
    fqsRevealedWeft = false;
    fqsRevealedReview = false;
    fqsGeneralCollapsed = false;
    fqsUpdateSummary();
    fqsSyncProgressiveSteps();
  }

  function fqsReadAxisRowsFromDom(axis) {
    const tbody = document.getElementById(`fqs-${axis}-tbody`);
    if (!tbody) return [];
    return [...tbody.querySelectorAll("tr")].map((tr) => {
      const endsRaw = tr.querySelector('[data-fqs-field="ends"]')?.value;
      let ends = "";
      if (endsRaw !== "" && endsRaw != null) {
        const n = Number(endsRaw);
        ends = Number.isFinite(n) ? n : "";
      }
      return {
        quality: tr.querySelector('[data-fqs-field="quality"]')?.value || "",
        count: tr.querySelector('[data-fqs-field="count"]')?.value || "",
        type: tr.querySelector('[data-fqs-field="type"]')?.value || "",
        texture: tr.querySelector('[data-fqs-field="texture"]')?.value || "",
        ends,
        extraEnds: Boolean(tr.querySelector('[data-fqs-field="extraEnds"]')?.checked),
      };
    });
  }

  function fqsBuildPresetLabelFromForm() {
    const mat = document.getElementById("fqs-material");
    const matLabel = mat?.selectedOptions[0]?.text?.trim() || "Fabric";
    const reed = document.getElementById("fqs-reed")?.value ?? "";
    const pick = document.getElementById("fqs-pick")?.value ?? "";
    const width = document.getElementById("fqs-width")?.value ?? "";
    return `${matLabel} — reed ${reed} / pick ${pick} / width ${width}`;
  }

  function fqsReadCurrentFormAsPreset(id, label) {
    const mat = document.getElementById("fqs-material");
    return {
      id,
      label,
      customerSearch: document.getElementById("fqs-customer-search")?.value?.trim() ?? "",
      email: document.getElementById("fqs-email")?.value?.trim() ?? "",
      material: mat?.value || "",
      reed: (() => {
        const v = document.getElementById("fqs-reed")?.value;
        const n = Number(v);
        return Number.isFinite(n) ? n : null;
      })(),
      pick: (() => {
        const v = document.getElementById("fqs-pick")?.value;
        const n = Number(v);
        return Number.isFinite(n) ? n : null;
      })(),
      width: (() => {
        const v = document.getElementById("fqs-width")?.value;
        const n = Number(v);
        return Number.isFinite(n) ? n : null;
      })(),
      certificate: document.getElementById("fqs-certificate")?.value || "",
      warp: fqsReadAxisRowsFromDom("warp"),
      weft: fqsReadAxisRowsFromDom("weft"),
    };
  }

  function fabricQualityModalApplyToPricing() {
    const warpTbody = document.getElementById("fqs-warp-tbody");
    const weftTbody = document.getElementById("fqs-weft-tbody");
    if (!fqsIsGeneralComplete()) {
      window.alert("Complete general fabric properties (material, reed, pick, width, certificate) before applying to pricing.");
      return;
    }
    if (!fqsIsAxisTbodyComplete(warpTbody) || !fqsIsAxisTbodyComplete(weftTbody)) {
      window.alert("Complete all warp and weft rows (quality, count, type, texture, and ends/picks) before applying to pricing.");
      return;
    }
    const id = `fq-new-${Date.now()}`;
    const label = fqsBuildPresetLabelFromForm();
    const preset = fqsReadCurrentFormAsPreset(id, label);
    prsRuntimeFabricQualityPresets.push(preset);
    fqsPopulateSavedSpecSelect();
    prsPopulateFabricQualitySelect();
    const hidden = document.getElementById("prs-fabric-quality");
    if (hidden) {
      hidden.value = id;
      fqsRefreshCustomSelectDisplay(hidden.closest(".fqs-custom-select"));
      hidden.dispatchEvent(new Event("change", { bubbles: true }));
    }
    closeFabricQualityModal();
    prsSyncProgressiveSteps();
  }

  function openFabricQualityModalFromPricing() {
    fabricQualityModalRestoreSectionToMain();
    const section = document.getElementById("view-fabric-quality");
    const modal = document.getElementById("fabric-quality-modal");
    const modalBody = document.getElementById("fabric-quality-modal-body");
    if (!section || !modal || !modalBody) return;
    if (!fqmSectionParent) {
      fqmSectionParent = section.parentNode;
      fqmSectionNext = section.nextSibling;
    }
    fqsResetFormToNew();
    modalBody.appendChild(section);
    section.hidden = false;
    document.body.style.overflow = "hidden";
    modal.hidden = false;
    modal.setAttribute("aria-hidden", "false");
    modal.classList.remove("is-open");
    requestAnimationFrame(() => {
      requestAnimationFrame(() => modal.classList.add("is-open"));
    });
    document.getElementById("fabric-quality-modal-apply")?.focus();
  }

  function initFabricQualityPricingModal() {
    const modal = document.getElementById("fabric-quality-modal");
    if (!modal || modal.dataset.fqmInit) return;
    modal.dataset.fqmInit = "1";
    modal.querySelector("[data-fabric-quality-modal-dismiss]")?.addEventListener("click", () => closeFabricQualityModal());
    document.getElementById("fabric-quality-modal-cancel")?.addEventListener("click", () => closeFabricQualityModal());
    document.getElementById("fabric-quality-modal-close-x")?.addEventListener("click", () => closeFabricQualityModal());
    document.getElementById("fabric-quality-modal-apply")?.addEventListener("click", () => fabricQualityModalApplyToPricing());
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape" || !modal || modal.hidden) return;
      if (document.querySelector(".fqs-custom-select[data-fqs-cs][data-open]")) return;
      closeFabricQualityModal();
    });
  }

  /**
   * Preset width: values <= 6 treated as metres (legacy); larger values as inches.
   * Returns inches for the fabric specifications field.
   */
  function prsPresetWidthToFabricSpecInches(width) {
    const w = Number(width);
    if (!Number.isFinite(w) || w <= 0) return NaN;
    if (w <= 6) return w / 0.0254;
    return w;
  }

  /** Build display line from a fabric-quality warp/weft row (matches Fabric quality axis table). */
  function prsFormatFqAxisRow(row) {
    if (!row || typeof row !== "object") return "";
    const parts = [];
    if (row.quality) parts.push(String(row.quality));
    if (row.count) parts.push(String(row.count));
    if (row.type) parts.push(String(row.type));
    if (row.texture) parts.push(String(row.texture));
    if (row.ends != null && row.ends !== "") parts.push(`${row.ends} picks`);
    if (row.extraEnds) parts.push("extra picks");
    return parts.join(" · ");
  }

  function prsSetPricingAxisHeading(headingId, prefix, index, row) {
    const el = document.getElementById(headingId);
    if (!el) return;
    const spec = prsFormatFqAxisRow(row);
    el.textContent = spec ? `${prefix} ${index} · ${spec}` : `${prefix} ${index} · —`;
  }

  function prsApplyPricingWarpWeftHeadingsFromPreset(p) {
    const warps = Array.isArray(p.warp) ? p.warp : [];
    const wefts = Array.isArray(p.weft) ? p.weft : [];
    prsSetPricingAxisHeading("prs-warp-block-1-heading", "Warp", 1, warps[0]);
    prsSetPricingAxisHeading("prs-warp-block-2-heading", "Warp", 2, warps[1]);
    prsSetPricingAxisHeading("prs-weft-block-1-heading", "Weft", 1, wefts[0]);
    prsSetPricingAxisHeading("prs-weft-block-2-heading", "Weft", 2, wefts[1]);
  }

  function prsResetPricingWarpWeftHeadings() {
    const set = (id, text) => {
      const el = document.getElementById(id);
      if (el) el.textContent = text;
    };
    set("prs-warp-block-1-heading", "Warp 1 · CPT 40 S");
    set("prs-warp-block-2-heading", "Warp 2 · CPT 2/40 S");
    set("prs-weft-block-1-heading", "Weft 1 · CPT 40 S");
    set("prs-weft-block-2-heading", "Weft 2 · CPT 2/40 S");
  }

  function prsApplyFabricQualityPreset(id) {
    const p = prsFabricQualityPresets().find((x) => x.id === id);
    if (!p) return;
    const setNum = (elId, v) => {
      const el = document.getElementById(elId);
      if (!el || v == null) return;
      const n = Number(v);
      if (!Number.isFinite(n)) return;
      el.value = String(n);
    };
    if (p.reed != null) {
      setNum("prs-reed", p.reed);
      setNum("prs-reed-loom", p.reed);
    }
    if (p.pick != null) {
      setNum("prs-pick", p.pick);
      setNum("prs-pick-table", p.pick);
    }
    const wIn = prsPresetWidthToFabricSpecInches(p.width);
    if (Number.isFinite(wIn)) setNum("prs-width", wIn);
    const cs = document.getElementById("prs-customer-search");
    const em = document.getElementById("prs-email");
    if (cs && p.customerSearch) cs.value = p.customerSearch;
    if (em && p.email) em.value = p.email;
    prsApplyPricingWarpWeftHeadingsFromPreset(p);
  }

  function prsSelectedFabricLabel(id) {
    const p = prsFabricQualityPresets().find((x) => x.id === id);
    return p ? p.label : "";
  }

  function readPricingWireframe() {
    const fabricQualityId = document.getElementById("prs-fabric-quality")?.value?.trim() ?? "";
    const warps = [
      {
        yarn: prsNum("prs-w1-yarn"),
        dye: prsNum("prs-w1-dye"),
        crimp: prsNum("prs-w1-crimp"),
        dyeWaste: prsNum("prs-w1-dye-waste"),
        warping: prsNum("prs-w1-warping"),
      },
      {
        yarn: prsNum("prs-w2-yarn"),
        dye: prsNum("prs-w2-dye"),
        crimp: prsNum("prs-w2-crimp"),
        dyeWaste: prsNum("prs-w2-dye-waste"),
        warping: prsNum("prs-w2-warping"),
      },
    ];
    const wefts = [
      { yarn: prsNum("prs-t1-yarn"), dye: prsNum("prs-t1-dye"), dyeWaste: prsNum("prs-t1-dye-waste"), selvedge: prsNum("prs-t1-selvedge") },
      { yarn: prsNum("prs-t2-yarn"), dye: prsNum("prs-t2-dye"), dyeWaste: prsNum("prs-t2-dye-waste"), selvedge: prsNum("prs-t2-selvedge") },
    ];
    return {
      customerSearch: document.getElementById("prs-customer-search")?.value.trim() ?? "",
      email: document.getElementById("prs-email")?.value.trim() ?? "",
      fabricQualityId,
      reed: prsNum("prs-reed"),
      pick: prsNum("prs-pick"),
      reedOnLoom: prsNum("prs-reed-loom"),
      pickOnTable: prsNum("prs-pick-table"),
      width: prsNum("prs-width"),
      overall: {
        pickRate: prsNum("prs-pick-rate"),
        finishing: prsNum("prs-finishing-cost"),
        kgPerM: prsNum("prs-kg-conv"),
        overheads: prsNum("prs-overheads"),
        processLoss: prsNum("prs-process-loss"),
      },
      warps,
      wefts,
    };
  }

  function prsFmtSpecCell(n) {
    return Number.isFinite(n) ? String(n) : "—";
  }

  function mapWireframeToLegacyQuoteInputs(w) {
    const wYarnAvg = (w.warps[0].yarn + w.warps[1].yarn) / 2;
    const fqLabel = prsSelectedFabricLabel(w.fabricQualityId) || w.fabricQualityId || "—";
    const material = `${w.customerSearch || "Customer"} · ${fqLabel}`;
    const sample = `Reed ${prsFmtSpecCell(w.reed)} / pick ${prsFmtSpecCell(w.pick)} · loom ${prsFmtSpecCell(w.reedOnLoom)} · table ${prsFmtSpecCell(w.pickOnTable)}`;
    const readpick = Number.isFinite(w.pick) && w.pick > 0 ? w.pick : (Number.isFinite(w.reed) && w.reed > 0 ? w.reed : 40);
    const widthInches = Number.isFinite(w.width) && w.width > 0 ? w.width : 60;
    const width = widthInches * 0.0254;
    const crimp = ((w.warps[0].crimp || 0) + (w.warps[1].crimp || 0)) / 2 || 2.5;
    const shrink = Number.isFinite(w.overall.processLoss) ? w.overall.processLoss : 3;
    return {
      material,
      readpick,
      rawprice: Number.isFinite(wYarnAvg) ? wYarnAvg : 12.5,
      width,
      sample,
      crimp,
      shrink,
      buffer: 8,
    };
  }

  function readPricingForm() {
    return mapWireframeToLegacyQuoteInputs(readPricingWireframe());
  }

  /** True when fabric specs, overall, warp, and weft grids are all filled (same bar as quotation inputs). */
  function prsIsQuotationCostInputsComplete() {
    return (
      prsIsAutoComplete() &&
      prsIsOverallComplete() &&
      prsIsWarpComplete() &&
      prsIsWeftComplete()
    );
  }

  /** Mid estimate (£/m) from the stepped pricing wireframe; NaN if any required step is incomplete. */
  function prsComputeGeneratedCostMid() {
    if (!prsIsQuotationCostInputsComplete()) return NaN;
    const { mid } = computeQuote(readPricingForm());
    return Number.isFinite(mid) && mid >= 0 ? mid : NaN;
  }

  /** Keep line 2 cost price in sync with the generated mid; refreshes final cost column. */
  function prsSyncQuotationCostFromCalculator() {
    const el = document.getElementById("qprice-cost-2");
    if (!el) return;
    const mid = prsComputeGeneratedCostMid();
    if (Number.isFinite(mid)) el.value = mid.toFixed(2);
    else el.value = "";
    syncQuotationPricingTable();
  }

  function prsFiniteNonNeg(n) {
    return Number.isFinite(n) && n >= 0;
  }

  function prsIsAutoComplete() {
    const w = readPricingWireframe();
    return (
      prsFiniteNonNeg(w.reed) &&
      prsFiniteNonNeg(w.pick) &&
      prsFiniteNonNeg(w.reedOnLoom) &&
      prsFiniteNonNeg(w.pickOnTable) &&
      Number.isFinite(w.width) &&
      w.width > 0
    );
  }

  function prsIsOverallComplete() {
    const o = readPricingWireframe().overall;
    return [o.pickRate, o.finishing, o.kgPerM, o.overheads, o.processLoss].every((x) => Number.isFinite(x) && x >= 0);
  }

  function prsIsWarpBlockComplete(x) {
    return [x.yarn, x.dye, x.crimp, x.dyeWaste, x.warping].every((n) => Number.isFinite(n) && n >= 0);
  }

  function prsIsWeftBlockComplete(x) {
    return [x.yarn, x.dye, x.dyeWaste, x.selvedge].every((n) => Number.isFinite(n) && n >= 0);
  }

  function prsIsWarpComplete() {
    const w = readPricingWireframe();
    return w.warps.every(prsIsWarpBlockComplete);
  }

  function prsIsWeftComplete() {
    const w = readPricingWireframe();
    return w.wefts.every(prsIsWeftBlockComplete);
  }

  function prsSetPanelLocked(id, locked) {
    const panel = document.getElementById(id);
    if (!panel) return;
    panel.classList.toggle("is-locked", locked);
    if (locked) panel.setAttribute("aria-disabled", "true");
    else panel.removeAttribute("aria-disabled");
  }

  function prsUpdateAutoCompactSummary() {
    const el = document.getElementById("prs-auto-compact-text");
    if (!el) return;
    const reed = document.getElementById("prs-reed")?.value ?? "—";
    const pick = document.getElementById("prs-pick")?.value ?? "—";
    const loom = document.getElementById("prs-reed-loom")?.value ?? "—";
    const table = document.getElementById("prs-pick-table")?.value ?? "—";
    const w = document.getElementById("prs-width")?.value ?? "—";
    el.textContent = `Reed ${reed} · Pick ${pick} · Loom ${loom} · Table ${table} · Width ${w} in`;
  }

  function prsUpdateOverallCompactSummary() {
    const el = document.getElementById("prs-overall-compact-text");
    if (!el) return;
    const o = readPricingWireframe().overall;
    el.textContent = `Pick rate ${prsFmtSpecCell(o.pickRate)} · Finishing ${prsFmtSpecCell(o.finishing)} £/m · Kg/m ${prsFmtSpecCell(o.kgPerM)} · Overheads ${prsFmtSpecCell(o.overheads)} · Loss ${prsFmtSpecCell(o.processLoss)}%`;
  }

  function prsUpdateWarpCompactSummary() {
    const el = document.getElementById("prs-warp-compact-text");
    if (!el) return;
    const h1 = document.getElementById("prs-warp-block-1-heading")?.textContent?.trim() || "Warp 1";
    const h2 = document.getElementById("prs-warp-block-2-heading")?.textContent?.trim() || "Warp 2";
    const y1 = document.getElementById("prs-w1-yarn")?.value ?? "—";
    const y2 = document.getElementById("prs-w2-yarn")?.value ?? "—";
    el.textContent = `${h1}: yarn £${y1} · ${h2}: yarn £${y2}`;
  }

  function prsUpdateWeftCompactSummary() {
    const el = document.getElementById("prs-weft-compact-text");
    if (!el) return;
    const h1 = document.getElementById("prs-weft-block-1-heading")?.textContent?.trim() || "Weft 1";
    const h2 = document.getElementById("prs-weft-block-2-heading")?.textContent?.trim() || "Weft 2";
    const y1 = document.getElementById("prs-t1-yarn")?.value ?? "—";
    const y2 = document.getElementById("prs-t2-yarn")?.value ?? "—";
    el.textContent = `${h1}: yarn £${y1} · ${h2}: yarn £${y2}`;
  }

  function prsUpdateFinalCompactSummary() {
    const el = document.getElementById("prs-final-compact-text");
    if (!el) return;
    const range = document.getElementById("quote-range")?.textContent?.trim() || "";
    const fin = document.getElementById("qprice-final-2")?.value?.trim() || "";
    const ops = document.getElementById("qprice-ops-2")?.value ?? "";
    const pct = document.getElementById("qprice-profit-2")?.value ?? "";
    const parts = [];
    if (range) parts.push(range);
    if (fin) parts.push(`Final ${fin}`);
    if (ops !== "" || pct !== "") parts.push(`Ops ${ops || "—"} · Profit ${pct || "—"}%`);
    el.textContent = parts.length ? parts.join(" · ") : "Quotation — complete costing to calculate.";
  }

  function prsSyncCollapsePanelsUi() {
    const autoOk = prsIsAutoComplete();
    const overallOk = autoOk && prsIsOverallComplete();
    const warpOk = overallOk && prsIsWarpComplete();
    const weftOk = warpOk && prsIsWeftComplete();

    if (!autoOk) prsAutoCollapsed = false;
    if (!overallOk) prsOverallCollapsed = false;
    if (!warpOk) prsWarpCollapsed = false;
    if (!weftOk) prsWeftCollapsed = false;
    if (!prsRevealedFinal) prsFinalCollapsed = false;

    const toggle = (id, collapsed) => {
      const panel = document.getElementById(id);
      if (panel) panel.classList.toggle("is-collapsed", Boolean(collapsed));
    };

    toggle("prs-panel-auto", prsAutoCollapsed && autoOk);
    toggle("prs-panel-overall", prsOverallCollapsed && overallOk);
    toggle("prs-panel-warp", prsWarpCollapsed && warpOk);
    toggle("prs-panel-weft", prsWeftCollapsed && weftOk);
    toggle(
      "prs-panel-final",
      prsFinalCollapsed && prsRevealedFinal && prsIsQuotationCostInputsComplete()
    );

    if (prsAutoCollapsed && autoOk) prsUpdateAutoCompactSummary();
    if (prsOverallCollapsed && overallOk) prsUpdateOverallCompactSummary();
    if (prsWarpCollapsed && warpOk) prsUpdateWarpCompactSummary();
    if (prsWeftCollapsed && weftOk) prsUpdateWeftCompactSummary();
    if (prsFinalCollapsed && prsRevealedFinal) prsUpdateFinalCompactSummary();
  }

  function prsSyncProgressiveSteps() {
    const autoFieldsOk = prsIsAutoComplete();
    const overallOk = autoFieldsOk && prsIsOverallComplete();
    const warpOk = overallOk && prsIsWarpComplete();
    const weftOk = warpOk && prsIsWeftComplete();

    if (!autoFieldsOk) {
      prsRevealedOverall = false;
      prsRevealedWarp = false;
      prsRevealedWeft = false;
      prsRevealedFinal = false;
    } else if (!overallOk) {
      prsRevealedWarp = false;
      prsRevealedWeft = false;
      prsRevealedFinal = false;
    } else if (!warpOk) {
      prsRevealedWeft = false;
      prsRevealedFinal = false;
    } else if (!weftOk) {
      prsRevealedFinal = false;
    }

    const overallUnlocked = autoFieldsOk && prsRevealedOverall;
    const warpUnlocked = overallUnlocked && overallOk && prsRevealedWarp;
    const weftUnlocked = warpUnlocked && warpOk && prsRevealedWeft;
    const finalUnlocked = weftUnlocked && weftOk && prsRevealedFinal;

    prsSetPanelLocked("prs-panel-auto", false);
    prsSetPanelLocked("prs-panel-overall", !overallUnlocked);
    prsSetPanelLocked("prs-panel-warp", !warpUnlocked);
    prsSetPanelLocked("prs-panel-weft", !weftUnlocked);
    prsSetPanelLocked("prs-panel-final", !finalUnlocked);

    prsSyncCollapsePanelsUi();

    const n2 = document.getElementById("prs-next-auto");
    if (n2) n2.disabled = !prsIsAutoComplete() || prsAutoCollapsed;

    const n3 = document.getElementById("prs-next-overall");
    if (n3) n3.disabled = !(overallUnlocked && overallOk) || prsOverallCollapsed;

    const n4 = document.getElementById("prs-next-warp");
    if (n4) n4.disabled = !(warpUnlocked && warpOk) || prsWarpCollapsed;

    const n5 = document.getElementById("prs-next-weft");
    if (n5) n5.disabled = !(weftUnlocked && weftOk) || prsWeftCollapsed;

    prsSyncQuotationCostFromCalculator();
    if (prsRevealedFinal && prsIsQuotationCostInputsComplete()) runPricing();
  }

  function initPricingPage() {
    const root = document.getElementById("view-pricing");
    if (!root || root.dataset.prsInit) return;
    root.dataset.prsInit = "1";
    prsRevealedOverall = false;
    prsRevealedWarp = false;
    prsRevealedWeft = false;
    prsRevealedFinal = false;
    prsAutoCollapsed = false;
    prsOverallCollapsed = false;
    prsWarpCollapsed = false;
    prsWeftCollapsed = false;
    prsFinalCollapsed = false;
    prsPopulateFabricQualitySelect();
    prsSyncProgressiveSteps();
    root.addEventListener("input", () => {
      prsSyncProgressiveSteps();
    });
    root.addEventListener("change", (e) => {
      if (e.target.matches("input, select, textarea")) prsSyncProgressiveSteps();
    });

    document.getElementById("prs-fabric-quality")?.addEventListener("change", (e) => {
      const id = e.target.value?.trim() ?? "";
      if (id) prsApplyFabricQualityPreset(id);
      else prsResetPricingWarpWeftHeadings();
      prsSyncProgressiveSteps();
    });

    document.getElementById("prs-fabric-create-new")?.addEventListener("click", () => {
      openFabricQualityModalFromPricing();
    });

    initFabricQualityPricingModal();

    document.getElementById("prs-next-auto")?.addEventListener("click", () => {
      if (!prsIsAutoComplete()) return;
      prsAutoCollapsed = true;
      prsRevealedOverall = true;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-overall")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("prs-next-overall")?.addEventListener("click", () => {
      if (!prsIsOverallComplete()) return;
      prsOverallCollapsed = true;
      prsRevealedWarp = true;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-warp")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("prs-next-warp")?.addEventListener("click", () => {
      if (!prsIsWarpComplete()) return;
      prsWarpCollapsed = true;
      prsRevealedWeft = true;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-weft")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("prs-next-weft")?.addEventListener("click", () => {
      if (!prsIsWeftComplete()) return;
      prsWeftCollapsed = true;
      prsRevealedFinal = true;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-final")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("prs-auto-edit")?.addEventListener("click", () => {
      prsAutoCollapsed = false;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-auto")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("prs-overall-edit")?.addEventListener("click", () => {
      prsOverallCollapsed = false;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-overall")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("prs-warp-edit")?.addEventListener("click", () => {
      prsWarpCollapsed = false;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-warp")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("prs-weft-edit")?.addEventListener("click", () => {
      prsWeftCollapsed = false;
      prsSyncProgressiveSteps();
      document.getElementById("prs-panel-weft")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });
    document.getElementById("prs-final-edit")?.addEventListener("click", () => {
      prsFinalCollapsed = false;
      prsSyncCollapsePanelsUi();
      document.getElementById("prs-panel-final")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("prs-quotation-save")?.addEventListener("click", () => {
      syncQuotationPricingTable();
      const line2 = {
        cost: document.getElementById("qprice-cost-2")?.value ?? "",
        opsBuffer: document.getElementById("qprice-ops-2")?.value ?? "",
        profitPct: document.getElementById("qprice-profit-2")?.value ?? "",
        final: document.getElementById("quote-final")?.value ?? "",
      };
      try {
        sessionStorage.setItem("elite_quotation_pricing", JSON.stringify({ savedAt: Date.now(), line2 }));
      } catch (_) {
        /* ignore quota / private mode */
      }
      window.alert("Quotation saved for this browser session (demo).");
      prsFinalCollapsed = true;
      prsSyncCollapsePanelsUi();
    });

    const prsFinalPanel = document.getElementById("prs-panel-final");
    prsFinalPanel?.addEventListener("input", (e) => {
      if (!e.target.closest("#quotation-pricing-table")) return;
      syncQuotationPricingTable();
      if (prsRevealedFinal && prsIsQuotationCostInputsComplete()) runPricing();
    });
    prsFinalPanel?.addEventListener("change", (e) => {
      if (!e.target.closest("#quotation-pricing-table")) return;
      syncQuotationPricingTable();
      if (prsRevealedFinal && prsIsQuotationCostInputsComplete()) runPricing();
    });
  }

  function readMailPricingForm() {
    return {
      material: document.getElementById("mpf-material").value.trim(),
      readpick: Number(document.getElementById("mpf-readpick").value) || 0,
      rawprice: Number(document.getElementById("mpf-rawprice").value) || 0,
      width: Number(document.getElementById("mpf-width").value) || 0,
      sample: document.getElementById("mpf-sample").value.trim(),
      crimp: Number(document.getElementById("mpf-crimp").value) || 0,
      shrink: Number(document.getElementById("mpf-shrink").value) || 0,
      buffer: Number(document.getElementById("mpf-buffer").value) || 0,
    };
  }

  function syncPricingPageToMailForm() {
    const legacy = readPricingForm();
    const set = (id, v) => {
      const el = document.getElementById(id);
      if (el) el.value = v;
    };
    set("mpf-material", legacy.material);
    set("mpf-readpick", String(legacy.readpick));
    set("mpf-rawprice", String(legacy.rawprice));
    set("mpf-width", String(legacy.width));
    set("mpf-sample", legacy.sample);
    set("mpf-crimp", String(legacy.crimp));
    set("mpf-shrink", String(legacy.shrink));
    set("mpf-buffer", String(legacy.buffer));
  }

  function syncMailFormToPricingPage() {
    const mMat = document.getElementById("mpf-material")?.value.trim() ?? "";
    const mRead = Number(document.getElementById("mpf-readpick")?.value);
    const mRaw = Number(document.getElementById("mpf-rawprice")?.value);
    const mW = Number(document.getElementById("mpf-width")?.value);
    const mSample = document.getElementById("mpf-sample")?.value ?? "";
    const mCr = Number(document.getElementById("mpf-crimp")?.value);
    const mSh = Number(document.getElementById("mpf-shrink")?.value);
    const prsFqSel = document.getElementById("prs-fabric-quality");
    if (document.getElementById("prs-customer-search") && mMat) {
      const parts = mMat.split("·").map((s) => s.trim());
      if (parts.length >= 2) {
        document.getElementById("prs-customer-search").value = parts[0];
        const rest = parts.slice(1).join(" · ").trim();
        const presets = prsFabricQualityPresets();
        const preset =
          presets.find((p) => p.label === rest) ||
          presets.find((p) => rest.includes(p.label) || p.label.includes(rest));
        if (prsFqSel) {
          prsFqSel.value = preset ? preset.id : "";
          fqsRefreshCustomSelectDisplay(prsFqSel.closest(".fqs-custom-select"));
        }
      }
    }
    if (Number.isFinite(mRead) && document.getElementById("prs-pick")) document.getElementById("prs-pick").value = String(mRead);
    if (Number.isFinite(mRaw) && document.getElementById("prs-w1-yarn")) document.getElementById("prs-w1-yarn").value = String(mRaw);
    if (Number.isFinite(mW) && mW > 0 && document.getElementById("prs-width"))
      document.getElementById("prs-width").value = String(mW / 0.0254);
    if (Number.isFinite(mCr) && document.getElementById("prs-w1-crimp")) document.getElementById("prs-w1-crimp").value = String(mCr);
    if (Number.isFinite(mSh) && document.getElementById("prs-process-loss")) document.getElementById("prs-process-loss").value = String(mSh);
    if (mSample && document.getElementById("prs-reed")) {
      const reedM = mSample.match(/Reed\s+([\d.]+)/i);
      if (reedM) document.getElementById("prs-reed").value = reedM[1];
    }
    prsSyncProgressiveSteps();
  }

  function computeQuote(inputs) {
    const { readpick, rawprice, width, sample, crimp, shrink, buffer } = inputs;
    const yarnFactor = 1 + (readpick - 30) * 0.012;
    const widthFactor = Math.max(width, 0.1);
    const patternFactor = 1 + Math.min(sample.length * 0.002, 0.28);
    let base =
      rawprice * widthFactor * yarnFactor * patternFactor;
    base *= 1 + crimp / 100;
    base *= 1 + shrink / 100;
    const beforeBuffer = base;
    base *= 1 + buffer / 100;
    const mid = base;
    const low = mid * 0.93;
    const high = mid * 1.09;
    return { low, mid, high, beforeBuffer };
  }

  function formatMoney(n) {
    return `£${n.toFixed(2)}`;
  }

  function qpriceNum(elId) {
    const el = document.getElementById(elId);
    if (!el) return NaN;
    const raw = el.value;
    if (raw === "" || (typeof raw === "string" && raw.trim() === "")) return NaN;
    const v = Number(raw);
    return Number.isFinite(v) ? v : NaN;
  }

  /** Final cost = cost + ops buffer + profit% of cost. */
  function syncQuotationPricingTable() {
    const c2 = qpriceNum("qprice-cost-2");
    const o2 = qpriceNum("qprice-ops-2");
    const p2 = qpriceNum("qprice-profit-2");
    const out2 = document.getElementById("qprice-final-2");
    const hiddenFinal = document.getElementById("quote-final");
    if (out2) {
      if (!Number.isFinite(c2) || c2 < 0 || !Number.isFinite(o2) || o2 < 0 || !Number.isFinite(p2) || p2 < 0) {
        out2.value = "";
        if (hiddenFinal) hiddenFinal.value = "";
      } else {
        const f = c2 + o2 + (c2 * p2) / 100;
        out2.value = formatMoney(f);
        if (hiddenFinal) hiddenFinal.value = f.toFixed(2);
      }
    }
  }

  function runPricing(e) {
    if (e) e.preventDefault();
    if (!prsRevealedFinal) {
      document.getElementById("quote-range").textContent =
        "Complete each step and use Next until this section opens; the quotation row will update automatically.";
      document.getElementById("quote-final").value = "";
      const o2 = document.getElementById("qprice-final-2");
      if (o2) o2.value = "";
      document.getElementById("quote-breakdown").textContent = "";
      lastQuote = null;
      syncMailPricingMini();
      return;
    }
    const inputs = readPricingForm();
    lastQuote = { inputs, ...computeQuote(inputs) };
    const { low, mid, high, beforeBuffer } = lastQuote;
    document.getElementById("quote-range").textContent = `Suggested range: ${formatMoney(low)} — ${formatMoney(high)} per metre (mid ${formatMoney(mid)})`;
    prsSyncQuotationCostFromCalculator();
    document.getElementById("quote-breakdown").textContent =
      `After crimp & shrink, before buffer: ${formatMoney(beforeBuffer)} / m. Cost price follows the generated mid from your steps; adjust Ops Buffer and Profit % for the customer quotation.`;
    syncMailPricingMini();
    syncPricingPageToMailForm();
    const mqf = document.getElementById("mail-quote-final");
    const qf = document.getElementById("quote-final")?.value;
    if (mqf && qf) mqf.value = qf;
    const mqr = document.getElementById("mail-quote-range");
    if (mqr)
      mqr.textContent = `Range: ${formatMoney(low)} — ${formatMoney(high)} / m`;
    if (prsFinalCollapsed) prsSyncCollapsePanelsUi();
  }

  function runMailPricing() {
    const inputs = readMailPricingForm();
    lastQuote = { inputs, ...computeQuote(inputs) };
    const { low, mid, high } = lastQuote;
    const mqr = document.getElementById("mail-quote-range");
    if (mqr)
      mqr.textContent = `Suggested range: ${formatMoney(low)} — ${formatMoney(high)} per metre (mid ${formatMoney(mid)})`;
    syncMailFormToPricingPage();
    const c2 = document.getElementById("qprice-cost-2");
    const o2 = document.getElementById("qprice-ops-2");
    const p2 = document.getElementById("qprice-profit-2");
    if (c2) c2.value = mid.toFixed(2);
    if (o2) o2.value = "0";
    if (p2) p2.value = "0";
    syncQuotationPricingTable();
    const qf = document.getElementById("quote-final")?.value;
    const mqf = document.getElementById("mail-quote-final");
    if (mqf && qf) mqf.value = qf;
    document.getElementById("quote-range").textContent = `Suggested range: ${formatMoney(low)} — ${formatMoney(high)} per metre (mid ${formatMoney(mid)})`;
    syncMailPricingMini();
  }

  function syncMailPricingMini() {
    const mini = document.getElementById("mail-pricing-mini");
    if (!mini) return;
    if (!lastQuote) {
      mini.textContent = "Use Pricing (stepped flow) to complete costing, or Calculate in Compose below.";
      return;
    }
    const m = lastQuote.inputs.material;
    mini.innerHTML = `Cloth type: ${escapeHtml(m)} · Range ${formatMoney(lastQuote.low)} — ${formatMoney(lastQuote.high)} / m`;
  }

  document.getElementById("btn-print-quote").addEventListener("click", () => {
    if (!lastQuote) runPricing();
    if (!lastQuote) return;
    const final = Number(document.getElementById("quote-final").value) || lastQuote.mid;
    const { inputs } = lastQuote;
    const existing = document.getElementById("print-root");
    if (existing) existing.remove();
    const root = document.createElement("div");
    root.id = "print-root";
    root.style.cssText = "background:#fff;color:#111;padding:2rem;";
    root.innerHTML = `
      <h1 style="font-family:system-ui;margin:0 0 0.5rem;">Elite Textile — Quotation</h1>
      <p style="margin:0 0 1rem;color:#333;">${new Date().toLocaleString()}</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px;">
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Material</strong></td><td style="padding:6px;border:1px solid #ccc;">${escapeHtml(inputs.material)}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Width</strong></td><td style="padding:6px;border:1px solid #ccc;">${inputs.width} m</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Read / pick</strong></td><td style="padding:6px;border:1px solid #ccc;">${inputs.readpick}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Sample / pattern</strong></td><td style="padding:6px;border:1px solid #ccc;">${escapeHtml(inputs.sample)}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Quoted price / m</strong></td><td style="padding:6px;border:1px solid #ccc;font-size:18px;"><strong>${formatMoney(final)}</strong></td></tr>
      </table>
      <p style="margin-top:1rem;font-size:12px;color:#666;">System range was ${formatMoney(lastQuote.low)} — ${formatMoney(lastQuote.high)} / m before manual edit.</p>
    `;
    document.body.appendChild(root);
    window.print();
    root.remove();
  });

  function escapeHtml(s) {
    const d = document.createElement("div");
    d.textContent = s;
    return d.innerHTML;
  }

  function fqsOpt(value, label, selected) {
    return `<option value="${escapeAttr(value)}"${selected === value ? " selected" : ""}>${escapeHtml(label)}</option>`;
  }

  function fqsAxisRowHtml(axis, index, d) {
    const q = d.quality || "";
    const c = d.count || "";
    const t = d.type || "";
    const tex = d.texture || "";
    const ends = d.ends !== undefined && d.ends !== "" ? d.ends : "";
    const ex = Boolean(d.extraEnds);
    const capAxis = axis.charAt(0).toUpperCase() + axis.slice(1);
    const endsVal = ends === "" ? "" : escapeAttr(String(ends));
    const countLabel = axis === "weft" ? "picks" : "ends";
    const extraCountLabel = axis === "weft" ? "extra picks" : "extra ends";
    return `<tr data-fqs-axis="${escapeAttr(axis)}" data-fqs-index="${index}">
      <td><span class="fqs-axis-pill">${escapeHtml(capAxis)} ${index}</span></td>
      <td><select class="select select--table" data-fqs-field="quality" aria-label="${escapeAttr(capAxis)} ${index} quality">
        <option value="">—</option>
        ${fqsOpt("premium-a", "Premium A", q)}
        ${fqsOpt("standard-b", "Standard B", q)}
        ${fqsOpt("mill-c", "Mill run C", q)}
      </select></td>
      <td><select class="select select--table" data-fqs-field="count" aria-label="${escapeAttr(capAxis)} ${index} count">
        <option value="">—</option>
        ${fqsOpt("20s", "20s", c)}
        ${fqsOpt("30s", "30s", c)}
        ${fqsOpt("40s", "40s", c)}
        ${fqsOpt("60s", "60s", c)}
      </select></td>
      <td><select class="select select--table" data-fqs-field="type" aria-label="${escapeAttr(capAxis)} ${index} type">
        <option value="">—</option>
        ${fqsOpt("carded", "Carded", t)}
        ${fqsOpt("combed", "Combed", t)}
        ${fqsOpt("ring-spun", "Ring spun", t)}
      </select></td>
      <td><select class="select select--table" data-fqs-field="texture" aria-label="${escapeAttr(capAxis)} ${index} texture">
        <option value="">—</option>
        ${fqsOpt("plain", "Plain", tex)}
        ${fqsOpt("slub", "Slub", tex)}
        ${fqsOpt("nep", "Nep", tex)}
      </select></td>
      <td><input type="number" class="input" data-fqs-field="ends" min="0" step="1" placeholder="—" value="${endsVal}" aria-label="${escapeAttr(capAxis)} ${index} ${escapeAttr(countLabel)}" /></td>
      <td class="table--fqs-axis__check"><input type="checkbox" class="fqs-check" data-fqs-field="extraEnds"${ex ? " checked" : ""} aria-label="${escapeAttr(capAxis)} ${index} ${escapeAttr(extraCountLabel)}" /></td>
      <td><button type="button" class="btn btn--ghost btn--small" data-fqs-remove-row aria-label="Remove row">Remove</button></td>
    </tr>`;
  }

  function fqsRefreshAddButtons(axis) {
    const tbody = document.getElementById(`fqs-${axis}-tbody`);
    const btn = document.getElementById(`fqs-${axis}-add`);
    if (!tbody || !btn) return;
    const n = tbody.querySelectorAll("tr").length;
    btn.disabled = n >= 3;
    tbody.querySelectorAll("[data-fqs-remove-row]").forEach((b) => {
      b.disabled = n <= 1;
    });
  }

  function fqsRenumberAxis(axis) {
    const tbody = document.getElementById(`fqs-${axis}-tbody`);
    if (!tbody) return;
    const capAxis = axis.charAt(0).toUpperCase() + axis.slice(1);
    [...tbody.querySelectorAll("tr")].forEach((tr, i) => {
      tr.dataset.fqsIndex = String(i + 1);
      const pill = tr.querySelector(".fqs-axis-pill");
      if (pill) pill.textContent = `${capAxis} ${i + 1}`;
    });
  }

  function fqsSetAxisRows(axis, rows) {
    const tbody = document.getElementById(`fqs-${axis}-tbody`);
    if (!tbody) return;
    tbody.innerHTML = "";
    const list = rows && rows.length ? rows : [{}];
    list.forEach((rd, i) => {
      tbody.insertAdjacentHTML("beforeend", fqsAxisRowHtml(axis, i + 1, rd));
    });
    fqsRefreshAddButtons(axis);
  }

  function fqsApplyPreset(p) {
    if (!p) return;
    const cs = document.getElementById("fqs-customer-search");
    const em = document.getElementById("fqs-email");
    if (cs) cs.value = p.customerSearch || "";
    if (em) {
      const addr = p.email || "";
      em.value = addr;
      if (addr) {
        const list = document.getElementById("fqs-email-listbox");
        const exists = list?.querySelector(`button[data-value="${CSS.escape(addr)}"]`);
        if (list && !exists) {
          list.insertAdjacentHTML(
            "beforeend",
            `<li role="presentation"><button type="button" role="option" class="fqs-custom-select__opt" data-value="${escapeAttr(addr)}">${escapeHtml(addr)}</button></li>`
          );
        }
      }
      const emWrap = em.closest(".fqs-custom-select");
      if (emWrap) fqsRefreshCustomSelectDisplay(emWrap);
    }
    const mat = document.getElementById("fqs-material");
    if (mat) mat.value = p.material || "";
    const reed = document.getElementById("fqs-reed");
    if (reed) reed.value = p.reed != null ? p.reed : "";
    const pick = document.getElementById("fqs-pick");
    if (pick) pick.value = p.pick != null ? p.pick : "";
    const width = document.getElementById("fqs-width");
    if (width) width.value = p.width != null ? p.width : "";
    const cert = document.getElementById("fqs-certificate");
    if (cert) cert.value = p.certificate || "";
    fqsSetAxisRows("warp", p.warp);
    fqsSetAxisRows("weft", p.weft);
    const specH = document.getElementById("fqs-search-old");
    if (specH && p.id) {
      specH.value = p.id;
      const specWrap = specH.closest(".fqs-custom-select");
      if (specWrap) fqsRefreshCustomSelectDisplay(specWrap);
    }
    fqsRevealedWarp = true;
    fqsRevealedWeft = true;
    fqsRevealedReview = true;
    fqsGeneralCollapsed = true;
    fqsUpdateSummary();
    fqsSyncProgressiveSteps();
  }

  function fqsUpdateSummary() {
    const out = document.getElementById("fqs-fabric-quality-summary");
    if (!out) return;
    const mat = document.getElementById("fqs-material");
    const matLabel = mat?.selectedOptions[0]?.text?.trim() || "—";
    const reed = document.getElementById("fqs-reed")?.value || "—";
    const pick = document.getElementById("fqs-pick")?.value || "—";
    const width = document.getElementById("fqs-width")?.value || "—";
    const certEl = document.getElementById("fqs-certificate");
    const certLabel = certEl?.selectedOptions[0]?.text?.trim() || "—";
    const wn = document.getElementById("fqs-warp-tbody")?.querySelectorAll("tr").length || 0;
    const tn = document.getElementById("fqs-weft-tbody")?.querySelectorAll("tr").length || 0;
    out.textContent = `Material: ${matLabel} · Reed ${reed} · Pick ${pick} · Width ${width} · Certificate: ${certLabel} · Warp sets: ${wn} · Weft sets: ${tn}`;
  }

  function fqsOnAxisClick(e) {
    const rm = e.target.closest("[data-fqs-remove-row]");
    if (!rm || rm.disabled) return;
    const tr = rm.closest("tr");
    const tbody = tr && tr.closest("tbody");
    if (!tr || !tbody) return;
    if (tbody.querySelectorAll("tr").length <= 1) return;
    tr.remove();
    const axis = tbody.id.replace("fqs-", "").replace("-tbody", "");
    fqsRenumberAxis(axis);
    fqsRefreshAddButtons(axis);
    fqsUpdateSummary();
    fqsSyncProgressiveSteps();
  }

  function fqsIsGeneralComplete() {
    const mat = document.getElementById("fqs-material")?.value?.trim();
    const cert = document.getElementById("fqs-certificate")?.value?.trim();
    const reed = document.getElementById("fqs-reed")?.value;
    const pick = document.getElementById("fqs-pick")?.value;
    const width = document.getElementById("fqs-width")?.value;
    if (!mat || !cert) return false;
    if (reed === "" || reed == null || pick === "" || pick == null || width === "" || width == null) return false;
    const rn = Number(reed);
    const pn = Number(pick);
    const wn = Number(width);
    if (!Number.isFinite(rn) || !Number.isFinite(pn) || !Number.isFinite(wn)) return false;
    if (rn < 0 || pn < 0 || wn < 0) return false;
    return true;
  }

  function fqsIsAxisTbodyComplete(tbody) {
    if (!tbody) return false;
    const rows = tbody.querySelectorAll("tr");
    if (!rows.length) return false;
    for (const tr of rows) {
      const q = tr.querySelector('[data-fqs-field="quality"]')?.value;
      const c = tr.querySelector('[data-fqs-field="count"]')?.value;
      const t = tr.querySelector('[data-fqs-field="type"]')?.value;
      const tex = tr.querySelector('[data-fqs-field="texture"]')?.value;
      const ends = tr.querySelector('[data-fqs-field="ends"]')?.value;
      if (!q || !c || !t || !tex || ends === "" || ends == null) return false;
      const en = Number(ends);
      if (!Number.isFinite(en) || en < 0) return false;
    }
    return true;
  }

  function fqsSetProgressPanelLocked(panel, locked) {
    if (!panel) return;
    panel.classList.toggle("is-locked", locked);
    if (locked) panel.setAttribute("aria-disabled", "true");
    else panel.removeAttribute("aria-disabled");
  }

  function fqsUpdateGeneralCompactSummary() {
    const el = document.getElementById("fqs-general-compact-text");
    if (!el) return;
    const mat = document.getElementById("fqs-material");
    const matLabel = mat?.selectedOptions[0]?.text?.trim() || "—";
    const reed = document.getElementById("fqs-reed")?.value ?? "—";
    const pick = document.getElementById("fqs-pick")?.value ?? "—";
    const width = document.getElementById("fqs-width")?.value ?? "—";
    const certEl = document.getElementById("fqs-certificate");
    const certLabel = certEl?.selectedOptions[0]?.text?.trim() || "—";
    el.textContent = `${matLabel} · Reed ${reed} · Pick ${pick} · Width ${width} · ${certLabel}`;
  }

  function fqsSyncGeneralPanelCollapsedUi() {
    const panel = document.getElementById("fqs-panel-general");
    const genOk = fqsIsGeneralComplete();
    if (!genOk) fqsGeneralCollapsed = false;
    if (!panel) return;
    const collapsed = Boolean(fqsGeneralCollapsed && genOk);
    panel.classList.toggle("is-collapsed", collapsed);
    if (collapsed) fqsUpdateGeneralCompactSummary();
  }

  function fqsSyncProgressiveSteps() {
    const genOk = fqsIsGeneralComplete();
    const warpTbody = document.getElementById("fqs-warp-tbody");
    const weftTbody = document.getElementById("fqs-weft-tbody");
    const warpRowsOk = genOk && fqsIsAxisTbodyComplete(warpTbody);
    const weftRowsOk = warpRowsOk && fqsIsAxisTbodyComplete(weftTbody);

    if (!genOk) {
      fqsRevealedWarp = false;
      fqsRevealedWeft = false;
      fqsRevealedReview = false;
    } else if (!warpRowsOk) {
      fqsRevealedWeft = false;
      fqsRevealedReview = false;
    } else if (!weftRowsOk) {
      fqsRevealedReview = false;
    }

    fqsSyncGeneralPanelCollapsedUi();

    const warpUnlocked = genOk && fqsRevealedWarp;
    fqsSetProgressPanelLocked(document.getElementById("fqs-panel-warp"), !warpUnlocked);

    const weftUnlocked = warpUnlocked && warpRowsOk && fqsRevealedWeft;
    fqsSetProgressPanelLocked(document.getElementById("fqs-panel-weft"), !weftUnlocked);

    const reviewUnlocked = weftUnlocked && weftRowsOk && fqsRevealedReview;
    fqsSetProgressPanelLocked(document.getElementById("fqs-panel-review"), !reviewUnlocked);

    const generalNext = document.getElementById("fqs-general-next");
    if (generalNext) generalNext.disabled = !genOk || fqsGeneralCollapsed;

    const warpNext = document.getElementById("fqs-warp-next");
    if (warpNext) warpNext.disabled = !(warpUnlocked && warpRowsOk);

    const weftNext = document.getElementById("fqs-weft-next");
    if (weftNext) weftNext.disabled = !(weftUnlocked && weftRowsOk);

    const reviewNext = document.getElementById("fqs-review-next");
    if (reviewNext) reviewNext.disabled = !reviewUnlocked;
  }

  let fqsCustomSelectDocListenersBound = false;

  function fqsSetCustomSelectOpen(wrap, open) {
    const panel = wrap.querySelector(".fqs-custom-select__panel");
    const btn = wrap.querySelector(".fqs-custom-select__trigger");
    if (!panel || !btn) return;
    panel.hidden = !open;
    btn.setAttribute("aria-expanded", open ? "true" : "false");
    if (open) wrap.dataset.open = "1";
    else delete wrap.dataset.open;
  }

  function fqsRefreshCustomSelectDisplay(wrap) {
    if (!wrap) return;
    const hidden = wrap.querySelector('input[type="hidden"]');
    const display = wrap.querySelector(".fqs-custom-select__value");
    if (!hidden || !display) return;
    const v = hidden.value;
    let label = "";
    wrap.querySelectorAll(".fqs-custom-select__opt").forEach((b) => {
      const bv = b.dataset.value != null ? b.dataset.value : "";
      const on = bv === v;
      b.setAttribute("aria-selected", on ? "true" : "false");
      b.classList.toggle("is-selected", on);
      if (on) label = b.textContent.trim();
    });
    display.textContent = label || "—";
  }

  function fqsEnsureFqsCustomSelectDocListeners() {
    if (fqsCustomSelectDocListenersBound) return;
    fqsCustomSelectDocListenersBound = true;
    document.addEventListener("mousedown", (e) => {
      document.querySelectorAll(".fqs-custom-select[data-fqs-cs][data-open]").forEach((wrap) => {
        if (!wrap.contains(e.target)) fqsSetCustomSelectOpen(wrap, false);
      });
    });
    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      document.querySelectorAll(".fqs-custom-select[data-fqs-cs][data-open]").forEach((wrap) => {
        fqsSetCustomSelectOpen(wrap, false);
        wrap.querySelector(".fqs-custom-select__trigger")?.focus();
      });
    });
  }

  function fqsBindFabricQualityCustomSelects() {
    fqsEnsureFqsCustomSelectDocListeners();
    document.querySelectorAll(".fqs-custom-select[data-fqs-cs]").forEach((wrap) => {
      if (wrap.dataset.fqsCsUiBound) return;
      wrap.dataset.fqsCsUiBound = "1";
      const btn = wrap.querySelector(".fqs-custom-select__trigger");
      const panel = wrap.querySelector(".fqs-custom-select__panel");
      const hidden = wrap.querySelector('input[type="hidden"]');
      if (!btn || !panel || !hidden) return;
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        const willOpen = !wrap.dataset.open;
        document.querySelectorAll(".fqs-custom-select[data-fqs-cs][data-open]").forEach((w) => {
          if (w !== wrap) fqsSetCustomSelectOpen(w, false);
        });
        fqsSetCustomSelectOpen(wrap, willOpen);
      });
      panel.addEventListener("click", (e) => {
        const opt = e.target.closest(".fqs-custom-select__opt");
        if (!opt) return;
        e.preventDefault();
        hidden.value = opt.dataset.value != null ? opt.dataset.value : "";
        fqsRefreshCustomSelectDisplay(wrap);
        fqsSetCustomSelectOpen(wrap, false);
        hidden.dispatchEvent(new Event("change", { bubbles: true }));
      });
    });
  }

  function fqsPopulateEmailSelect() {
    const hidden = document.getElementById("fqs-email");
    const list = document.getElementById("fqs-email-listbox");
    if (!hidden || !list) return;
    const rows =
      typeof MOCK_FABRIC_EMAIL_OPTIONS !== "undefined" ? MOCK_FABRIC_EMAIL_OPTIONS : [{ value: "", label: "Not specified" }];
    list.innerHTML = rows
      .map(
        (r) =>
          `<li role="presentation"><button type="button" role="option" class="fqs-custom-select__opt" data-value="${escapeAttr(r.value)}">${escapeHtml(r.label)}</button></li>`
      )
      .join("");
    hidden.value = rows[0] ? rows[0].value : "";
    fqsRefreshCustomSelectDisplay(hidden.closest(".fqs-custom-select"));
  }

  function fqsPopulateSavedSpecSelect() {
    const hidden = document.getElementById("fqs-search-old");
    const list = document.getElementById("fqs-search-old-listbox");
    if (!hidden || !list) return;
    const presets = prsFabricQualityPresets();
    const rows = [
      { value: "", label: "Select a saved specification…" },
      ...presets.map((p) => ({ value: p.id, label: p.label })),
    ];
    list.innerHTML = rows
      .map(
        (r) =>
          `<li role="presentation"><button type="button" role="option" class="fqs-custom-select__opt" data-value="${escapeAttr(r.value)}">${escapeHtml(r.label)}</button></li>`
      )
      .join("");
    hidden.value = "";
    fqsRefreshCustomSelectDisplay(hidden.closest(".fqs-custom-select"));
  }

  function initFabricQualityPage() {
    fqsRevealedWarp = false;
    fqsRevealedWeft = false;
    fqsRevealedReview = false;
    fqsGeneralCollapsed = false;
    fqsPopulateEmailSelect();
    fqsPopulateSavedSpecSelect();
    fqsBindFabricQualityCustomSelects();

    const searchOld = document.getElementById("fqs-search-old");
    if (searchOld) {
      searchOld.addEventListener("change", () => {
        const id = searchOld.value;
        if (!id) return;
        const p = prsFabricQualityPresets().find((x) => x.id === id);
        if (p) fqsApplyPreset(p);
      });
    }

    fqsSetAxisRows("warp", [{}]);
    fqsSetAxisRows("weft", [{}]);
    fqsUpdateSummary();
    fqsSyncProgressiveSteps();

    const fabricRoot = document.getElementById("view-fabric-quality");
    if (fabricRoot) {
      fabricRoot.addEventListener("input", () => {
        fqsUpdateSummary();
        fqsSyncProgressiveSteps();
      });
      fabricRoot.addEventListener("change", (e) => {
        if (e.target.matches("select, input")) {
          fqsUpdateSummary();
          fqsSyncProgressiveSteps();
        }
      });
    }

    document.getElementById("fqs-warp-add")?.addEventListener("click", () => {
      const tbody = document.getElementById("fqs-warp-tbody");
      if (!tbody) return;
      const n = tbody.querySelectorAll("tr").length;
      if (n >= 3) return;
      tbody.insertAdjacentHTML("beforeend", fqsAxisRowHtml("warp", n + 1, {}));
      fqsRenumberAxis("warp");
      fqsRefreshAddButtons("warp");
      fqsUpdateSummary();
      fqsSyncProgressiveSteps();
    });

    document.getElementById("fqs-weft-add")?.addEventListener("click", () => {
      const tbody = document.getElementById("fqs-weft-tbody");
      if (!tbody) return;
      const n = tbody.querySelectorAll("tr").length;
      if (n >= 3) return;
      tbody.insertAdjacentHTML("beforeend", fqsAxisRowHtml("weft", n + 1, {}));
      fqsRenumberAxis("weft");
      fqsRefreshAddButtons("weft");
      fqsUpdateSummary();
      fqsSyncProgressiveSteps();
    });

    document.getElementById("fqs-warp-tbody")?.addEventListener("click", fqsOnAxisClick);
    document.getElementById("fqs-weft-tbody")?.addEventListener("click", fqsOnAxisClick);

    document.getElementById("fqs-general-next")?.addEventListener("click", () => {
      if (!fqsIsGeneralComplete()) return;
      fqsGeneralCollapsed = true;
      fqsRevealedWarp = true;
      fqsSyncProgressiveSteps();
      document.getElementById("fqs-panel-warp")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("fqs-general-edit")?.addEventListener("click", () => {
      fqsGeneralCollapsed = false;
      fqsSyncProgressiveSteps();
      document.getElementById("fqs-panel-general")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("fqs-warp-next")?.addEventListener("click", () => {
      const warpTbody = document.getElementById("fqs-warp-tbody");
      if (!fqsIsGeneralComplete() || !fqsIsAxisTbodyComplete(warpTbody)) return;
      fqsRevealedWeft = true;
      fqsSyncProgressiveSteps();
      document.getElementById("fqs-panel-weft")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("fqs-weft-next")?.addEventListener("click", () => {
      const weftTbody = document.getElementById("fqs-weft-tbody");
      if (!fqsRevealedWeft || !fqsIsAxisTbodyComplete(weftTbody)) return;
      fqsRevealedReview = true;
      fqsSyncProgressiveSteps();
      document.getElementById("fqs-panel-review")?.scrollIntoView({ behavior: "smooth", block: "start" });
    });

    document.getElementById("fqs-review-next")?.addEventListener("click", () => {
      fqsUpdateSummary();
      window.alert("Review recorded for this session (demo). Connect workflow rules when backend is ready.");
    });
  }

  const ONGOING_STEPPER_LABELS = {
    "ENQUIRY / REACHOUT": "ENQUIRY",
    "SAMPLE CONFIRMATION": "SAMPLE CONFIRMATION",
    "QUOTATION PROVIDED": "QUOTATION PROVIDED",
    "PRICE CONFIRMATION": "PRICE CONFIRMATION",
    "ORDER CONFIRMATION": "ORDER CONFIRMED",
  };

  function formatOngoingDay(iso) {
    const d = parseISODate(iso);
    return d ? d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" }) : "—";
  }

  function formatOngoingLogAt(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    return Number.isNaN(d.getTime())
      ? iso
      : d.toLocaleString(undefined, { month: "short", day: "numeric", year: "numeric", hour: "2-digit", minute: "2-digit" });
  }

  function ongoingProgressPercent(r) {
    const n = Number(r.progressPercent);
    if (Number.isFinite(n)) return Math.min(100, Math.max(0, n));
    const stages = typeof PIPELINE_STAGES !== "undefined" ? PIPELINE_STAGES : [];
    const i = stages.indexOf(r.progress);
    if (i < 0) return 0;
    return Math.round(((i + 1) / Math.max(stages.length, 1)) * 100);
  }

  /** Per-vendor line count and average % across that vendor’s ongoing rows (same filtered list). */
  function buildOngoingVendorStats(rowList) {
    const byVendor = new Map();
    for (const r of rowList) {
      const k = r.vendor || "";
      if (!byVendor.has(k)) byVendor.set(k, []);
      byVendor.get(k).push(r);
    }
    const stats = new Map();
    for (const [k, list] of byVendor) {
      let sum = 0;
      for (const x of list) sum += ongoingProgressPercent(x);
      stats.set(k, {
        count: list.length,
        overallPct: list.length ? Math.round(sum / list.length) : 0,
      });
    }
    return stats;
  }

  function formatOngoingMaterialCell(r, vendorStats, ctx) {
    const mat = (r.material || "").split("(")[0].trim();
    const escaped = escapeHtml(mat);
    if (ctx && ctx.expandedLine) return escaped;
    const s = vendorStats.get(r.vendor || "");
    const extra = s && s.count > 1 ? s.count - 1 : 0;
    if (extra <= 0) return escaped;
    const title = `${extra} more material line${extra === 1 ? "" : "s"} for this vendor`;
    return `${escaped} <span class="ongoing-material-suffix" title="${escapeAttr(title)}">+${extra}</span>`;
  }

  function formatOngoingProgressTableCell(r, vendorStats, ctx) {
    const s = vendorStats.get(r.vendor || "");
    const pct =
      ctx && ctx.expandedLine
        ? ongoingProgressPercent(r)
        : s && s.count > 1
          ? s.overallPct
          : ongoingProgressPercent(r);
    return `<span class="ongoing-table-pct">${pct}%</span>`;
  }

  function getOngoingCols(vendorStats) {
    return [
      { label: "Order ID", sortKey: "id", render: (r, ctx) => r.id },
      { label: "Vendor name", sortKey: "vendor", render: (r, ctx) => r.vendor },
      {
        label: "Material name",
        sortKey: "material",
        render: (r, ctx) => formatOngoingMaterialCell(r, vendorStats, ctx),
      },
      {
        label: "Current progress",
        sortKey: "_ongoingSortProgress",
        render: (r, ctx) => formatOngoingProgressTableCell(r, vendorStats, ctx),
      },
      { label: "Order received date", sortKey: "received", render: (r, ctx) => r.received },
      { label: "Est. completion date", sortKey: "eta", render: (r, ctx) => r.eta },
      { label: "Order in-charge", sortKey: "incharge", render: (r, ctx) => r.incharge },
    ];
  }

  function ongoingProgressCell(r) {
    const pct = ongoingProgressPercent(r);
    return `<div class="ongoing-progress" aria-label="Progress ${pct} percent">
      <div class="ongoing-progress__track"><span class="ongoing-progress__fill" style="width:${pct}%"></span></div>
      <span class="ongoing-progress__pct">${pct}%</span>
    </div>`;
  }

  function buildOngoingStepperHtml(row) {
    const stages = typeof PIPELINE_STAGES !== "undefined" ? PIPELINE_STAGES : [];
    const currentIdx = stages.indexOf(row.progress);
    const safeIdx = currentIdx >= 0 ? currentIdx : 0;
    const dates = row.stepDates || {};
    const segs = stages.map((name, i) => {
      let state = "ongoing-stepper__seg--upcoming";
      if (i < safeIdx) state = "ongoing-stepper__seg--done";
      else if (i === safeIdx) state = "ongoing-stepper__seg--current";
      const shortLabel = ONGOING_STEPPER_LABELS[name] || name;
      const sub = dates[name] ? formatOngoingDay(dates[name]) : "";
      const lineDone = i < stages.length - 1 && i < safeIdx;
      const line =
        i < stages.length - 1
          ? `<span class="ongoing-stepper__line${lineDone ? " ongoing-stepper__line--done" : ""}" aria-hidden="true"></span>`
          : "";
      return `<div class="ongoing-stepper__seg ${state}" role="listitem">
        <div class="ongoing-stepper__node-row">
          <span class="ongoing-stepper__dot"></span>
          ${line}
        </div>
        <span class="ongoing-stepper__label">${escapeHtml(shortLabel)}</span>
        ${sub ? `<span class="ongoing-stepper__sub">${escapeHtml(sub)}</span>` : ""}
      </div>`;
    });
    return `<div class="ongoing-stepper" role="list">${segs.join("")}</div>`;
  }

  function buildOngoingLogHtml(row) {
    const items = (row.log || []).map((ev) => {
      const cls = ev.done ? "ongoing-log__item--done" : "ongoing-log__item--pending";
      const when = ev.done && ev.at ? formatOngoingLogAt(ev.at) : ev.done ? "" : "Pending";
      const detail = ev.detail ? `<span class="ongoing-log__detail">${escapeHtml(ev.detail)}</span>` : "";
      return `<li class="ongoing-log__item ${cls}">
        <span class="ongoing-log__dot" aria-hidden="true"></span>
        <div class="ongoing-log__body">
          <span class="ongoing-log__title">${escapeHtml(ev.title)}</span>
          ${detail}
          ${when ? `<time class="ongoing-log__time">${escapeHtml(when)}</time>` : ""}
        </div>
      </li>`;
    });
    return `<p class="ongoing-log__heading">Order log &amp; history</p><ul class="ongoing-log">${items.join("")}</ul>`;
  }

  function buildOngoingStageOptionsHtml(currentStage) {
    const stages = typeof PIPELINE_STAGES !== "undefined" ? PIPELINE_STAGES : [];
    return stages
      .map((s) => {
        const lbl = ONGOING_STEPPER_LABELS[s] || s;
        const sel = s === currentStage ? " selected" : "";
        return `<option value="${escapeAttr(s)}"${sel}>${escapeHtml(lbl)}</option>`;
      })
      .join("");
  }

  function applyOngoingStageChange(row, newStage) {
    const stages = typeof PIPELINE_STAGES !== "undefined" ? PIPELINE_STAGES : [];
    if (!row || stages.indexOf(newStage) < 0) return;
    if (row.progress === newStage) return;
    const prev = row.progress;
    const newIdx = stages.indexOf(newStage);
    const today = toIsoDateLocal(new Date());
    row.progress = newStage;
    row.progressPercent = Math.round(((newIdx + 1) / Math.max(stages.length, 1)) * 100);
    if (!row.stepDates) row.stepDates = {};
    stages.forEach((name, i) => {
      if (i <= newIdx && !row.stepDates[name]) row.stepDates[name] = today;
    });
    row.stepDates[newStage] = today;
    if (!row.log) row.log = [];
    if (prev !== newStage) {
      const lbl = ONGOING_STEPPER_LABELS[newStage] || newStage;
      row.log.unshift({
        title: `Stage updated — ${lbl}`,
        at: new Date().toISOString(),
        done: true,
      });
    }
  }

  function ongoingStripChipLabel(r) {
    const raw = r.displayRef || r.id || "";
    const ref = raw.startsWith("TX-") ? raw.replace(/^TX-/, "TX - ") : raw;
    const mat = (r.material || "").split("(")[0].trim();
    return mat ? `#${ref} · ${mat}` : `#${ref}`;
  }

  function buildOngoingOrderStripHtml(allRows, vendorKey, activeOrderId) {
    if (!allRows || allRows.length === 0) return "";
    const vk = vendorKey || "";
    const cohort = vk
      ? allRows.filter((r) => (r.vendor || "") === vk)
      : allRows;
    if (cohort.length === 0) return "";
    const header = vk
      ? `<div class="ongoing-order-strip__header">
          <p class="ongoing-order-strip__vendor">${escapeHtml(vk)}</p>
          <p class="ongoing-order-strip__sub muted small">${cohort.length} material order line${
            cohort.length === 1 ? "" : "s"
          } for this vendor · scroll sideways to switch</p>
        </div>`
      : `<p class="ongoing-order-strip__label muted small">Orders — scroll sideways to switch</p>`;
    const chips = cohort
      .map((r) => {
        const active = r.id === activeOrderId;
        return `<button type="button" role="tab" class="ongoing-order-chip${active ? " is-active" : ""}" data-ongoing-strip="${escapeAttr(
          r.id
        )}" aria-selected="${active ? "true" : "false"}">${escapeHtml(ongoingStripChipLabel(r))}</button>`;
      })
      .join("");
    return `<div class="ongoing-order-strip-wrap">
      ${header}
      <div class="ongoing-order-strip" role="tablist" aria-label="Material orders for this vendor">${chips}</div>
    </div>`;
  }

  function buildOngoingExpandHtml(row) {
    const editOpen = ongoingEditOpenId === row.id;
    const ref = row.displayRef || row.id;
    const idLine = `#${escapeHtml(ref)} · ${escapeHtml(row.material || "—")}`;
    const qty =
      row.qty != null ? `${Number(row.qty).toLocaleString()} Meters` : "—";
    const etaFmt = formatOngoingDay(row.eta);
    const stageOptions = buildOngoingStageOptionsHtml(row.progress);
    const editPanelHidden = editOpen ? "" : " hidden";
    return `<div class="ongoing-detail ongoing-detail--embedded">
      <section class="ongoing-am" aria-label="Active material order">
        <p class="ongoing-am__eyebrow">Active material order</p>
        <p class="ongoing-am__vendor-line">Vendor: <strong>${escapeHtml(row.vendor || "—")}</strong></p>
        <div class="ongoing-am__metrics">
          <div class="ongoing-am__metric ongoing-am__metric--wide">
            <span class="ongoing-am__metric-label">Order ID / material</span>
            <p class="ongoing-am__metric-value">${idLine}</p>
          </div>
          <div class="ongoing-am__metric">
            <span class="ongoing-am__metric-label">Quantity</span>
            <p class="ongoing-am__metric-value">${escapeHtml(qty)}</p>
          </div>
          <div class="ongoing-am__metric ongoing-am__metric--progress">
            <span class="ongoing-am__metric-label">Progress</span>
            <div class="ongoing-detail__progress-wrap">${ongoingProgressCell(row)}</div>
          </div>
          <div class="ongoing-am__metric">
            <span class="ongoing-am__metric-label">Est. completion</span>
            <p class="ongoing-am__metric-value">${escapeHtml(etaFmt)}</p>
          </div>
        </div>
        <div class="ongoing-am__footer">
          <div class="ongoing-am__log">${buildOngoingLogHtml(row)}</div>
          <div class="ongoing-am__actions">
            <button type="button" class="btn ongoing-am__update" data-ongoing-edit="${escapeAttr(row.id)}" aria-expanded="${editOpen ? "true" : "false"}">
              ${editOpen ? "Close updater" : "Update ORDER"}
            </button>
            <div class="ongoing-edit-panel ongoing-am__edit${editOpen ? " is-open" : ""}"${editPanelHidden}>
              <label class="ongoing-edit-label" for="ongoing-stage-${escapeAttr(row.id)}">Pipeline stage</label>
              <select id="ongoing-stage-${escapeAttr(row.id)}" class="select ongoing-edit-select">${stageOptions}</select>
              <div class="ongoing-edit-actions">
                <button type="button" class="btn btn--ongoing-ok" data-ongoing-ok="${escapeAttr(row.id)}">OK</button>
                <button type="button" class="btn btn--ongoing-cancel" data-ongoing-edit-cancel="${escapeAttr(row.id)}">Cancel</button>
              </div>
            </div>
            <p class="ongoing-am__hint muted small">Pick a stage and confirm. Progress and log update for this session.</p>
          </div>
        </div>
      </section>

      <div class="ongoing-detail__card ongoing-detail__card--pipeline">
        <p class="ongoing-detail__card-label">Pipeline overview</p>
        ${buildOngoingStepperHtml(row)}
      </div>
    </div>`;
  }

  function syncOngoingDrawerOrderId(rows) {
    if (!ongoingExpandedId) {
      ongoingDrawerOrderId = null;
      return;
    }
    const anchor = rows.find((x) => x.id === ongoingExpandedId);
    if (!anchor) {
      ongoingDrawerOrderId = null;
      return;
    }
    const av = anchor.vendor || "";
    const drawer = ongoingDrawerOrderId ? rows.find((x) => x.id === ongoingDrawerOrderId) : null;
    if (!drawer || (drawer.vendor || "") !== av) {
      ongoingDrawerOrderId = ongoingExpandedId;
    }
  }

  function renderOngoingTableBody(table, rows) {
    syncOngoingDrawerOrderId(rows);
    const tbody = table.querySelector("tbody");
    const cols = tableColumns.ongoing || getOngoingCols(buildOngoingVendorStats(rows));
    const colCount = cols.length;
    if (rows.length === 0) {
      tbody.innerHTML = `<tr><td colspan="${colCount}" class="muted">No rows in this range.</td></tr>`;
      return;
    }
    const vendorStats = buildOngoingVendorStats(
      rows.map(({ _ongoingSortProgress, ...rest }) => rest)
    );
    const colsResolved = tableColumns.ongoing || getOngoingCols(vendorStats);
    const frag = rows
      .map((r) => {
        const open = ongoingExpandedId === r.id;
        if (open) {
          const drawerRow =
            rows.find((x) => x.id === ongoingDrawerOrderId) || r;
          const cells = colsResolved
            .map((c) => `<div class="ongoing-inline__cell">${c.render(drawerRow, { expandedLine: true })}</div>`)
            .join("");
          const vk = r.vendor || "";
          return `<tr class="ongoing-row ongoing-row--open is-expanded" data-ongoing-id="${escapeAttr(
            r.id
          )}" tabindex="0" aria-expanded="true" role="button">
            <td colspan="${colCount}">
              <div class="ongoing-inline">
                <div class="ongoing-inline__cells">${cells}</div>
                <div class="ongoing-inline__drawer">${buildOngoingOrderStripHtml(
                  rows,
                  vk,
                  ongoingDrawerOrderId || r.id
                )}${buildOngoingExpandHtml(drawerRow)}</div>
              </div>
            </td>
          </tr>`;
        }
        const tds = colsResolved.map((c) => `<td>${c.render(r)}</td>`).join("");
        return `<tr class="ongoing-row" data-ongoing-id="${escapeAttr(
          r.id
        )}" tabindex="0" aria-expanded="false" role="button">${tds}</tr>`;
      })
      .join("");
    tbody.innerHTML = frag;
  }

  function setOngoingExpanded(id) {
    if (ongoingExpandedId === id) {
      ongoingExpandedId = null;
      ongoingDrawerOrderId = null;
      ongoingEditOpenId = null;
    } else {
      ongoingExpandedId = id;
      ongoingDrawerOrderId = id;
      ongoingEditOpenId = null;
    }
    const table = document.getElementById("table-ongoing");
    if (table) renderOngoingTableBody(table, tableRowCache.ongoing || []);
  }

  function bindOngoingTable() {
    const table = document.getElementById("table-ongoing");
    if (!table) return;
    table.addEventListener("click", (e) => {
      const okBtn = e.target.closest("[data-ongoing-ok]");
      if (okBtn && table.contains(okBtn)) {
        e.stopPropagation();
        const id = okBtn.getAttribute("data-ongoing-ok");
        const sel = document.getElementById(`ongoing-stage-${id}`);
        const row = (tableRowCache.ongoing || []).find((r) => r.id === id);
        if (row && sel) applyOngoingStageChange(row, sel.value);
        ongoingEditOpenId = null;
        renderOngoingTableBody(table, tableRowCache.ongoing || []);
        return;
      }
      const cancelBtn = e.target.closest("[data-ongoing-edit-cancel]");
      if (cancelBtn && table.contains(cancelBtn)) {
        e.stopPropagation();
        ongoingEditOpenId = null;
        renderOngoingTableBody(table, tableRowCache.ongoing || []);
        return;
      }
      const editBtn = e.target.closest("[data-ongoing-edit]");
      if (editBtn && table.contains(editBtn)) {
        e.stopPropagation();
        const id = editBtn.getAttribute("data-ongoing-edit");
        ongoingEditOpenId = ongoingEditOpenId === id ? null : id;
        renderOngoingTableBody(table, tableRowCache.ongoing || []);
        return;
      }
      const th = e.target.closest("th[data-sort]");
      if (th && table.contains(th)) {
        const key = th.dataset.sort;
        const current = table.dataset.sortKey;
        const dir = current === `${key}-asc` ? "desc" : "asc";
        table.dataset.sortKey = `${key}-${dir}`;
        table.querySelectorAll("th[data-sort]").forEach((h) => h.classList.remove("is-sorted"));
        th.classList.add("is-sorted");
        const sorted = sortRows(tableRowCache.ongoing, key, dir);
        tableRowCache.ongoing = sorted;
        renderOngoingTableBody(table, sorted);
        return;
      }
      const stripBtn = e.target.closest("[data-ongoing-strip]");
      if (stripBtn && table.contains(stripBtn)) {
        e.stopPropagation();
        const id = stripBtn.getAttribute("data-ongoing-strip");
        if (id && id !== ongoingDrawerOrderId) {
          ongoingDrawerOrderId = id;
          ongoingEditOpenId = null;
          renderOngoingTableBody(table, tableRowCache.ongoing || []);
          requestAnimationFrame(() => {
            const next = [...document.querySelectorAll("[data-ongoing-strip]")].find(
              (el) => el.getAttribute("data-ongoing-strip") === id
            );
            next?.scrollIntoView({ behavior: "smooth", inline: "nearest", block: "nearest" });
          });
        }
        return;
      }
      const tr = e.target.closest("tr.ongoing-row");
      if (!tr || !table.contains(tr)) return;
      if (e.target.closest(".ongoing-inline__drawer")) return;
      const id = tr.getAttribute("data-ongoing-id");
      if (id) setOngoingExpanded(id);
    });
    table.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        if (ongoingEditOpenId) {
          e.preventDefault();
          ongoingEditOpenId = null;
          renderOngoingTableBody(table, tableRowCache.ongoing || []);
          return;
        }
        if (ongoingExpandedId) {
          e.preventDefault();
          ongoingExpandedId = null;
          ongoingDrawerOrderId = null;
          renderOngoingTableBody(table, tableRowCache.ongoing || []);
        }
        return;
      }
      if (e.key !== "Enter" && e.key !== " ") return;
      const tr = e.target.closest("tr.ongoing-row");
      if (!tr || !table.contains(tr)) return;
      e.preventDefault();
      const id = tr.getAttribute("data-ongoing-id");
      if (id) setOngoingExpanded(id);
    });
  }

  /* --- Mail --- */
  function formatSentAt(iso) {
    const d = new Date(iso);
    return Number.isNaN(d.getTime()) ? iso : d.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });
  }

  /** Inbox “Send / generate mail” — only after compose has body text or a saved draft is loaded in Compose */
  function refreshMailSendGenerateEnabled() {
    const btn = document.getElementById("btn-mail-send-generate");
    if (!btn) return;
    const body = document.getElementById("mail-draft-body")?.value?.trim() || "";
    const editingSaved = !!document.getElementById("mail-compose-draft-id")?.value?.trim();
    const ready = body.length > 0 || editingSaved;
    btn.disabled = !ready;
    btn.setAttribute("aria-disabled", ready ? "false" : "true");
    btn.title = ready ? "" : "Save a draft or add message text in Compose first";
  }

  function setMailFolder(folder) {
    mailFolder = folder;
    const viewMail = document.getElementById("view-mail");
    if (viewMail && viewMail.classList.contains("is-visible")) {
      document.getElementById("page-title").textContent = mailFolderLabel(folder);
    }

    document.querySelectorAll(".nav-mail__item").forEach((btn) => {
      const on = btn.dataset.mailFolder === folder;
      btn.classList.toggle("is-active", on);
      if (on) btn.setAttribute("aria-current", "page");
      else btn.removeAttribute("aria-current");
    });

    const inboxWrap = document.getElementById("mail-inbox-wrap");
    const draftUl = document.getElementById("mail-draft-list");
    const sentUl = document.getElementById("mail-sent-list");
    const composeHint = document.getElementById("mail-context-compose-hint");
    if (inboxWrap) inboxWrap.hidden = folder !== "inbox";
    if (draftUl) draftUl.hidden = folder !== "draft";
    if (sentUl) sentUl.hidden = folder !== "sent";
    if (composeHint) composeHint.hidden = folder !== "compose";

    document.querySelectorAll(".mail-pane").forEach((pane) => {
      pane.hidden = pane.getAttribute("data-mail-pane") !== folder;
    });

    if (folder === "sent") updateSentDetail();
    refreshMailSendGenerateEnabled();
  }

  function initNavMail() {
    const wrap = document.getElementById("nav-mail");
    const trig = document.getElementById("nav-mail-trigger");
    const panel = document.getElementById("nav-mail-menu");
    if (!wrap || !trig || !panel) return;

    trig.addEventListener("click", (e) => {
      e.stopPropagation();
      const willOpen = panel.hidden;
      if (willOpen) {
        showView("mail");
        setMailFolder("inbox");
        openMailMenu();
      } else {
        closeMailMenu();
      }
    });

    wrap.addEventListener("click", (e) => e.stopPropagation());

    panel.querySelectorAll("[data-mail-folder]").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        showView("mail");
        setMailFolder(btn.getAttribute("data-mail-folder"));
      });
    });

    document.addEventListener("click", () => closeMailMenu());

    document.querySelector(".main")?.addEventListener("pointerdown", () => {
      if (!panel.hidden) closeMailMenu();
    });

    document.addEventListener("keydown", (e) => {
      if (e.key !== "Escape") return;
      if (panel.hidden) return;
      closeMailMenu();
      trig.focus();
    });

  }

  function mailInboxDetailTitle(m) {
    if (m.ref) return `Material Specs #${m.ref} Enquiry`;
    return m.subject;
  }

  function formatMailDetailMeta(m) {
    const d = parseISODate(m.date);
    const datePart = d
      ? d.toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric", year: "numeric" })
      : m.date;
    const timePart = d
      ? d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" })
      : "";
    return timePart ? `From: ${m.from} · ${datePart}, ${timePart}` : `From: ${m.from} · ${datePart}`;
  }

  /** Fade the inbox heading as the thread list scrolls (keeps focus on messages). */
  function syncMailInboxHeadingFade() {
    const list = document.getElementById("mail-inbox");
    const heading = document.getElementById("mail-inbox-heading");
    if (!list || !heading) return;
    const y = list.scrollTop;
    const fadePx = 64;
    const opacity = Math.max(0, Math.min(1, 1 - y / fadePx));
    if (opacity >= 1) heading.style.removeProperty("opacity");
    else heading.style.opacity = String(opacity);
  }

  let mailInboxHeadingScrollBound = false;

  function initMailInboxHeadingFade() {
    const list = document.getElementById("mail-inbox");
    if (!list || mailInboxHeadingScrollBound) return;
    mailInboxHeadingScrollBound = true;
    list.addEventListener("scroll", syncMailInboxHeadingFade, { passive: true });
    syncMailInboxHeadingFade();
  }

  function renderMailInbox() {
    const ul = document.getElementById("mail-inbox");
    const countEl = document.getElementById("mail-inbox-count");
    if (countEl) {
      const n = MOCK_MAIL.length;
      countEl.textContent = String(n);
      countEl.setAttribute("aria-label", `${n} message${n === 1 ? "" : "s"} in inbox`);
    }
    if (!ul) return;
    ul.innerHTML = MOCK_MAIL.map((m) => {
      const unread = Boolean(m.unread);
      const unreadCls = unread ? " is-unread" : "";
      const titleAttr = unread ? ` title="Unread message"` : "";
      return `
      <li>
        <button type="button" class="mail-inbox-card${unreadCls}" data-id="${escapeAttr(m.id)}"${titleAttr}>
          ${unread ? '<span class="mail-inbox-card__dot" aria-hidden="true"></span>' : ""}
          <span class="mail-inbox-card__subject">${escapeHtml(m.subject)}</span>
          <span class="mail-inbox-card__meta">${escapeHtml(m.from)} · ${escapeHtml(m.date)}</span>
        </button>
      </li>`;
    }).join("");
    ul.querySelectorAll(".mail-inbox-card").forEach((btn) => {
      btn.addEventListener("click", () => selectMail(btn.dataset.id));
    });
    initMailInboxHeadingFade();
    syncMailInboxHeadingFade();
  }

  function renderDraftList() {
    const ul = document.getElementById("mail-draft-list");
    if (!ul) return;
    ul.innerHTML = mailDraftsRuntime
      .map(
        (d) => `
      <li>
        <button type="button" class="mail-thread-btn" data-draft-id="${escapeAttr(d.id)}">
          <strong>${escapeHtml(d.subject)}</strong>
          <span>${escapeHtml(d.to)} · ${escapeHtml(d.updatedAt || "—")}</span>
        </button>
      </li>`
      )
      .join("");
    ul.querySelectorAll(".mail-thread-btn").forEach((btn) => {
      btn.addEventListener("click", () => selectDraft(btn.dataset.draftId));
    });
  }

  function renderSentList() {
    const ul = document.getElementById("mail-sent-list");
    if (!ul) return;
    if (mailSentSession.length === 0) {
      ul.innerHTML = "";
      return;
    }
    const rows = [...mailSentSession].reverse();
    ul.innerHTML = rows
      .map(
        (s) => `
      <li>
        <button type="button" class="mail-thread-btn" data-sent-id="${escapeAttr(s.id)}">
          <strong>${escapeHtml(s.subject)}</strong>
          <span>${escapeHtml(s.to)} · ${escapeHtml(formatSentAt(s.sentAt))}</span>
        </button>
      </li>`
      )
      .join("");
    ul.querySelectorAll(".mail-thread-btn").forEach((btn) => {
      btn.addEventListener("click", () => selectSent(btn.dataset.sentId));
    });
  }

  function selectMail(id) {
    selectedMailId = id;
    const m = MOCK_MAIL.find((x) => x.id === id);
    if (m) m.unread = false;
    renderMailInbox();
    document.querySelectorAll("#mail-inbox .mail-inbox-card").forEach((b) => {
      b.classList.toggle("is-selected", b.dataset.id === id);
    });
    const summaryEl = document.getElementById("mail-summary");
    const replyBtn = document.getElementById("btn-mail-reply-compose");
    const detailCard = document.getElementById("mail-detail-card");
    const detailEmpty = document.getElementById("mail-detail-empty");
    const titleEl = document.getElementById("mail-detail-title");
    const metaEl = document.getElementById("mail-detail-meta");
    const bodyEl = document.getElementById("mail-detail-body");
    if (!m) {
      if (summaryEl) summaryEl.textContent = "Select a thread to read a short summary of what the sender needs.";
      if (replyBtn) replyBtn.disabled = true;
      if (detailCard) detailCard.hidden = true;
      if (detailEmpty) detailEmpty.hidden = false;
      return;
    }
    if (summaryEl) summaryEl.textContent = m.summary;
    if (titleEl) titleEl.textContent = mailInboxDetailTitle(m);
    if (metaEl) metaEl.textContent = formatMailDetailMeta(m);
    if (bodyEl) bodyEl.textContent = m.body || m.summary;
    if (detailCard) detailCard.hidden = false;
    if (detailEmpty) detailEmpty.hidden = true;
    if (replyBtn) replyBtn.disabled = false;
  }

  function selectDraft(id) {
    selectedDraftId = id;
    document.querySelectorAll("#mail-draft-list .mail-thread-btn").forEach((b) => {
      b.classList.toggle("is-selected", b.dataset.draftId === id);
    });
    const preview = document.getElementById("mail-draft-preview");
    const editBtn = document.getElementById("btn-mail-edit-compose");
    const d = mailDraftsRuntime.find((x) => x.id === id);
    if (!d || !preview) {
      if (preview) preview.textContent = "Select a saved draft or save one from Compose.";
      if (editBtn) editBtn.disabled = true;
      return;
    }
    preview.innerHTML = `<p style="margin:0 0 0.5rem;"><strong>To</strong> ${escapeHtml(d.to)}</p><p style="margin:0 0 0.5rem;"><strong>Subject</strong> ${escapeHtml(d.subject)}</p><div style="white-space:pre-wrap;">${escapeHtml(d.body)}</div>`;
    if (editBtn) editBtn.disabled = false;
  }

  function updateSentDetail() {
    const empty = document.getElementById("mail-sent-empty");
    const bodyEl = document.getElementById("mail-sent-detail-body");
    if (!empty || !bodyEl) return;

    if (mailSentSession.length === 0) {
      empty.hidden = false;
      bodyEl.hidden = true;
      bodyEl.innerHTML = "";
      return;
    }

    empty.hidden = true;
    bodyEl.hidden = false;
    const s = mailSentSession.find((x) => x.id === selectedSentId);
    if (!s) {
      bodyEl.innerHTML = `<p class="muted" style="margin:0;">Select a message from the list.</p>`;
      return;
    }
    bodyEl.innerHTML = `<p class="mail-sent-meta">To ${escapeHtml(s.to)} · ${escapeHtml(formatSentAt(s.sentAt))}</p><h3 class="panel__title" style="font-size:1rem;margin:0 0 0.5rem;">${escapeHtml(s.subject)}</h3><pre class="mail-sent-body" style="margin:0;font-family:inherit;white-space:pre-wrap;">${escapeHtml(s.body)}</pre>`;
  }

  function selectSent(id) {
    selectedSentId = id;
    document.querySelectorAll("#mail-sent-list .mail-thread-btn").forEach((b) => {
      b.classList.toggle("is-selected", b.dataset.sentId === id);
    });
    updateSentDetail();
  }

  function getMailFinalPrice() {
    const mq = document.getElementById("mail-quote-final");
    if (mq && mq.value !== "") return Number(mq.value) || 0;
    return Number(document.getElementById("quote-final").value) || (lastQuote && lastQuote.mid) || 0;
  }

  function initMailWorkbench() {
    initMailInboxHeadingFade();

    document.getElementById("mail-sync-pricing").addEventListener("click", () => {
      syncPricingPageToMailForm();
    });

    document.getElementById("btn-mail-calc").addEventListener("click", () => {
      runMailPricing();
    });

    function openComposeFromSelectedInbox() {
      const m = MOCK_MAIL.find((x) => x.id === selectedMailId);
      if (!m) return;
      document.getElementById("mail-draft-to").value = m.from;
      document.getElementById("mail-draft-subject").value = `Re: ${m.subject}`;
      document.getElementById("mail-draft-body").value = "";
      document.getElementById("mail-compose-draft-id").value = "";
      const mqf = document.getElementById("mail-quote-final");
      if (mqf) mqf.value = lastQuote ? lastQuote.mid.toFixed(2) : "";
      const mqr = document.getElementById("mail-quote-range");
      if (mqr) {
        if (lastQuote)
          mqr.textContent = `Suggested range: ${formatMoney(lastQuote.low)} — ${formatMoney(lastQuote.high)} per metre`;
        else mqr.textContent = "";
      }
      syncMailPricingMini();
      setMailFolder("compose");
      refreshMailSendGenerateEnabled();
    }

    document.getElementById("btn-mail-reply-compose").addEventListener("click", () => {
      openComposeFromSelectedInbox();
    });

    document.getElementById("btn-mail-send-generate")?.addEventListener("click", () => {
      const btn = document.getElementById("btn-mail-send-generate");
      if (btn?.disabled) return;
      if (selectedMailId && MOCK_MAIL.some((x) => x.id === selectedMailId)) {
        openComposeFromSelectedInbox();
      } else {
        setMailFolder("compose");
      }
    });

    document.getElementById("btn-mail-compose-draft")?.addEventListener("click", () => {
      setMailFolder("compose");
    });

    document.getElementById("btn-mail-older")?.addEventListener("click", () => {
      window.alert("Older mail archive is not connected yet — demo inbox only.");
    });

    document.getElementById("btn-mail-sample-input")?.addEventListener("click", () => {
      setMailFolder("compose");
      window.setTimeout(() => document.getElementById("mpf-sample")?.focus(), 0);
    });

    document.getElementById("btn-mail-attachment")?.addEventListener("click", () => {
      window.alert("Attach files from your device when you use Copy or PDF in Compose, or connect storage later.");
    });

    document.getElementById("btn-mail-edit-compose").addEventListener("click", () => {
      const d = mailDraftsRuntime.find((x) => x.id === selectedDraftId);
      if (!d) return;
      document.getElementById("mail-draft-to").value = d.to;
      document.getElementById("mail-draft-subject").value = d.subject;
      document.getElementById("mail-draft-body").value = d.body;
      document.getElementById("mail-compose-draft-id").value = d.id;
      setMailFolder("compose");
    });

    document.getElementById("btn-mail-save-draft").addEventListener("click", () => {
      const to = document.getElementById("mail-draft-to").value.trim();
      const subject = document.getElementById("mail-draft-subject").value.trim();
      const body = document.getElementById("mail-draft-body").value;
      const existingId = document.getElementById("mail-compose-draft-id").value.trim();
      if (!to && !subject && !body.trim()) return;
      const today = new Date().toISOString().slice(0, 10);
      if (existingId) {
        const d = mailDraftsRuntime.find((x) => x.id === existingId);
        if (d) {
          d.to = to || d.to;
          d.subject = subject || "(No subject)";
          d.body = body;
          d.updatedAt = today;
        }
      } else {
        const nid = `d-${Date.now()}`;
        mailDraftsRuntime.push({
          id: nid,
          to: to || "(No address)",
          subject: subject || "(No subject)",
          body,
          updatedAt: today,
        });
        document.getElementById("mail-compose-draft-id").value = nid;
      }
      renderDraftList();
      refreshMailSendGenerateEnabled();
    });

    document.getElementById("btn-mail-send").addEventListener("click", () => {
      const to = document.getElementById("mail-draft-to").value.trim();
      if (!to) {
        alert("Add a recipient in To before sending.");
        return;
      }
      const subject = document.getElementById("mail-draft-subject").value.trim() || "(No subject)";
      const body = document.getElementById("mail-draft-body").value;
      const draftRowId = document.getElementById("mail-compose-draft-id").value.trim();
      if (draftRowId) {
        const idx = mailDraftsRuntime.findIndex((x) => x.id === draftRowId);
        if (idx >= 0) mailDraftsRuntime.splice(idx, 1);
        renderDraftList();
        if (selectedDraftId === draftRowId) {
          selectedDraftId = null;
          const preview = document.getElementById("mail-draft-preview");
          if (preview) preview.textContent = "Select a saved draft or save one from Compose.";
          const editBtn = document.getElementById("btn-mail-edit-compose");
          if (editBtn) editBtn.disabled = true;
        }
      }
      const sid = `s-${Date.now()}`;
      mailSentSession.push({
        id: sid,
        to,
        subject,
        body,
        sentAt: new Date().toISOString(),
      });
      document.getElementById("mail-compose-draft-id").value = "";
      document.getElementById("mail-draft-to").value = "";
      document.getElementById("mail-draft-subject").value = "";
      document.getElementById("mail-draft-body").value = "";
      renderSentList();
      selectedSentId = sid;
      selectSent(sid);
      setMailFolder("sent");
      refreshMailSendGenerateEnabled();
    });

    document.getElementById("btn-mail-clear").addEventListener("click", () => {
      document.getElementById("mail-draft-to").value = "";
      document.getElementById("mail-draft-subject").value = "";
      document.getElementById("mail-draft-body").value = "";
      document.getElementById("mail-compose-draft-id").value = "";
      refreshMailSendGenerateEnabled();
    });

    document.getElementById("btn-insert-quote").addEventListener("click", () => {
      if (!lastQuote) {
        if (document.getElementById("mpf-material")) runMailPricing();
        else runPricing();
      }
      const final = getMailFinalPrice() || lastQuote.mid;
      const mat = lastQuote.inputs.material;
      const w = lastQuote.inputs.width;
      const block = `\n\n---\nQuotation summary\nCloth type / material: ${mat}\nPrice per metre: ${formatMoney(final)}\nWidth: ${w} m\nSample / pattern notes: ${lastQuote.inputs.sample || "—"}\n(PDF export available for attachment.)\n`;
      document.getElementById("mail-draft-body").value += block;
      refreshMailSendGenerateEnabled();
    });

    ["mail-draft-to", "mail-draft-subject", "mail-draft-body"].forEach((id) => {
      document.getElementById(id)?.addEventListener("input", refreshMailSendGenerateEnabled);
    });

    document.getElementById("btn-mail-pdf").addEventListener("click", () => {
      if (!lastQuote) runMailPricing();
      const final = getMailFinalPrice() || lastQuote.mid;
      const inputs = lastQuote.inputs;
      const existing = document.getElementById("print-root");
      if (existing) existing.remove();
      const root = document.createElement("div");
      root.id = "print-root";
      root.style.cssText = "background:#fff;color:#111;padding:2rem;";
      root.innerHTML = `
      <h1 style="font-family:system-ui;margin:0 0 0.5rem;">Elite Textile — Quotation</h1>
      <p style="margin:0 0 1rem;color:#333;">${new Date().toLocaleString()}</p>
      <table style="width:100%;border-collapse:collapse;font-size:14px;">
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Material</strong></td><td style="padding:6px;border:1px solid #ccc;">${escapeHtml(inputs.material)}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Width</strong></td><td style="padding:6px;border:1px solid #ccc;">${inputs.width} m</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Read / pick</strong></td><td style="padding:6px;border:1px solid #ccc;">${inputs.readpick}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Sample type</strong></td><td style="padding:6px;border:1px solid #ccc;">${escapeHtml(inputs.sample)}</td></tr>
        <tr><td style="padding:6px;border:1px solid #ccc;"><strong>Quoted price / m</strong></td><td style="padding:6px;border:1px solid #ccc;font-size:18px;"><strong>${formatMoney(final)}</strong></td></tr>
      </table>
      <p style="margin-top:1rem;font-size:12px;color:#666;">Range was ${formatMoney(lastQuote.low)} — ${formatMoney(lastQuote.high)} / m before final edit.</p>
    `;
      document.body.appendChild(root);
      window.print();
      root.remove();
    });

    document.getElementById("btn-copy-draft").addEventListener("click", async () => {
      const text = [
        `To: ${document.getElementById("mail-draft-to").value}`,
        `Subject: ${document.getElementById("mail-draft-subject").value}`,
        "",
        document.getElementById("mail-draft-body").value,
      ].join("\n");
      try {
        await navigator.clipboard.writeText(text);
        const btn = document.getElementById("btn-copy-draft");
        const prev = btn.textContent;
        btn.textContent = "Copied";
        setTimeout(() => {
          btn.textContent = prev;
        }, 1500);
      } catch {
        alert("Could not copy — select the body and copy manually.");
      }
    });
  }

  /* --- Wire refresh --- */
  function refreshAll() {
    updateDateLabel();
    refreshOverview();
    refreshCompleted();
    refreshOngoing();
    refreshReceived();
    refreshAlerts();
  }

  function onGlobalRangeInputChange() {
    setGlobalDateAll(false);
    refreshAll();
  }
  document.getElementById("global-start").addEventListener("change", onGlobalRangeInputChange);
  document.getElementById("global-end").addEventListener("change", onGlobalRangeInputChange);
  document.getElementById("overview-awaiting").addEventListener("change", refreshOverview);
  document.getElementById("overview-filter-stage").addEventListener("change", refreshOverview);
  document.getElementById("overview-filter-vendor").addEventListener("change", refreshOverview);
  document.getElementById("overview-awaiting-search").addEventListener("input", () => renderAwaitingTable(false));

  document.getElementById("table-awaiting").addEventListener("click", (e) => {
    const btn = e.target.closest("[data-awaiting-detail]");
    if (!btn) return;
    const id = btn.getAttribute("data-awaiting-detail");
    const row = tableRowCache.awaiting.find((r) => r.id === id);
    if (row) openAwaitingModal(row);
  });

  document.querySelectorAll("[data-modal-close]").forEach((el) => {
    el.addEventListener("click", closeAwaitingModal);
  });
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;
    const modal = document.getElementById("awaiting-modal");
    if (modal && modal.classList.contains("is-open")) closeAwaitingModal();
  });
  ["filter-completed-vendor", "filter-completed-material", "filter-completed-type"].forEach(
    (id) => document.getElementById(id).addEventListener("change", refreshCompleted)
  );
  ["filter-ongoing-vendor", "filter-ongoing-material", "filter-ongoing-progress", "filter-ongoing-incharge"].forEach(
    (id) => document.getElementById(id)?.addEventListener("change", refreshOngoing)
  );

  /* Init */
  setDefaultDates();
  fillSelect(
    "filter-completed-vendor",
    [...new Set(MOCK_COMPLETED.map((r) => r.vendor))].sort(),
    "All vendors"
  );
  fillSelect(
    "filter-completed-material",
    [...new Set(MOCK_COMPLETED.map((r) => r.material))].sort(),
    "All materials"
  );
  fillSelect(
    "filter-completed-type",
    [...new Set(MOCK_COMPLETED.map((r) => r.orderType))].sort(),
    "All order types"
  );

  fillSelect(
    "filter-ongoing-vendor",
    [...new Set(MOCK_ONGOING.map((r) => r.vendor))].sort(),
    "All vendors"
  );
  fillSelect(
    "filter-ongoing-material",
    [...new Set(MOCK_ONGOING.map((r) => r.material))].sort(),
    "All materials"
  );
  fillSelect(
    "filter-ongoing-progress",
    typeof PIPELINE_STAGES !== "undefined" ? [...PIPELINE_STAGES] : [],
    "All pipeline stages"
  );
  fillSelect(
    "filter-ongoing-incharge",
    [...new Set(MOCK_ONGOING.map((r) => r.incharge).filter(Boolean))].sort(),
    "All handlers"
  );

  fillOverviewFilterSelects();
  tableColumns.completed = COMPLETED_COLS;
  tableColumns.ongoing = getOngoingCols(buildOngoingVendorStats([]));
  tableColumns.received = RECEIVED_COLS;
  bindDelegatedSort("table-awaiting", "awaiting");
  bindDelegatedSort("table-completed", "completed");
  bindOngoingTable();
  bindDelegatedSort("table-received", "received");

  mailDraftsRuntime = (typeof MOCK_MAIL_DRAFTS !== "undefined" ? MOCK_MAIL_DRAFTS : []).map((d) => ({ ...d }));
  renderMailInbox();
  if (MOCK_MAIL.some((x) => x.id === "m2")) selectMail("m2");
  else if (MOCK_MAIL[0]) selectMail(MOCK_MAIL[0].id);
  renderDraftList();
  renderSentList();
  initNavMail();
  initFabricQualityPage();
  initPricingPage();
  initMailWorkbench();
  initSidebarToggle();
  initDateRangePopover();
  showView("overview");
  refreshAll();
})();
