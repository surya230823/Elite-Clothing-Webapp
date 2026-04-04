/**
 * Dashboard behaviour: navigation, tables, filters, pricing calculator, mail draft.
 * Requires: index.html markup, data.js (MOCK_* globals), auth.js (optional for mail).
 */
(function () {
  const TITLES = {
    overview: "Overview",
    completed: "Completed",
    ongoing: "Ongoing",
    pricing: "Pricing",
    alerts: "Alerts",
    received: "Received",
    mail: "Mail",
  };

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

  function setDefaultDates() {
    const end = new Date();
    const start = new Date();
    start.setDate(start.getDate() - 29);
    document.getElementById("global-end").value = end.toISOString().slice(0, 10);
    document.getElementById("global-start").value = start.toISOString().slice(0, 10);
  }

  function updateDateLabel() {
    const { start, end } = getGlobalRange();
    const fmt = (d) =>
      d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
    const el = document.getElementById("global-date-line");
    if (el) el.textContent = `${fmt(start)} → ${fmt(end)}`;
  }

  function toIsoDateLocal(d) {
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const day = d.getDate();
    return `${y}-${String(m).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  function applyDateRangePreset(preset) {
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
    const startEl = document.getElementById("global-start");
    const endEl = document.getElementById("global-end");
    if (!startEl || !endEl) return;
    startEl.value = toIsoDateLocal(start);
    endEl.value = toIsoDateLocal(end);
    startEl.dispatchEvent(new Event("change", { bubbles: true }));
    endEl.dispatchEvent(new Event("change", { bubbles: true }));
  }

  /* --- Navigation --- */
  function showView(key) {
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
    document.getElementById("page-title").textContent = TITLES[key] || key;
    updateDateLabel();
    const stageWrap = document.getElementById("topbar-stage-wrap");
    if (stageWrap) stageWrap.hidden = key !== "overview";
  }

  document.querySelectorAll(".nav-item").forEach((btn) => {
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
    const range = getGlobalRange();
    const awaitingFilter = document.getElementById("overview-awaiting").value;
    const stageFilter = document.getElementById("overview-filter-stage")?.value || "";
    const vendorFilter = document.getElementById("overview-filter-vendor")?.value || "";

    let awaiting = MOCK_AWAITING_ELITE.filter((a) =>
      a.opened ? inRange(a.opened, range.start, range.end) : true
    );
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

  const ONGOING_COLS = [
    { label: "Order ID", sortKey: "id", render: (r) => r.id },
    { label: "Vendor name", sortKey: "vendor", render: (r) => r.vendor },
    { label: "Material name", sortKey: "material", render: (r) => r.material },
    {
      label: "Current progress",
      sortKey: "progress",
      render: (r) => `<span class="${badgeClass(r.progress)}">${r.progress}</span>`,
    },
    { label: "Order received date", sortKey: "received", render: (r) => r.received },
    { label: "Est. completion date", sortKey: "eta", render: (r) => r.eta },
    { label: "Order in-charge", sortKey: "incharge", render: (r) => r.incharge },
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
    const { start, end } = getGlobalRange();
    let rows = MOCK_COMPLETED.filter((r) => inRange(r.completed, start, end));
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
    const { start, end } = getGlobalRange();
    const rows = MOCK_ONGOING.filter((r) => inRange(r.received, start, end));
    tableColumns.ongoing = ONGOING_COLS;
    tableRowCache.ongoing = rows;
    document.getElementById("table-ongoing").dataset.sortKey = "";
    document.querySelectorAll("#table-ongoing th[data-sort]").forEach((h) => h.classList.remove("is-sorted"));
    const table = document.getElementById("table-ongoing");
    renderTableHead(table, ONGOING_COLS);
    renderTableBody(table, rows, ONGOING_COLS);
  }

  /* --- Received --- */
  function refreshReceived() {
    const { start, end } = getGlobalRange();
    const rows = MOCK_RECEIVED.filter((r) => inRange(r.received, start, end));
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

  /* --- Pricing --- */
  let lastQuote = null;

  function readPricingForm() {
    return {
      material: document.getElementById("pf-material").value.trim(),
      readpick: Number(document.getElementById("pf-readpick").value) || 0,
      rawprice: Number(document.getElementById("pf-rawprice").value) || 0,
      width: Number(document.getElementById("pf-width").value) || 0,
      sample: document.getElementById("pf-sample").value.trim(),
      crimp: Number(document.getElementById("pf-crimp").value) || 0,
      shrink: Number(document.getElementById("pf-shrink").value) || 0,
      buffer: Number(document.getElementById("pf-buffer").value) || 0,
    };
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
    const pairs = [
      ["pf-material", "mpf-material"],
      ["pf-readpick", "mpf-readpick"],
      ["pf-rawprice", "mpf-rawprice"],
      ["pf-width", "mpf-width"],
      ["pf-sample", "mpf-sample"],
      ["pf-crimp", "mpf-crimp"],
      ["pf-shrink", "mpf-shrink"],
      ["pf-buffer", "mpf-buffer"],
    ];
    pairs.forEach(([a, b]) => {
      const src = document.getElementById(a);
      const dst = document.getElementById(b);
      if (src && dst) dst.value = src.value;
    });
  }

  function syncMailFormToPricingPage() {
    const pairs = [
      ["mpf-material", "pf-material"],
      ["mpf-readpick", "pf-readpick"],
      ["mpf-rawprice", "pf-rawprice"],
      ["mpf-width", "pf-width"],
      ["mpf-sample", "pf-sample"],
      ["mpf-crimp", "pf-crimp"],
      ["mpf-shrink", "pf-shrink"],
      ["mpf-buffer", "pf-buffer"],
    ];
    pairs.forEach(([a, b]) => {
      const src = document.getElementById(a);
      const dst = document.getElementById(b);
      if (src && dst) dst.value = src.value;
    });
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

  function runPricing(e) {
    if (e) e.preventDefault();
    const inputs = readPricingForm();
    lastQuote = { inputs, ...computeQuote(inputs) };
    const { low, mid, high, beforeBuffer } = lastQuote;
    document.getElementById("quote-range").textContent = `Suggested range: ${formatMoney(low)} — ${formatMoney(high)} per metre (mid ${formatMoney(mid)})`;
    document.getElementById("quote-final").value = mid.toFixed(2);
    document.getElementById("quote-breakdown").textContent =
      `After crimp & shrink, before buffer: ${formatMoney(beforeBuffer)} / m. Adjust the final figure for margin; export PDF to attach to your mail client.`;
    syncMailPricingMini();
    syncPricingPageToMailForm();
    const mqf = document.getElementById("mail-quote-final");
    if (mqf) mqf.value = mid.toFixed(2);
    const mqr = document.getElementById("mail-quote-range");
    if (mqr)
      mqr.textContent = `Range: ${formatMoney(low)} — ${formatMoney(high)} / m`;
  }

  function runMailPricing() {
    const inputs = readMailPricingForm();
    lastQuote = { inputs, ...computeQuote(inputs) };
    const { low, mid, high } = lastQuote;
    const mqr = document.getElementById("mail-quote-range");
    if (mqr)
      mqr.textContent = `Suggested range: ${formatMoney(low)} — ${formatMoney(high)} per metre (mid ${formatMoney(mid)})`;
    const mqf = document.getElementById("mail-quote-final");
    if (mqf) mqf.value = mid.toFixed(2);
    syncMailFormToPricingPage();
    document.getElementById("quote-range").textContent = `Suggested range: ${formatMoney(low)} — ${formatMoney(high)} per metre (mid ${formatMoney(mid)})`;
    document.getElementById("quote-final").value = mid.toFixed(2);
    syncMailPricingMini();
  }

  function syncMailPricingMini() {
    const mini = document.getElementById("mail-pricing-mini");
    if (!mini) return;
    if (!lastQuote) {
      mini.textContent = "Run Calculate quotation on this page or on Pricing overview.";
      return;
    }
    const m = lastQuote.inputs.material;
    mini.innerHTML = `Cloth type: ${escapeHtml(m)} · Range ${formatMoney(lastQuote.low)} — ${formatMoney(lastQuote.high)} / m`;
  }

  document.getElementById("pricing-form").addEventListener("submit", runPricing);

  document.getElementById("btn-print-quote").addEventListener("click", () => {
    if (!lastQuote) runPricing();
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

  /* --- Mail --- */
  let selectedMailId = null;

  function renderMailInbox() {
    const ul = document.getElementById("mail-inbox");
    ul.innerHTML = MOCK_MAIL.map(
      (m) => `
      <li>
        <button type="button" class="mail-thread-btn" data-id="${m.id}">
          <strong>${escapeHtml(m.subject)}</strong>
          <span>${escapeHtml(m.from)} · ${m.date}</span>
        </button>
      </li>`
    ).join("");
    ul.querySelectorAll(".mail-thread-btn").forEach((btn) => {
      btn.addEventListener("click", () => selectMail(btn.dataset.id));
    });
  }

  function selectMail(id) {
    selectedMailId = id;
    document.querySelectorAll(".mail-thread-btn").forEach((b) => {
      b.classList.toggle("is-selected", b.dataset.id === id);
    });
    const m = MOCK_MAIL.find((x) => x.id === id);
    if (!m) return;
    document.getElementById("mail-summary").innerHTML = `<p style="margin:0;">${escapeHtml(m.summary)}</p>`;
    document.getElementById("mail-draft-to").value = m.from;
    document.getElementById("mail-draft-subject").value = `Re: ${m.subject}`;
    document.getElementById("mail-draft-body").value = "";
    const mqf = document.getElementById("mail-quote-final");
    if (mqf) mqf.value = lastQuote ? lastQuote.mid.toFixed(2) : "";
    const mqr = document.getElementById("mail-quote-range");
    if (mqr) {
      if (lastQuote)
        mqr.textContent = `Suggested range: ${formatMoney(lastQuote.low)} — ${formatMoney(lastQuote.high)} per metre`;
      else mqr.textContent = "";
    }
    syncMailPricingMini();
  }

  document.getElementById("mail-sync-pricing").addEventListener("click", () => {
    syncPricingPageToMailForm();
  });

  document.getElementById("btn-mail-calc").addEventListener("click", () => {
    runMailPricing();
  });

  function getMailFinalPrice() {
    const mq = document.getElementById("mail-quote-final");
    if (mq && mq.value !== "") return Number(mq.value) || 0;
    return Number(document.getElementById("quote-final").value) || (lastQuote && lastQuote.mid) || 0;
  }

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

  /* --- Wire refresh --- */
  function refreshAll() {
    updateDateLabel();
    refreshOverview();
    refreshCompleted();
    refreshOngoing();
    refreshReceived();
    refreshAlerts();
  }

  document.getElementById("global-start").addEventListener("change", refreshAll);
  document.getElementById("global-end").addEventListener("change", refreshAll);
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

  fillOverviewFilterSelects();
  tableColumns.completed = COMPLETED_COLS;
  tableColumns.ongoing = ONGOING_COLS;
  tableColumns.received = RECEIVED_COLS;
  bindDelegatedSort("table-awaiting", "awaiting");
  bindDelegatedSort("table-completed", "completed");
  bindDelegatedSort("table-ongoing", "ongoing");
  bindDelegatedSort("table-received", "received");

  renderMailInbox();
  initSidebarToggle();
  initDateRangePopover();
  showView("overview");
  refreshAll();
  runPricing();
})();
