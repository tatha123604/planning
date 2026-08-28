(() => {
  const DEFAULT_PAGE_SIZE = 10;
  const isFilterHidden = (row) => (
    row?.dataset?.filterHidden === "true"
    || row?.dataset?.smartDistanceHidden === "true"
  );

  function normalizeText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function parseSortableNumber(text) {
    const normalized = normalizeText(text).replace(/,/g, "");
    if (!normalized) return null;
    const num = Number(normalized);
    return Number.isFinite(num) ? num : null;
  }

  function parseSortableDateTime(text) {
    // Matches "dd-mm-yyyy hh:mm" or "dd-mm-yyyy hh:mm:ss" with optional timezone suffix (IST/UTC).
    const normalized = normalizeText(text);
    const match = normalized.match(/(\d{2})-(\d{2})-(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?/);
    if (!match) return null;
    const [, dd, mm, yyyy, hh, mi, ss] = match;
    // Use a sortable yyyymmddhhmmss integer (timezone label doesn't matter if consistent).
    return Number(`${yyyy}${mm}${dd}${hh}${mi}${(ss || "00").padStart(2, "0")}`);
  }

  function parseSortableDuration(text) {
    // Supports:
    // - "HH hh MM mm SS ss"
    // - "DD DD HH hh MM mm SS ss"
    // - "MM MM DD DD HH hh MM mm SS ss"
    // - "YY YY MM MM DD DD HH hh MM mm SS ss"
    const normalized = normalizeText(text).toUpperCase();
    if (!normalized) return null;
    const grab = (unit) => {
      const m = normalized.match(new RegExp(`(\\d+)\\s*${unit}\\b`));
      return m ? parseInt(m[1], 10) : 0;
    };

    const years = grab("YY");
    const months = grab("MM");
    const days = grab("DD");
    const hours = grab("HH");
    const secs = grab("SS");

    // Disambiguate months vs minutes by looking for the time-part "hh ... mm ... ss".
    const timeMinMatch = normalized.match(/(\d+)\s*HH\b.*?(\d+)\s*MM\b.*?(\d+)\s*SS\b/);
    const minutes = timeMinMatch ? parseInt(timeMinMatch[2], 10) : 0;

    // Approximate conversion: 1Y=365d, 1M=30d.
    const totalDays = years * 365 + months * 30 + days;
    return totalDays * 86400 + hours * 3600 + minutes * 60 + secs;
  }

  function styleGroup(group, justifyContent = "flex-start") {
    group.style.display = "flex";
    group.style.alignItems = "center";
    group.style.gap = "8px";
    group.style.flexWrap = "wrap";
    group.style.justifyContent = justifyContent;
  }

  function styleInfo(info) {
    info.style.color = "#c9cee0";
    info.style.fontSize = "13px";
  }

  function createPagerBar({ showRowsSelector = false, showSearch = false } = {}) {
    const wrapper = document.createElement("div");
    wrapper.className = "table-pager";
    wrapper.style.display = "flex";
    wrapper.style.justifyContent = "space-between";
    wrapper.style.alignItems = "center";
    wrapper.style.gap = "8px";
    wrapper.style.margin = "6px 0";
    wrapper.style.flexWrap = "wrap";

    const left = document.createElement("div");
    left.className = "table-pager-left";
    styleGroup(left);

    const right = document.createElement("div");
    right.className = "table-pager-right";
    styleGroup(right, "flex-end");
    right.style.marginLeft = "auto";

    let select = null;
    let search = null;

    if (showRowsSelector) {
      const label = document.createElement("label");
      label.className = "table-pager-label";
      label.textContent = "Rows:";
      styleInfo(label);

      select = document.createElement("select");
      select.className = "table-pager-select";
      ["10", "20", "All"].forEach((optText) => {
        const opt = document.createElement("option");
        opt.value = optText === "All" ? "0" : optText;
        opt.textContent = optText;
        select.appendChild(opt);
      });
      select.value = String(DEFAULT_PAGE_SIZE);
      left.append(label, select);
    }

    if (showSearch) {
      search = document.createElement("input");
      search.type = "search";
      search.className = "table-pager-search";
      search.placeholder = "Search table...";
      search.setAttribute("aria-label", "Search table rows");
      left.append(search);
    }

    const info = document.createElement("span");
    info.className = "table-pager-info";
    styleInfo(info);

    const prev = document.createElement("button");
    prev.type = "button";
    prev.textContent = "Prev";
    prev.className = "ghost small table-pager-button table-pager-button-prev";

    const next = document.createElement("button");
    next.type = "button";
    next.textContent = "Next";
    next.className = "ghost small table-pager-button table-pager-button-next";

    left.append(info);
    right.append(prev, next);
    wrapper.append(left, right);

    return { wrapper, info, prev, next, select, search };
  }

  function injectControls(table) {
    if (table.dataset.pagerReady === "true") return;
    if (!table.tBodies.length) return;
    const rows = Array.from(table.tBodies[0].rows || []);
    if (!rows.length) return;
    if (table.dataset.pager === "server") {
      rows.forEach((row, idx) => {
        row.addEventListener("click", () => {
          rows.forEach((r) =>
            r.classList.remove("table-paged-active", "table-paged-active-odd", "table-paged-active-even")
          );
          row.classList.add("table-paged-active");
          row.classList.add(idx % 2 === 0 ? "table-paged-active-odd" : "table-paged-active-even");
        });
      });
      table.dataset.pagerReady = "true";
      return;
    }
    const scrollContainer = table.closest(".table-wrap");
    const pagerAnchor = scrollContainer || table;
    const pagerHost = pagerAnchor.parentElement;
    if (!pagerHost) return;
    const isTotalRow = (row) =>
      Array.from(row.cells || []).some((cell) => {
        const text = String(cell.textContent || "").trim().toUpperCase();
        return text === "TOTAL" || text.startsWith("TOTAL ");
      });
    const stickyRows = rows.filter(isTotalRow);
    const pageableRows = rows.filter((row) => !isTotalRow(row));
    if (!pageableRows.length) return;

    let pageSize = table.id === "cli-distribution-table" ? 0 : DEFAULT_PAGE_SIZE;
    let page = 0;
    let sortColumn = null;
    let sortDirection = 1; // 1 = asc, -1 = desc

    const topPager = createPagerBar({
      showRowsSelector: true,
      showSearch: table.dataset.searchable === "true" || window.location.pathname === "/ssts-report",
    });
    if (table.id === "cli-distribution-table" && topPager.select) {
      topPager.select.value = "0";
    }
    const bottomPager = createPagerBar();
    const infos = [topPager.info, bottomPager.info];
    const prevButtons = [topPager.prev, bottomPager.prev];
    const nextButtons = [topPager.next, bottomPager.next];
    const syncPagerWidths = () => {
      const hostWidth = pagerAnchor.clientWidth || pagerHost.clientWidth || table.clientWidth || 0;
      [topPager.wrapper, bottomPager.wrapper].forEach((wrapper) => {
        wrapper.style.width = hostWidth > 0 ? `${hostWidth}px` : "";
      });
    };

    const tbody = table.tBodies[0];
    const originalIndex = new Map(pageableRows.map((row, idx) => [row, idx]));
    const updateSortIndicators = () => {
      const headCells = Array.from(table.tHead?.rows?.[0]?.cells || []);
      headCells.forEach((cell, idx) => {
        const indicator = cell.querySelector(".table-sort-indicator");
        if (!indicator) return;
        if (idx !== sortColumn) {
          indicator.textContent = "";
          indicator.setAttribute("data-dir", "");
          return;
        }
        indicator.textContent = sortDirection === 1 ? "▲" : "▼";
        indicator.setAttribute("data-dir", sortDirection === 1 ? "asc" : "desc");
      });
    };

    const coerceSortKey = (headerText, cellText) => {
      const header = normalizeText(headerText).toLowerCase();
      const value = normalizeText(cellText);
      if (!value) return { kind: "empty", value: "" };

      if (header.includes("update") || header.includes("snapshot") || header.includes("time") || header.includes("date")) {
        const dt = parseSortableDateTime(value);
        if (dt !== null) return { kind: "number", value: dt };
      }

      if (header.includes("offline") || header.includes("duration") || header.includes("hours") || header.includes("status")) {
        const dur = parseSortableDuration(value);
        if (dur !== null) return { kind: "number", value: dur };
      }

      const num = parseSortableNumber(value);
      if (num !== null) return { kind: "number", value: num };

      return { kind: "text", value: value.toLowerCase() };
    };

    const sortRows = () => {
      if (sortColumn === null) return;
      const headCells = Array.from(table.tHead?.rows?.[0]?.cells || []);
      const headerText = headCells[sortColumn]?.textContent || "";
      const getCellText = (row) => {
        const cell = row.cells?.[sortColumn];
        if (!cell) return "";
        const explicitSortValue = cell.dataset?.sortValue;
        if (explicitSortValue) return normalizeText(explicitSortValue);
        // Prefer visible text; ignore inputs/buttons for remark editor etc.
        return normalizeText(cell.innerText || cell.textContent || "");
      };

      pageableRows.sort((a, b) => {
        const ak = coerceSortKey(headerText, getCellText(a));
        const bk = coerceSortKey(headerText, getCellText(b));

        if (ak.kind === "empty" && bk.kind !== "empty") return 1;
        if (bk.kind === "empty" && ak.kind !== "empty") return -1;

        if (ak.kind === "number" && bk.kind === "number") {
          if (ak.value === bk.value) return (originalIndex.get(a) ?? 0) - (originalIndex.get(b) ?? 0);
          return sortDirection * (ak.value - bk.value);
        }

        if (ak.value === bk.value) return (originalIndex.get(a) ?? 0) - (originalIndex.get(b) ?? 0);
        return sortDirection * String(ak.value).localeCompare(String(bk.value));
      });

      // Re-append rows in sorted order so paging/filtering uses the new ordering.
      pageableRows.forEach((row) => tbody.appendChild(row));
      stickyRows.forEach((row) => tbody.appendChild(row));
      updateSortIndicators();
    };

    const enableSorting = () => {
      const headRow = table.tHead?.rows?.[0];
      if (!headRow) return;
      const headCells = Array.from(headRow.cells || []);
      if (!headCells.length) return;

      headCells.forEach((cell, idx) => {
        // Avoid double-injecting.
        if (cell.querySelector(".table-sort-indicator")) return;
        cell.classList.add("table-sortable");
        cell.tabIndex = 0;
        const indicator = document.createElement("span");
        indicator.className = "table-sort-indicator";
        indicator.setAttribute("aria-hidden", "true");
        cell.appendChild(indicator);

        const activate = () => {
          if (sortColumn === idx) {
            sortDirection = sortDirection === 1 ? -1 : 1;
          } else {
            sortColumn = idx;
            sortDirection = 1;
          }
          page = 0;
          sortRows();
          render();
        };

        cell.addEventListener("click", activate);
        cell.addEventListener("keydown", (event) => {
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            activate();
          }
        });
      });
    };

    function render() {
      syncPagerWidths();
      const visiblePageableRows = pageableRows.filter((row) => !isFilterHidden(row));
      const total = visiblePageableRows.length;
      const ps = pageSize === 0 ? total : pageSize;
      const maxPage = ps === 0 ? 0 : Math.max(0, Math.ceil(total / ps) - 1);
      if (page > maxPage) page = maxPage;

      pageableRows.forEach((row) => {
        row.classList.remove("table-paged-active");
      });

      visiblePageableRows.forEach((row, idx) => {
        if (ps === 0) {
          row.style.display = "";
          return;
        }
        const start = page * ps;
        const end = start + ps;
        row.style.display = idx >= start && idx < end ? "" : "none";
      });

      pageableRows
        .filter((row) => isFilterHidden(row))
        .forEach((row) => {
          row.style.display = "none";
        });

      stickyRows.forEach((row) => {
        row.style.display = total > 0 ? "" : "none";
      });

      if (total === 0) {
        infos.forEach((info) => {
          info.textContent = "0 results";
        });
      } else {
        const startIdx = ps === 0 ? 1 : page * ps + 1;
        const endIdx = ps === 0 ? total : Math.min(total, (page + 1) * ps);
        const message = ps === 0
          ? `Showing all ${total}`
          : `${startIdx}-${endIdx} of ${total}`;
        infos.forEach((info) => {
          info.textContent = message;
        });
      }

      prevButtons.forEach((button) => {
        button.disabled = ps === 0 || page === 0;
      });
      nextButtons.forEach((button) => {
        button.disabled = ps === 0 || page >= maxPage;
      });
    }

    topPager.select?.addEventListener("change", () => {
      pageSize = parseInt(topPager.select.value, 10);
      page = 0;
      render();
    });

    topPager.search?.addEventListener("input", () => {
      const query = String(topPager.search.value || "").trim().toLowerCase();
      pageableRows.forEach((row) => {
        const text = String(row.textContent || "").replace(/\s+/g, " ").trim().toLowerCase();
        row.dataset.filterHidden = query && !text.includes(query) ? "true" : "false";
      });
      page = 0;
      render();
    });

    prevButtons.forEach((button) => {
      button.addEventListener("click", () => {
        if (page > 0) {
          page -= 1;
          render();
        }
      });
    });

    nextButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const total = pageableRows.filter((row) => !isFilterHidden(row)).length;
        const ps = pageSize === 0 ? total : pageSize;
        const maxPage = ps === 0 ? 0 : Math.max(0, Math.ceil(total / ps) - 1);
        if (page < maxPage) {
          page += 1;
          render();
        }
      });
    });

    rows.forEach((row, idx) => {
      row.addEventListener("click", () => {
        rows.forEach((r) => r.classList.remove("table-paged-active", "table-paged-active-odd", "table-paged-active-even"));
        row.classList.add("table-paged-active");
        row.classList.add(idx % 2 === 0 ? "table-paged-active-odd" : "table-paged-active-even");
      });
    });

    pagerHost.insertBefore(topPager.wrapper, pagerAnchor);
    pagerAnchor.insertAdjacentElement("afterend", bottomPager.wrapper);
    const exportBar = pagerHost.querySelector(".table-export-bar");
    if (exportBar) {
      topPager.wrapper.insertAdjacentElement("afterend", exportBar);
    }
    window.addEventListener("resize", syncPagerWidths);
    table.addEventListener("table-pager:refresh", (event) => {
      if (event?.detail?.resetPage) {
        page = 0;
      }
      render();
    });
    table.dataset.pagerReady = "true";
    enableSorting();
    sortRows();
    render();
  }

  window.__tablePagerInject = injectControls;

  const injectTopScrollbar = (wrap) => {
    if (!wrap || wrap.dataset.topScrollReady === "true") return;
    const table = wrap.querySelector("table");
    if (!table) return;

    const topBar = document.createElement("div");
    topBar.className = "table-scrollbar-top";
    const spacer = document.createElement("div");
    spacer.className = "table-scrollbar-spacer";
    topBar.appendChild(spacer);

    const syncSizes = () => {
      spacer.style.width = `${table.scrollWidth}px`;
    };

    const syncing = { top: false, body: false };
    topBar.addEventListener("scroll", () => {
      if (syncing.top) return;
      syncing.body = true;
      wrap.scrollLeft = topBar.scrollLeft;
      syncing.body = false;
    });
    wrap.addEventListener("scroll", () => {
      if (syncing.body) return;
      syncing.top = true;
      topBar.scrollLeft = wrap.scrollLeft;
      syncing.top = false;
    });

    wrap.parentElement?.insertBefore(topBar, wrap);
    window.addEventListener("resize", syncSizes);
    syncSizes();
    wrap.dataset.topScrollReady = "true";
  };

  window.__tableTopScrollbarInject = injectTopScrollbar;

  window.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("table.paged").forEach(injectControls);
    document.querySelectorAll(".table-wrap").forEach(injectTopScrollbar);
  });
})();
