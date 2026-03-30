(() => {
  const DEFAULT_PAGE_SIZE = 10;
  const isFilterHidden = (row) => row?.dataset?.filterHidden === "true";

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

  function createPagerBar({ showRowsSelector = false } = {}) {
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

    return { wrapper, info, prev, next, select };
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
      Array.from(row.cells || []).some((cell) =>
        String(cell.textContent || "").trim().toUpperCase() === "TOTAL"
      );
    const stickyRows = rows.filter(isTotalRow);
    const pageableRows = rows.filter((row) => !isTotalRow(row));
    if (!pageableRows.length) return;

    let pageSize = DEFAULT_PAGE_SIZE;
    let page = 0;

    const topPager = createPagerBar({ showRowsSelector: true });
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
    window.addEventListener("resize", syncPagerWidths);
    table.addEventListener("table-pager:refresh", (event) => {
      if (event?.detail?.resetPage) {
        page = 0;
      }
      render();
    });
    table.dataset.pagerReady = "true";
    render();
  }

  window.__tablePagerInject = injectControls;

  window.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("table.paged").forEach(injectControls);
  });
})();
