(() => {
  const DEFAULT_PAGE_SIZE = 10;

  function injectControls(table) {
    if (!table.tBodies.length) return;
    const rows = Array.from(table.tBodies[0].rows || []);
    if (!rows.length) return;

    let pageSize = DEFAULT_PAGE_SIZE;
    let page = 0;

    const wrapper = document.createElement("div");
    wrapper.style.display = "flex";
    wrapper.style.justifyContent = "space-between";
    wrapper.style.alignItems = "center";
    wrapper.style.gap = "8px";
    wrapper.style.margin = "6px 0";

    const controls = document.createElement("div");
    controls.style.display = "flex";
    controls.style.gap = "8px";
    controls.style.alignItems = "center";

    const label = document.createElement("label");
    label.textContent = "Rows:";
    label.style.color = "#c9cee0";
    label.style.fontSize = "13px";

    const select = document.createElement("select");
    ["10", "20", "All"].forEach((optText) => {
      const opt = document.createElement("option");
      opt.value = optText === "All" ? "0" : optText;
      opt.textContent = optText;
      select.appendChild(opt);
    });
    select.value = String(DEFAULT_PAGE_SIZE);

    const prev = document.createElement("button");
    prev.type = "button";
    prev.textContent = "Prev";
    prev.className = "ghost small";

    const next = document.createElement("button");
    next.type = "button";
    next.textContent = "Next";
    next.className = "ghost small";

    const info = document.createElement("span");
    info.style.color = "#c9cee0";
    info.style.fontSize = "13px";

    function render() {
      const total = rows.length;
      const ps = pageSize === 0 ? total : pageSize;
      const maxPage = ps === 0 ? 0 : Math.max(0, Math.ceil(total / ps) - 1);
      if (page > maxPage) page = maxPage;

      rows.forEach((row, idx) => {
        row.classList.remove("table-paged-active");
        if (ps === 0) {
          row.style.display = "";
        } else {
          const start = page * ps;
          const end = start + ps;
          row.style.display = idx >= start && idx < end ? "" : "none";
        }
      });

      const startIdx = ps === 0 ? 1 : page * ps + 1;
      const endIdx = ps === 0 ? total : Math.min(total, (page + 1) * ps);
      info.textContent = ps === 0
        ? `Showing all ${total}`
        : `${startIdx}-${endIdx} of ${total}`;

      prev.disabled = ps === 0 || page === 0;
      next.disabled = ps === 0 || page >= maxPage;
    }

    select.addEventListener("change", () => {
      pageSize = parseInt(select.value, 10);
      page = 0;
      render();
    });

    prev.addEventListener("click", () => {
      if (page > 0) {
        page -= 1;
        render();
      }
    });

    next.addEventListener("click", () => {
      const ps = pageSize === 0 ? rows.length : pageSize;
      const maxPage = ps === 0 ? 0 : Math.max(0, Math.ceil(rows.length / ps) - 1);
      if (page < maxPage) {
        page += 1;
        render();
      }
    });

    rows.forEach((row, idx) => {
      row.addEventListener("click", () => {
        rows.forEach((r) => r.classList.remove("table-paged-active", "table-paged-active-odd", "table-paged-active-even"));
        row.classList.add("table-paged-active");
        row.classList.add(idx % 2 === 0 ? "table-paged-active-odd" : "table-paged-active-even");
      });
    });

    controls.append(label, select, prev, next);
    wrapper.append(controls, info);
    table.parentElement?.insertBefore(wrapper, table);
    render();
  }

  window.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("table.paged").forEach(injectControls);
  });
})();
