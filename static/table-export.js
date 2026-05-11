(() => {
  function sanitizeFilename(value) {
    return String(value || "table")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "") || "table";
  }

  function getTableTitle(table) {
    return (
      table.dataset.exportName ||
      table.closest(".card")?.querySelector("h2, summary")?.textContent ||
      "table"
    );
  }

  function getVisibleText(cell) {
    return (cell.textContent || "").replace(/\s+/g, " ").trim();
  }

  function buildExportTableMarkup(table) {
    const clone = table.cloneNode(true);
    clone.querySelectorAll("tbody tr").forEach((row) => {
      row.style.display = "";
    });
    clone.querySelectorAll("th, td").forEach((cell) => {
      cell.textContent = getVisibleText(cell);
    });
    return clone.outerHTML;
  }

  function downloadBlob(filename, blob) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function exportExcel(table) {
    const title = getTableTitle(table);
    const safeTitle = sanitizeFilename(title);
    const tableMarkup = buildExportTableMarkup(table);
    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <style>
    body { font-family: Arial, sans-serif; padding: 16px; }
    h1 { margin: 0 0 12px; font-size: 20px; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #999; padding: 8px; text-align: left; }
    th { background: #e9eef6; font-weight: 700; }
  </style>
</head>
<body>
  <h1>${title}</h1>
  ${tableMarkup}
</body>
</html>`;

    downloadBlob(
      `${safeTitle}.xls`,
      new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8;" }),
    );
  }

  function exportPdf(table) {
    const title = getTableTitle(table);
    const tableMarkup = buildExportTableMarkup(table);
    const printWindow = window.open("", "_blank", "noopener,noreferrer,width=1200,height=800");
    if (!printWindow) return;

    const html = `<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8" />
  <title>${title}</title>
  <style>
    @page { size: A4 landscape; margin: 12mm; }
    body { font-family: Arial, sans-serif; padding: 12px; color: #111; }
    h1 { margin: 0 0 12px; font-size: 18px; }
    table { border-collapse: collapse; width: 100%; font-size: 11px; }
    th, td { border: 1px solid #666; padding: 6px 8px; text-align: left; }
    th { background: #eceff4; }
  </style>
</head>
<body>
  <h1>${title}</h1>
  ${tableMarkup}
  <script>
    window.onload = () => {
      window.print();
    };
  <\/script>
</body>
</html>`;

    printWindow.document.open();
    printWindow.document.write(html);
    printWindow.document.close();
  }

  function injectExportButtons(table) {
    const wrapper = table.parentElement;
    if (!wrapper || wrapper.querySelector(".table-export-actions")) return;

    const actions = document.createElement("div");
    actions.className = "table-export-actions";
    actions.style.display = "flex";
    actions.style.justifyContent = "flex-end";
    actions.style.gap = "8px";
    actions.style.margin = "0 0 10px";
    actions.style.flexWrap = "wrap";

    const pdfButton = document.createElement("button");
    pdfButton.type = "button";
    pdfButton.className = "ghost small table-export-button table-export-button-pdf";
    pdfButton.textContent = "Download PDF";
    pdfButton.addEventListener("click", () => exportPdf(table));

    const excelButton = document.createElement("button");
    excelButton.type = "button";
    excelButton.className = "ghost small table-export-button table-export-button-excel";
    excelButton.textContent = "Download Excel";
    excelButton.addEventListener("click", () => exportExcel(table));

    actions.append(pdfButton, excelButton);
    wrapper.insertBefore(actions, table);
  }

  window.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("table[data-exportable='true']").forEach((table) => {
      injectExportButtons(table);
    });
  });
})();
