(() => {
  const ACTION_HEADER_RE = /^actions$/i;

  const normalizeText = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

  const escapeHtml = (value) =>
    String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");

  const sanitizeFilename = (value, fallback = "table_export") => {
    const cleaned = normalizeText(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    return cleaned || fallback;
  };

  const extractExpandedTexts = (cells) => {
    const values = [];
    Array.from(cells || []).forEach((cell) => {
      const text = normalizeText(cell.textContent);
      const span = Math.max(1, parseInt(cell.getAttribute("colspan") || "1", 10) || 1);
      values.push(text);
      for (let index = 1; index < span; index += 1) {
        values.push("");
      }
    });
    return values;
  };

  const getTableTitle = (table, tableIndex) => {
    const details = table.closest("details");
    if (details) {
      const summary = Array.from(details.children).find((child) => child.tagName === "SUMMARY");
      if (summary) {
        return normalizeText(summary.textContent);
      }
    }

    const card = table.closest(".card");
    if (card) {
      const heading = Array.from(card.children).find((child) => /^H[12]$/.test(child.tagName));
      if (heading) {
        return normalizeText(heading.textContent);
      }
    }

    const pageHeading = document.querySelector("main.content h1");
    const baseTitle = pageHeading ? normalizeText(pageHeading.textContent) : "Table";
    return `${baseTitle} Table ${tableIndex + 1}`;
  };

  const buildSnapshot = (table, tableIndex) => {
    const headerRows = Array.from(table.tHead?.rows || []);
    const headerSource = headerRows.length ? headerRows[headerRows.length - 1] : null;
    const rawHeaders = headerSource ? extractExpandedTexts(headerSource.cells) : [];
    const bodyRows = Array.from(table.tBodies || []).flatMap((tbody) => Array.from(tbody.rows || []));
    const rawRows = bodyRows
      .filter((row) => row?.dataset?.filterHidden !== "true")
      .map((row) => extractExpandedTexts(row.cells));

    const columnCount = Math.max(
      rawHeaders.length,
      rawRows.reduce((max, row) => Math.max(max, row.length), 0),
      1
    );
    const headers = Array.from({ length: columnCount }, (_, index) =>
      normalizeText(rawHeaders[index]) || `Column ${index + 1}`
    );
    const ignoredIndexes = new Set(
      headers
        .map((header, index) => (ACTION_HEADER_RE.test(header) ? index : -1))
        .filter((index) => index >= 0)
    );
    const keepIndexes = headers.map((_, index) => index).filter((index) => !ignoredIndexes.has(index));
    const filteredHeaders = keepIndexes.map((index) => headers[index]);
    const rows = rawRows.map((row) => {
      const padded = row.slice(0, columnCount);
      while (padded.length < columnCount) {
        padded.push("");
      }
      return keepIndexes.map((index) => normalizeText(padded[index]));
    });

    return {
      title: getTableTitle(table, tableIndex),
      headers: filteredHeaders,
      rows,
    };
  };

  const parseDownloadFilename = (response, fallbackTitle) => {
    const header = response.headers.get("content-disposition") || "";
    const utfMatch = header.match(/filename\*=UTF-8''([^;]+)/i);
    if (utfMatch?.[1]) {
      return decodeURIComponent(utfMatch[1]);
    }
    const plainMatch = header.match(/filename="?([^"]+)"?/i);
    if (plainMatch?.[1]) {
      return plainMatch[1];
    }
    return `${sanitizeFilename(fallbackTitle)}.xlsx`;
  };

  const buildPrintDocument = (snapshot) => {
    const headerHtml = snapshot.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("");
    const rowHtml = snapshot.rows
      .map((row) => `<tr>${row.map((cell) => `<td>${escapeHtml(cell)}</td>`).join("")}</tr>`)
      .join("");

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <title>${escapeHtml(snapshot.title)} PDF</title>
  <style>
    @page { size: landscape; margin: 12mm; }
    body {
      font-family: "Segoe UI", Arial, sans-serif;
      margin: 0;
      color: #122235;
      background: #ffffff;
    }
    .print-shell {
      padding: 18px 20px;
    }
    h1 {
      margin: 0 0 8px;
      font-size: 24px;
      color: #173a5c;
    }
    .meta {
      margin: 0 0 16px;
      font-size: 12px;
      color: #4f6d87;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      table-layout: auto;
      font-size: 11px;
    }
    thead th {
      background: #173a5c;
      color: #f3fbff;
      padding: 8px 10px;
      border: 1px solid #d3dfeb;
      text-align: left;
    }
    tbody td {
      padding: 7px 10px;
      border: 1px solid #d3dfeb;
      vertical-align: top;
      word-break: break-word;
    }
    tbody tr:nth-child(even) td {
      background: #f7fbff;
    }
  </style>
</head>
<body>
  <div class="print-shell">
    <h1>${escapeHtml(snapshot.title)}</h1>
    <p class="meta">Rows: ${snapshot.rows.length} | Generated: ${escapeHtml(new Date().toLocaleString())}</p>
    <table>
      <thead><tr>${headerHtml}</tr></thead>
      <tbody>${rowHtml}</tbody>
    </table>
  </div>
  <script>
    window.addEventListener("load", () => {
      window.setTimeout(() => window.print(), 150);
    });
    window.onafterprint = () => window.close();
  </script>
</body>
</html>`;
  };

  const injectToolbar = (table, tableIndex) => {
    if (table.dataset.exportReady === "true" || !table.tBodies.length) {
      return;
    }

    const primaryAnchor = table.closest(".table-wrap") || table;
    const host = primaryAnchor.parentElement;
    if (!host) {
      return;
    }

    const toolbar = document.createElement("div");
    toolbar.className = "table-export-bar";

    const label = document.createElement("span");
    label.className = "table-export-label";
    label.textContent = "Table Export";

    const status = document.createElement("span");
    status.className = "table-export-status";

    const actions = document.createElement("div");
    actions.className = "table-export-actions";

    const pdfButton = document.createElement("button");
    pdfButton.type = "button";
    pdfButton.className = "ghost small table-export-button table-export-button-pdf";
    pdfButton.textContent = "Generate PDF";

    const excelButton = document.createElement("button");
    excelButton.type = "button";
    excelButton.className = "ghost small table-export-button table-export-button-excel";
    excelButton.textContent = "Download Excel";

    const setStatus = (message, isError = false) => {
      status.textContent = message;
      status.classList.toggle("is-error", isError);
    };

    pdfButton.addEventListener("click", () => {
      const snapshot = buildSnapshot(table, tableIndex);
      const printWindow = window.open("", "_blank", "noopener,noreferrer");
      if (!printWindow) {
        setStatus("PDF window blocked.", true);
        return;
      }
      printWindow.document.open();
      printWindow.document.write(buildPrintDocument(snapshot));
      printWindow.document.close();
      setStatus("PDF ready for save.");
      window.setTimeout(() => setStatus(""), 2400);
    });

    excelButton.addEventListener("click", async () => {
      const snapshot = buildSnapshot(table, tableIndex);
      try {
        setStatus("Preparing Excel...");
        excelButton.disabled = true;
        const response = await fetch("/exports/table.xlsx", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(snapshot),
        });
        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(errorText || "Excel export failed.");
        }
        const blob = await response.blob();
        const downloadUrl = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = downloadUrl;
        link.download = parseDownloadFilename(response, snapshot.title);
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
        setStatus(`Excel ready (${snapshot.rows.length} rows).`);
        window.setTimeout(() => setStatus(""), 2400);
      } catch (error) {
        console.error(error);
        setStatus("Excel export failed.", true);
      } finally {
        excelButton.disabled = false;
      }
    });

    actions.append(pdfButton, excelButton);
    toolbar.append(label, status, actions);

    const insertionPoint =
      primaryAnchor.previousElementSibling?.classList.contains("table-pager")
        ? primaryAnchor.previousElementSibling
        : primaryAnchor;
    host.insertBefore(toolbar, insertionPoint);
    table.dataset.exportReady = "true";
  };

  window.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("table").forEach((table, tableIndex) => injectToolbar(table, tableIndex));
  });
})();
