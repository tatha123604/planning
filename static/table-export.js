(() => {
  const ACTION_HEADER_RE = /^actions$/i;

  const normalizeText = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

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

  const parseDownloadFilename = (response, fallbackTitle, extension) => {
    const header = response.headers.get("content-disposition") || "";
    const utfMatch = header.match(/filename\*=UTF-8''([^;]+)/i);
    if (utfMatch?.[1]) {
      return decodeURIComponent(utfMatch[1]);
    }
    const plainMatch = header.match(/filename="?([^"]+)"?/i);
    if (plainMatch?.[1]) {
      return plainMatch[1];
    }
    return `${sanitizeFilename(fallbackTitle)}.${extension}`;
  };

  const downloadSnapshot = async (
    snapshot,
    endpoint,
    extension,
    button,
    setStatus,
    pendingMessage,
    readyMessage,
    failureMessage
  ) => {
    try {
      setStatus(pendingMessage);
      button.disabled = true;
      const response = await fetch(endpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(snapshot),
      });
      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(errorText || failureMessage);
      }
      const blob = await response.blob();
      const downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = downloadUrl;
      link.download = parseDownloadFilename(response, snapshot.title, extension);
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
      setStatus(`${readyMessage} (${snapshot.rows.length} rows).`);
      window.setTimeout(() => setStatus(""), 2400);
    } catch (error) {
      console.error(error);
      setStatus(failureMessage, true);
    } finally {
      button.disabled = false;
    }
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

    pdfButton.addEventListener("click", async () => {
      const snapshot = buildSnapshot(table, tableIndex);
      const confirmed = window.confirm(
        `Generate PDF for "${snapshot.title}"?\nRows: ${snapshot.rows.length}`
      );
      if (!confirmed) {
        setStatus("PDF export cancelled.");
        window.setTimeout(() => setStatus(""), 1800);
        return;
      }
      await downloadSnapshot(
        snapshot,
        "/exports/table.pdf",
        "pdf",
        pdfButton,
        setStatus,
        "Preparing PDF...",
        "PDF ready",
        "PDF export failed."
      );
    });

    excelButton.addEventListener("click", async () => {
      const snapshot = buildSnapshot(table, tableIndex);
      const confirmed = window.confirm(
        `Download Excel for "${snapshot.title}"?\nRows: ${snapshot.rows.length}`
      );
      if (!confirmed) {
        setStatus("Excel export cancelled.");
        window.setTimeout(() => setStatus(""), 1800);
        return;
      }
      await downloadSnapshot(
        snapshot,
        "/exports/table.xlsx",
        "xlsx",
        excelButton,
        setStatus,
        "Preparing Excel...",
        "Excel ready",
        "Excel export failed."
      );
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
