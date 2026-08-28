(() => {
  const ACTION_HEADER_RE = /^actions$/i;
  const EXPORT_IGNORED_HEADER_RE = /^graphical representation$/i;
  const SKIP_PDF_HEADER_RE = /^(select|chart)$/i;

  const normalizeText = (value) => String(value ?? "").replace(/\s+/g, " ").trim();

  const sanitizeFilename = (value, fallback = "table_export") => {
    const cleaned = normalizeText(value)
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    return cleaned || fallback;
  };

  const readBlobSignature = async (blob, size = 8) => {
    const bytes = new Uint8Array(await blob.slice(0, size).arrayBuffer());
    return Array.from(bytes);
  };

  const matchesSignature = (bytes, signature) =>
    signature.every((value, index) => bytes[index] === value);

  const getExpectedContentType = (extension) =>
    extension === "pdf"
      ? "application/pdf"
      : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  const validateDownloadBlob = async (blob, extension) => {
    const bytes = await readBlobSignature(blob, extension === "pdf" ? 5 : 4);
    if (extension === "pdf") {
      return matchesSignature(bytes, [0x25, 0x50, 0x44, 0x46, 0x2d]);
    }
    return matchesSignature(bytes, [0x50, 0x4b, 0x03, 0x04]);
  };

  const extractErrorMessage = async (response, fallbackMessage) => {
    const contentType = response.headers.get("content-type") || "";
    try {
      if (contentType.includes("application/json")) {
        const payload = await response.json();
        return normalizeText(payload?.detail || payload?.message || "") || fallbackMessage;
      }
      const text = await response.text();
      return normalizeText(text) || fallbackMessage;
    } catch (error) {
      console.error(error);
      return fallbackMessage;
    }
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
    const explicitTitle = normalizeText(table?.dataset?.exportTitle);
    if (explicitTitle) {
      return explicitTitle;
    }
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

  const buildSnapshot = (table, tableIndex, options = {}) => {
    const { excludeMarkedPdfRows = false, includeSelectedPdfRows = false } = options;
    const scopedReportDateInput = table.closest("section, .card, details")?.querySelector('[data-export-report-date-source="true"]');
    const reportDateInput = document.querySelector('input[name="report_date"]');
    const sourceReportDateInput = document.querySelector("#source-report-date");
    const reportDateValue = normalizeText(
      table?.dataset?.exportReportDate
      || scopedReportDateInput?.value
      || sourceReportDateInput?.value
      || reportDateInput?.value
      || ""
    );
    const headerRows = Array.from(table.tHead?.rows || []);
    const headerSource = headerRows.length ? headerRows[headerRows.length - 1] : null;
    const rawHeaders = headerSource ? extractExpandedTexts(headerSource.cells) : [];
    const bodyRows = Array.from(table.tBodies || []).flatMap((tbody) => Array.from(tbody.rows || []));
    const selectedPdfRows = includeSelectedPdfRows
      ? bodyRows.filter((row) => row.querySelector('input[data-pdf-select-row]:checked'))
      : [];
    const candidateRows = selectedPdfRows.length ? selectedPdfRows : bodyRows;
    const includedRows = candidateRows
      .filter((row) => row?.dataset?.filterHidden !== "true")
      .filter((row) => row?.dataset?.smartDistanceHidden !== "true")
      .filter((row) => (
        !excludeMarkedPdfRows || !row.querySelector('input[data-skip-pdf-row]:checked')
      ));
    const rawRows = includedRows
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
        .map((header, index) => (
          ACTION_HEADER_RE.test(header)
          || EXPORT_IGNORED_HEADER_RE.test(header)
          || SKIP_PDF_HEADER_RE.test(header)
            ? index
            : -1
        ))
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
    const cellClasses = includedRows
      .map((row) => {
        const cells = Array.from(row.cells || []);
        const expanded = cells.flatMap((cell) => {
          const className = normalizeText(cell.className);
          const span = Math.max(1, parseInt(cell.getAttribute("colspan") || "1", 10) || 1);
          const values = [className];
          for (let index = 1; index < span; index += 1) {
            values.push("");
          }
          return values;
        });
        while (expanded.length < columnCount) {
          expanded.push("");
        }
        return keepIndexes.map((index) => expanded[index] || "");
      });

    return {
      title: getTableTitle(table, tableIndex),
      headers: filteredHeaders,
      rows,
      cell_classes: cellClasses,
      report_date: reportDateValue,
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
        credentials: "same-origin",
        headers: {
          "Content-Type": "application/json",
          Accept: `${getExpectedContentType(extension)}, application/json, text/plain`,
          "X-Requested-With": "fetch",
        },
        body: JSON.stringify(snapshot),
      });
      if (!response.ok) {
        const errorText = await extractErrorMessage(response, failureMessage);
        throw new Error(errorText || failureMessage);
      }
      const blob = await response.blob();
      const expectedContentType = getExpectedContentType(extension);
      const contentType = response.headers.get("content-type") || blob.type || "";
      const validBlob = await validateDownloadBlob(blob, extension);
      if (!validBlob || !contentType.includes(expectedContentType)) {
        throw new Error(
          "The export response was not a valid file. Please sign in again and retry."
        );
      }
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
      return true;
    } catch (error) {
      console.error(error);
      setStatus(failureMessage, true);
      return false;
    } finally {
      button.disabled = false;
    }
  };

  const submitPdfFallbackForm = (snapshot, setStatus) => {
    const frameName = `table-export-pdf-frame-${Date.now()}`;
    const iframe = document.createElement("iframe");
    iframe.name = frameName;
    iframe.style.display = "none";
    document.body.appendChild(iframe);

    const form = document.createElement("form");
    form.method = "POST";
    form.action = "/exports/table.pdf/form";
    form.target = frameName;
    form.style.display = "none";

    const payloadInput = document.createElement("input");
    payloadInput.type = "hidden";
    payloadInput.name = "payload";
    payloadInput.value = JSON.stringify(snapshot);
    form.appendChild(payloadInput);
    document.body.appendChild(form);

    setStatus(`Trying fallback PDF export (${snapshot.rows.length} rows)...`);
    form.submit();
    window.setTimeout(() => {
      form.remove();
      iframe.remove();
      setStatus("Fallback PDF export started.");
      window.setTimeout(() => setStatus(""), 2400);
    }, 1500);
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
      const snapshot = buildSnapshot(table, tableIndex, {
        excludeMarkedPdfRows: true,
        includeSelectedPdfRows: true,
      });
      const confirmed = window.confirm(
        `Generate PDF for "${snapshot.title}"?\nRows: ${snapshot.rows.length}`
      );
      if (!confirmed) {
        setStatus("PDF export cancelled.");
        window.setTimeout(() => setStatus(""), 1800);
        return;
      }
      const ok = await downloadSnapshot(
        snapshot,
        "/exports/table.pdf",
        "pdf",
        pdfButton,
        setStatus,
        "Preparing PDF...",
        "PDF ready",
        "PDF export failed."
      );
      if (!ok) {
        submitPdfFallbackForm(snapshot, setStatus);
      }
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

    const pagerAbove = primaryAnchor.previousElementSibling;
    if (pagerAbove && pagerAbove.classList.contains("table-pager")) {
      pagerAbove.insertAdjacentElement("afterend", toolbar);
    } else {
      host.insertBefore(toolbar, primaryAnchor);
    }
    table.dataset.exportReady = "true";
  };

  window.__tableExportInject = injectToolbar;

  window.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("table").forEach((table, tableIndex) => injectToolbar(table, tableIndex));
  });
})();
