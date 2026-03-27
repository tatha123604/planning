(() => {
  const buildMessage = (button) => {
    const custom = button.dataset.confirm;
    if (custom) return custom;
    const label = (button.dataset.confirmLabel || button.textContent || "this action").trim();
    if (!label) return "Are you sure you want to continue?";
    return `Are you sure you want to ${label.toLowerCase()}?`;
  };

  document.addEventListener(
    "click",
    (event) => {
      const button = event.target.closest("button");
      if (!button) return;
      if (button.dataset.noConfirm === "true") return;

      const hasExplicitConfirm =
        button.hasAttribute("data-confirm") ||
        button.hasAttribute("data-confirm-label");

      const form = button.form;
      if (form) {
        if (form.dataset.noConfirm === "true") return;
        const method = (form.getAttribute("method") || "get").toLowerCase();
        if (method === "get" && !hasExplicitConfirm) return;
        if (form.querySelector('input[type="file"]')) return;
        const inlineConfirm = (form.getAttribute("onsubmit") || "").includes("confirm(");
        if (inlineConfirm) return;
      } else if (!hasExplicitConfirm) {
        return;
      }

      if (
        button.classList.contains("sidebar-toggle-btn") ||
        button.closest(".sidebar-toggle-btn") ||
        button.closest(".brand")
      ) {
        return;
      }

      const ok = window.confirm(buildMessage(button));
      if (!ok) {
        event.preventDefault();
        event.stopImmediatePropagation();
      }
    },
    true,
  );
})();
