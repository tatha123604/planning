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

      const form = button.form;
      if (form) {
        if (form.dataset.noConfirm === "true") return;
        if (form.querySelector('input[type="file"]')) return;
        const inlineConfirm = (form.getAttribute("onsubmit") || "").includes("confirm(");
        if (inlineConfirm) return;
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
