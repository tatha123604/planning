(() => {
  const STORAGE_KEY = "sidebarHidden";
  const btn = document.createElement("button");
  btn.className = "sidebar-toggle-btn";
  btn.setAttribute("aria-label", "Toggle sidebar");

  const update = () => {
    const hidden = document.body.classList.contains("nav-hidden");
    btn.textContent = hidden ? "\u203a" : "\u2039";
    btn.style.left = hidden ? "12px" : "252px";
  };

  const persistState = (hidden) => {
    try {
      localStorage.setItem(STORAGE_KEY, hidden ? "1" : "0");
    } catch (err) {}
  };

  btn.addEventListener("click", () => {
    document.body.classList.toggle("nav-hidden");
    persistState(document.body.classList.contains("nav-hidden"));
    update();
  });

  document.addEventListener("DOMContentLoaded", () => {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      if (stored === "1") {
        document.body.classList.add("nav-hidden");
      } else if (stored === "0") {
        document.body.classList.remove("nav-hidden");
      }
    } catch (err) {}
    document.body.appendChild(btn);
    update();

    const observer = new MutationObserver(update);
    observer.observe(document.body, { attributes: true, attributeFilter: ["class"] });
  });
})();
