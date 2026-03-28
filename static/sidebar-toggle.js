(() => {
  const NAV_STORAGE_KEY = "sdah-nav-hidden";
  const btn = document.createElement("button");
  btn.className = "sidebar-toggle-btn";
  btn.setAttribute("aria-label", "Toggle sidebar");

  const readStoredHidden = () => {
    try {
      return window.localStorage.getItem(NAV_STORAGE_KEY) === "1";
    } catch (error) {
      return document.body.classList.contains("nav-hidden");
    }
  };

  const persistHidden = (hidden) => {
    try {
      window.localStorage.setItem(NAV_STORAGE_KEY, hidden ? "1" : "0");
    } catch (error) {
      return;
    }
  };

  const applyHidden = (hidden) => {
    document.body.classList.toggle("nav-hidden", hidden);
    return hidden;
  };

  const update = () => {
    const hidden = document.body.classList.contains("nav-hidden");
    btn.textContent = hidden ? "\u203a" : "\u2039";
    btn.style.left = hidden ? "12px" : "252px";
  };

  btn.addEventListener("click", () => {
    const hidden = applyHidden(!document.body.classList.contains("nav-hidden"));
    persistHidden(hidden);
    update();
  });

  applyHidden(readStoredHidden());

  document.addEventListener("DOMContentLoaded", () => {
    applyHidden(readStoredHidden());
    document.body.appendChild(btn);
    update();

    const observer = new MutationObserver(update);
    observer.observe(document.body, { attributes: true, attributeFilter: ["class"] });
  });
})();
