(() => {
  const STORAGE_KEY_DESKTOP = "sidebarHiddenDesktop";
  const STORAGE_KEY_MOBILE = "sidebarHiddenMobile";
  const MOBILE_MEDIA = "(max-width: 640px)";
  let btn = null;
  let mediaQuery = null;

  const isMobile = () => !!mediaQuery?.matches;

  const storageKey = () => (isMobile() ? STORAGE_KEY_MOBILE : STORAGE_KEY_DESKTOP);
  const isDashboard = () => window.location.pathname === "/";

  const update = () => {
    if (!btn) return;
    const hidden = document.body.classList.contains("nav-hidden");
    const mobile = isMobile();
    document.body.classList.toggle("nav-mobile", mobile);
    document.body.classList.toggle("brand-icon-hidden", !isDashboard());
    btn.textContent = mobile ? (hidden ? "\u2630" : "\u00d7") : hidden ? "\u203a" : "\u2039";
    btn.setAttribute("aria-expanded", hidden ? "false" : "true");
    if (mobile) {
      btn.style.left = "";
    } else {
      btn.style.left = hidden ? "12px" : "252px";
    }
  };

  const persistState = (hidden) => {
    try {
      localStorage.setItem(storageKey(), hidden ? "1" : "0");
    } catch (err) {}
  };

  const applyStoredState = () => {
    const mobile = isMobile();
    let stored = null;
    try {
      stored = localStorage.getItem(storageKey());
    } catch (err) {}
    if (stored === "1") {
      document.body.classList.add("nav-hidden");
      return;
    }
    if (stored === "0") {
      document.body.classList.remove("nav-hidden");
      return;
    }
    if (mobile) {
      document.body.classList.add("nav-hidden");
    } else {
      document.body.classList.remove("nav-hidden");
    }
  };

  document.addEventListener("DOMContentLoaded", () => {
    mediaQuery = window.matchMedia(MOBILE_MEDIA);
    btn = document.querySelector(".sidebar-toggle-btn");
    if (!btn) {
      btn = document.createElement("button");
      btn.className = "sidebar-toggle-btn";
      btn.setAttribute("aria-label", "Toggle sidebar");
      document.body.appendChild(btn);
    }

    applyStoredState();

    btn.addEventListener("click", () => {
      document.body.classList.toggle("nav-hidden");
      persistState(document.body.classList.contains("nav-hidden"));
      update();
    });

    document.querySelectorAll(".sidebar a").forEach((link) => {
      link.addEventListener("click", () => {
        if (!isMobile()) return;
        document.body.classList.add("nav-hidden");
        persistState(true);
        update();
      });
    });

    document.addEventListener("keydown", (event) => {
      if (event.key !== "Escape" || !isMobile()) return;
      document.body.classList.add("nav-hidden");
      persistState(true);
      update();
    });

    const handleViewportChange = () => {
      applyStoredState();
      update();
    };
    if (typeof mediaQuery.addEventListener === "function") {
      mediaQuery.addEventListener("change", handleViewportChange);
    } else if (typeof mediaQuery.addListener === "function") {
      mediaQuery.addListener(handleViewportChange);
    }

    update();

    const observer = new MutationObserver(update);
    observer.observe(document.body, { attributes: true, attributeFilter: ["class"] });
  });
})();
