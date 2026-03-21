(() => {
  const btn = document.createElement("button");
  btn.className = "sidebar-toggle-btn";
  btn.setAttribute("aria-label", "Toggle sidebar");

  const update = () => {
    const hidden = document.body.classList.contains("nav-hidden");
    btn.textContent = hidden ? "›" : "‹";
    btn.style.left = hidden ? "12px" : "252px";
  };

  btn.addEventListener("click", () => {
    document.body.classList.toggle("nav-hidden");
    update();
  });

  document.addEventListener("DOMContentLoaded", () => {
    document.body.appendChild(btn);
    update();
  });
})();
