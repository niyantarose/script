window.setTimeout(() => {
  document.querySelectorAll(".flash").forEach((node) => {
    node.style.opacity = "0";
    node.style.transition = "opacity 0.3s ease";
    window.setTimeout(() => node.remove(), 320);
  });
}, 4000);
