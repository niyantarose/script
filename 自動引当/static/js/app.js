const UI_SCALE_KEY = "inventory-tool-ui-scale";
const VALID_UI_SCALES = ["small", "medium", "large"];

function setUiScale(scale) {
  const nextScale = VALID_UI_SCALES.includes(scale) ? scale : "small";
  document.body.dataset.uiScale = nextScale;
  try {
    window.localStorage.setItem(UI_SCALE_KEY, nextScale);
  } catch (error) {
    console.warn("UI scale could not be saved.", error);
  }
  document.querySelectorAll(".scale-btn").forEach((button) => {
    button.classList.toggle("active", button.dataset.uiScale === nextScale);
  });
}

function initializeUiScale() {
  let storedScale = "small";
  try {
    storedScale = window.localStorage.getItem(UI_SCALE_KEY) || "small";
  } catch (error) {
    console.warn("UI scale could not be loaded.", error);
  }
  setUiScale(storedScale);
  document.querySelectorAll(".scale-btn").forEach((button) => {
    button.addEventListener("click", () => {
      setUiScale(button.dataset.uiScale);
    });
  });
}

function toggleAllRows(group, checked) {
  document.querySelectorAll(`.row-chk-${group}`).forEach((checkbox) => {
    checkbox.checked = checked;
    const row = checkbox.closest("tr");
    if (row) {
      row.classList.toggle("selected", checked);
    }
  });
  updateCheckedRows(group);
}

function clearAllRows(group) {
  document.querySelectorAll(`.row-chk-${group}`).forEach((checkbox) => {
    checkbox.checked = false;
    const row = checkbox.closest("tr");
    if (row) {
      row.classList.remove("selected");
    }
  });
  const master = document.getElementById(`chk-all-${group}`);
  if (master) {
    master.checked = false;
  }
  updateCheckedRows(group);
}

function updateCheckedRows(group) {
  const rows = document.querySelectorAll(`.row-chk-${group}`);
  const checked = document.querySelectorAll(`.row-chk-${group}:checked`);
  const count = checked.length;
  const label = document.getElementById(`sel-count-${group}`);
  if (label) {
    label.textContent = `${count}件選択中`;
  }
  rows.forEach((checkbox) => {
    const row = checkbox.closest("tr");
    if (row) {
      row.classList.toggle("selected", checkbox.checked);
    }
  });
  const master = document.getElementById(`chk-all-${group}`);
  if (master) {
    master.checked = rows.length > 0 && count === rows.length;
  }
}

document.addEventListener("DOMContentLoaded", () => {
  initializeUiScale();

  window.setTimeout(() => {
    document.querySelectorAll(".flash").forEach((node) => {
      node.style.opacity = "0";
      node.style.transition = "opacity 0.3s ease";
      window.setTimeout(() => node.remove(), 320);
    });
  }, 4000);
});
