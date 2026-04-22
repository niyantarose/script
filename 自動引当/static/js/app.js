const UI_SCALE_KEY = "inventory-tool-ui-scale";
const VALID_UI_SCALES = ["small", "medium", "large"];
const DEFAULT_UI_SCALE = "large";

function setUiScale(scale) {
  const nextScale = VALID_UI_SCALES.includes(scale) ? scale : DEFAULT_UI_SCALE;
  document.body.dataset.uiScale = nextScale;
  try {
    window.localStorage.setItem(UI_SCALE_KEY, nextScale);
  } catch (error) {
    console.warn("表示サイズの保存に失敗しました。", error);
  }
  document.querySelectorAll(".scale-btn").forEach((button) => {
    button.classList.toggle("active", button.dataset.uiScale === nextScale);
  });
}

function initializeUiScale() {
  let storedScale = DEFAULT_UI_SCALE;
  try {
    storedScale = window.localStorage.getItem(UI_SCALE_KEY) || DEFAULT_UI_SCALE;
  } catch (error) {
    console.warn("表示サイズの読込に失敗しました。", error);
  }
  setUiScale(storedScale);
  document.querySelectorAll(".scale-btn").forEach((button) => {
    button.addEventListener("click", () => {
      setUiScale(button.dataset.uiScale);
    });
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function showToast(message, kind = "success") {
  const container = document.getElementById("app-toast-container");
  if (!container) {
    return;
  }

  const toast = document.createElement("div");
  toast.className = `toast align-items-center text-bg-${kind === "error" ? "danger" : "dark"} border-0`;
  toast.role = "alert";
  toast.ariaLive = "assertive";
  toast.ariaAtomic = "true";
  toast.innerHTML = `
    <div class="d-flex">
      <div class="toast-body">${escapeHtml(message)}</div>
      <button type="button" class="btn-close btn-close-white me-2 m-auto" data-bs-dismiss="toast" aria-label="閉じる"></button>
    </div>
  `;
  container.appendChild(toast);
  const toastInstance = new bootstrap.Toast(toast, { delay: 2600 });
  toast.addEventListener("hidden.bs.toast", () => toast.remove());
  toastInstance.show();
}

function parseFilename(disposition, fallbackName) {
  if (!disposition) {
    return fallbackName;
  }
  const match = disposition.match(/filename="?([^"]+)"?/i);
  return match ? match[1] : fallbackName;
}

async function downloadCsv(url, fallbackName) {
  const response = await window.fetch(url, { method: "GET" });
  const contentType = response.headers.get("content-type") || "";
  if (!response.ok || !contentType.includes("text/csv")) {
    throw new Error("CSV出力に失敗しました。");
  }

  const blob = await response.blob();
  const filename = parseFilename(response.headers.get("content-disposition"), fallbackName);
  const downloadUrl = window.URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = downloadUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.setTimeout(() => window.URL.revokeObjectURL(downloadUrl), 2000);
}

function initializeExportButtons() {
  document.querySelectorAll("[data-export-url]").forEach((button) => {
    button.addEventListener("click", async () => {
      try {
        await downloadCsv(button.dataset.exportUrl, button.dataset.exportFilename || "export.csv");
      } catch (error) {
        showToast(error.message || "CSV出力に失敗しました。", "error");
      }
    });
  });
}

function flashCellState(cell, className) {
  const element = cell.getElement();
  element.classList.add(className);
  window.setTimeout(() => {
    element.classList.remove(className);
  }, 900);
}

async function persistCell(cell, editConfig) {
  const row = cell.getRow().getData();
  const entity = typeof editConfig.entity === "function"
    ? editConfig.entity(row, cell)
    : editConfig.entity || row[editConfig.entityField];
  const recordId = typeof editConfig.idField === "function"
    ? editConfig.idField(row, cell)
    : row[editConfig.idField];
  const field = typeof editConfig.field === "function"
    ? editConfig.field(row, cell)
    : editConfig.field;
  const nextValue = typeof editConfig.valueTransform === "function"
    ? editConfig.valueTransform(cell.getValue(), row, cell)
    : cell.getValue();

  if (!recordId) {
    return;
  }

  try {
    const response = await window.fetch("/api/update-field", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        entity,
        id: recordId,
        field,
        value: nextValue,
      }),
    });
    const result = await response.json();
    if (!response.ok || !result.ok) {
      throw new Error(result.message || "保存に失敗しました。");
    }
    flashCellState(cell, "cell-saved");
  } catch (error) {
    if (typeof cell.restoreOldValue === "function") {
      cell.restoreOldValue();
    }
    flashCellState(cell, "cell-error");
    showToast(error.message || "保存に失敗しました。", "error");
  }
}

function normalizeColumns(columns) {
  return columns.map((column) => {
    const next = { ...column };
    if (column.editConfig) {
      const originalEdited = column.cellEdited;
      next.editable = column.editable || ((cell) => {
        const row = cell.getRow().getData();
        if (typeof column.editConfig.canEdit === "function") {
          return column.editConfig.canEdit(row, cell);
        }
        if (column.editConfig.requireIdField) {
          return Boolean(row[column.editConfig.requireIdField]);
        }
        return true;
      });
      next.cellEdited = async (cell) => {
        await persistCell(cell, column.editConfig);
        if (typeof originalEdited === "function") {
          originalEdited(cell);
        }
      };
    }
    return next;
  });
}

function selectionColumn() {
  return {
    formatter: "rowSelection",
    titleFormatter: "rowSelection",
    hozAlign: "center",
    headerHozAlign: "center",
    headerSort: false,
    width: 44,
    minWidth: 44,
    frozen: true,
    cellClick: (event, cell) => {
      event.stopPropagation();
      cell.getRow().toggleSelect();
    },
  };
}

function statusBadgeFormatter(field) {
  return (cell) => {
    const row = cell.getRow().getData();
    const label = row[field] || cell.getValue() || "";
    return `<span class="grid-badge">${escapeHtml(label)}</span>`;
  };
}

function choiceBadgeFormatter(choices) {
  return (cell) => {
    const label = choices?.[cell.getValue()] || cell.getValue() || "";
    return `<span class="grid-badge">${escapeHtml(label)}</span>`;
  };
}

function createGrid(config) {
  const element = typeof config.selector === "string" ? document.querySelector(config.selector) : config.selector;
  if (!element || typeof Tabulator === "undefined") {
    return null;
  }

  const table = new Tabulator(element, {
    data: config.data || [],
    index: config.index || "row_id",
    layout: config.layout || "fitDataStretch",
    height: config.height,
    maxHeight: config.maxHeight,
    headerVisible: config.headerVisible ?? true,
    placeholder: config.placeholder || "データがありません。",
    columnDefaults: {
      headerSort: true,
      tooltip: true,
      vertAlign: "middle",
      minWidth: 88,
    },
    movableColumns: false,
    selectableRows: true,
    selectableRowsPersistence: false,
    groupBy: config.groupBy,
    groupStartOpen: config.groupStartOpen ?? true,
    groupHeader: config.groupHeader,
    rowFormatter: config.rowFormatter,
    columns: normalizeColumns(config.columns || []),
    rowSelectionChanged: (data) => {
      if (config.selectionLabelSelector) {
        const label = document.querySelector(config.selectionLabelSelector);
        if (label) {
          label.textContent = `${data.length}件選択中`;
        }
      }
    },
  });

  if (config.initialSort) {
    table.setSort(config.initialSort);
  }

  return table;
}

async function postFormAction(url, payload = {}) {
  const body = new URLSearchParams();
  Object.entries(payload).forEach(([key, value]) => {
    if (value !== undefined && value !== null) {
      body.append(key, value);
    }
  });

  const response = await window.fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8" },
    body,
  });

  if (!response.ok) {
    throw new Error("処理に失敗しました。");
  }
  return response;
}

function tintRowByField(fieldName, classPrefix) {
  return (row) => {
    const level = row.getData()[fieldName];
    if (!level) {
      return;
    }
    row.getElement().classList.add(`${classPrefix}-${level}`);
  };
}

window.InventoryApp = {
  choiceValues: (choices) => ({ ...(choices || {}) }),
  createGrid,
  escapeHtml,
  postFormAction,
  selectionColumn,
  showToast,
  choiceBadgeFormatter,
  statusBadgeFormatter,
  tintRowByField,
};

document.addEventListener("DOMContentLoaded", () => {
  initializeUiScale();
  initializeExportButtons();
});
