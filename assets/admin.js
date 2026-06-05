const DB_NAME = "logistics-info-library";
const DB_VERSION = 2;
const FACT_STORE = "fact-files";
const SOURCE_SLOT_ID = "fact-1";
const LOCAL_LIBRARY_SOURCE = "local-upload";

const librarySlots = [
  {
    store: FACT_STORE,
    id: SOURCE_SLOT_ID,
    library: "事实表",
    label: "物流原表",
    description: "包含医疗-发货、医疗-退货、医疗-单独发货、医疗-专线"
  }
];

const adminEls = {
  slots: document.querySelector("#adminLibrarySlots"),
  clearCacheButton: document.querySelector("#clearLibraryCacheButton"),
  referenceState: document.querySelector("#adminReferenceState"),
  referenceRows: document.querySelector("#adminReferenceRows")
};

const adminState = {
  records: new Map()
};

function openAppDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      ["files", FACT_STORE].forEach((storeName) => {
        if (!db.objectStoreNames.contains(storeName)) {
          db.createObjectStore(storeName, { keyPath: "id" });
        }
      });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

function runStoreRequest(db, storeName, mode, createRequest) {
  return new Promise((resolve, reject) => {
    const transaction = db.transaction(storeName, mode);
    const request = createRequest(transaction.objectStore(storeName));
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    transaction.onerror = () => reject(transaction.error);
  });
}

function getRecord(db, storeName, key) {
  return runStoreRequest(db, storeName, "readonly", (store) => store.get(key));
}

function putRecord(db, storeName, record) {
  return runStoreRequest(db, storeName, "readwrite", (store) => store.put(record));
}

function deleteRecord(db, storeName, key) {
  return runStoreRequest(db, storeName, "readwrite", (store) => store.delete(key));
}

function deleteLibraryDatabase() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.deleteDatabase(DB_NAME);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("文件库正在被其他页面占用，请关闭其他看板页面后重试"));
  });
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        if (!window.XLSX) throw new Error("XLSX 解析库未加载");
        const workbook = XLSX.read(reader.result, { type: "array", cellDates: true });
        const sheets = workbook.SheetNames.map((sheetName) => ({
          name: sheetName,
          rows: XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            header: 1,
            defval: "",
            raw: false
          })
        }));
        resolve(sheets);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function bindAdminEvents() {
  adminEls.slots.addEventListener("change", async (event) => {
    const input = event.target.closest("[data-admin-upload]");
    if (!input) return;
    await saveFile(input.dataset.adminUpload, input.files[0]);
    input.value = "";
  });

  adminEls.slots.addEventListener("click", async (event) => {
    const applyButton = event.target.closest("[data-admin-apply]");
    if (applyButton) {
      await applySlot(applyButton.dataset.adminApply);
      return;
    }

    const deleteButton = event.target.closest("[data-admin-delete]");
    if (deleteButton) {
      await deleteSlot(deleteButton.dataset.adminDelete);
    }
  });

  adminEls.clearCacheButton.addEventListener("click", clearLibraryCache);
}

async function refreshAdmin() {
  const db = await openAppDb();
  const entries = await Promise.all(
    librarySlots.map(async (slot) => [slot.id, await getRecord(db, slot.store, slot.id)])
  );
  db.close();
  adminState.records = new Map(entries);
  renderLibrarySlots();
  renderReferenceRows();
}

async function saveFile(slotId, file) {
  if (!file) return;
  const slot = getSlot(slotId);
  const savedAt = new Date().toISOString();
  adminEls.referenceState.textContent = "读取文件中";
  const sheets = await readWorkbook(file);
  const existing = adminState.records.get(slotId) || { id: slotId };
  const record = {
    ...existing,
    id: slotId,
    pendingName: file.name,
    pendingSize: file.size,
    pendingTypeLabel: getFileTypeLabel(file),
    pendingRefreshMonth: getRefreshMonth(file.name, savedAt),
    pendingSavedAt: savedAt,
    pendingLibrarySource: LOCAL_LIBRARY_SOURCE,
    pendingSheets: sheets
  };
  const db = await openAppDb();
  await putRecord(db, slot.store, record);
  db.close();
  await refreshAdmin();
}

async function applySlot(slotId) {
  const slot = getSlot(slotId);
  const record = adminState.records.get(slotId);
  if (!record) return;
  const appliedAt = new Date().toISOString();
  const updatedRecord = record.pendingSheets
    ? clearPendingFields({
        ...record,
        name: record.pendingName,
        size: record.pendingSize,
        typeLabel: record.pendingTypeLabel,
        refreshMonth: record.pendingRefreshMonth,
        savedAt: record.pendingSavedAt,
        librarySource: LOCAL_LIBRARY_SOURCE,
        sheets: record.pendingSheets,
        applied: true,
        appliedAt
      })
    : {
        ...record,
        librarySource: record.librarySource || LOCAL_LIBRARY_SOURCE,
        applied: true,
        appliedAt
      };
  const db = await openAppDb();
  await putRecord(db, slot.store, updatedRecord);
  db.close();
  await refreshAdmin();
}

async function deleteSlot(slotId) {
  const slot = getSlot(slotId);
  const db = await openAppDb();
  await deleteRecord(db, slot.store, slotId);
  db.close();
  await refreshAdmin();
}

async function clearLibraryCache() {
  const confirmed = window.confirm("确认清除当前浏览器里的所有文件库缓存吗？清除后需要重新上传并确认应用。");
  if (!confirmed) return;
  adminEls.clearCacheButton.disabled = true;
  adminEls.referenceState.textContent = "清除中";
  try {
    await deleteLibraryDatabase();
    adminState.records = new Map();
    renderLibrarySlots();
    renderReferenceRows();
    adminEls.referenceState.textContent = "已清除缓存";
  } catch (error) {
    console.warn("clear library cache failed", error);
    adminEls.referenceState.textContent = "清除失败";
  } finally {
    adminEls.clearCacheButton.disabled = false;
  }
}

function renderLibrarySlots() {
  adminEls.slots.innerHTML = librarySlots.map(renderLibrarySlot).join("");
}

function renderLibrarySlot(slot) {
  const record = adminState.records.get(slot.id);
  const display = getDisplayRecord(record);
  const isApplied = Boolean(record?.applied && !record.pendingSheets);
  const hasPending = Boolean(record?.pendingSheets);
  const statusText = isApplied ? "已引用" : hasPending ? "待应用" : "未上传";
  return `
    <article class="admin-file-card ${isApplied ? "applied" : ""}">
      <div class="admin-file-card-head">
        <div>
          <p class="eyebrow">${escapeHtml(slot.library)}</p>
          <h2>${escapeHtml(slot.label)}</h2>
        </div>
        <span class="slot-status ${isApplied ? "applied" : hasPending ? "pending" : ""}">${statusText}</span>
      </div>
      <div class="admin-file-meta">
        <span>${escapeHtml(display?.name || "未上传文件")}</span>
        <strong>${display ? `${escapeHtml(display.typeLabel || "文件")} / ${formatFileSize(display.size)}` : "--"}</strong>
        <small>${escapeHtml(slot.description)}</small>
        <small>更新：${display ? formatDateTime(display.savedAt) : "--"}</small>
      </div>
      <div class="admin-file-actions">
        <label class="admin-upload-button">
          <input type="file" accept=".xlsx,.xls,.csv" data-admin-upload="${slot.id}">
          上传/替换
        </label>
        <button type="button" data-admin-apply="${slot.id}" ${hasPending || (record && !record.applied) ? "" : "disabled"}>确认应用</button>
        <button class="danger-button" type="button" data-admin-delete="${slot.id}" ${record ? "" : "disabled"}>删除</button>
      </div>
    </article>
  `;
}

function renderReferenceRows() {
  adminEls.referenceRows.innerHTML = librarySlots.map((slot) => renderReferenceRow(slot, adminState.records.get(slot.id))).join("");
  adminEls.referenceState.textContent = "本地文件库";
}

function renderReferenceRow(slot, record) {
  const applied = Boolean(record?.applied);
  return `
    <tr>
      <td>${escapeHtml(slot.library)}</td>
      <td>${escapeHtml(slot.label)}</td>
      <td>${escapeHtml(record?.name || "--")}</td>
      <td>${escapeHtml(record?.refreshMonth || "--")}</td>
      <td>${formatDateTime(record?.savedAt)}</td>
      <td>${formatDateTime(record?.appliedAt || "")}</td>
      <td><span class="slot-status ${applied ? "applied" : "pending"}">${applied ? "已引用" : "未引用"}</span></td>
    </tr>
  `;
}

function getSlot(slotId) {
  return librarySlots.find((slot) => slot.id === slotId);
}

function getDisplayRecord(record) {
  if (!record) return null;
  if (record.pendingSheets) {
    return {
      name: record.pendingName,
      size: record.pendingSize,
      typeLabel: record.pendingTypeLabel,
      savedAt: record.pendingSavedAt
    };
  }
  return record;
}

function clearPendingFields(record) {
  const nextRecord = { ...record };
  delete nextRecord.pendingName;
  delete nextRecord.pendingSize;
  delete nextRecord.pendingTypeLabel;
  delete nextRecord.pendingRefreshMonth;
  delete nextRecord.pendingSavedAt;
  delete nextRecord.pendingLibrarySource;
  delete nextRecord.pendingSheets;
  return nextRecord;
}

function getRefreshMonth(fileName, fallbackDate) {
  const match = String(fileName).match(/(20\d{2})[-_年 ]?(0?[1-9]|1[0-2])月?/);
  if (match) return `${match[1]}-${String(match[2]).padStart(2, "0")}`;
  const date = new Date(fallbackDate);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function formatDateTime(value) {
  if (!value) return "--";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "--";
  return new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Shanghai",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false
  }).format(date);
}

function formatFileSize(bytes) {
  if (!Number.isFinite(bytes)) return "--";
  if (bytes < 1024) return `${bytes} B`;
  const units = ["KB", "MB", "GB"];
  let size = bytes / 1024;
  let unitIndex = 0;
  while (size >= 1024 && unitIndex < units.length - 1) {
    size /= 1024;
    unitIndex += 1;
  }
  return `${size.toFixed(size >= 10 ? 1 : 2)} ${units[unitIndex]}`;
}

function getFileTypeLabel(file) {
  const extension = file.name.split(".").pop()?.toLowerCase();
  if (extension === "xlsx" || extension === "xls") return "Excel 工作簿";
  if (extension === "csv") return "CSV 文件";
  return file.type || "未知类型";
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  })[char]);
}

bindAdminEvents();
refreshAdmin().catch((error) => {
  console.error(error);
  adminEls.referenceState.textContent = "读取失败";
});
