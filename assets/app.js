const DB_NAME = "logistics-info-library";
const DB_VERSION = 2;
const STORE_NAME = "files";
const FACT_STORE_NAME = "fact-files";
const SOURCE_ID = "fact:logistics-source";
const SOURCE_SLOT_ID = "fact-1";

const EXPECTED_SHEETS = ["医疗-发货", "医疗-退货", "医疗-单独发货", "医疗-专线"];

const FIELD_ALIASES = {
  month: ["取走数据月份", "取数月份"],
  logistics: ["物流公司", "物流", "发货物流"],
  category: ["分类", "商品分类", "产品线", "产品线_1"],
  subject: ["公司主体", "主体", "主体_1", "付款主体"],
  department: ["财务回传部门1", "财务回传部门2", "事业部", "阿米巴"],
  shop: ["财务回传店铺1", "财务回传店铺2", "店铺", "店铺（必填）"],
  productLine: ["产品线", "产品线_1", "架构", "商品分类"],
  orderNo: ["有效订单号", "订单号", "订单号/审批单号（财务需求必填）", "OA审批单号", "查询编码"],
  trackingNo: ["有效物流单号", "物流单号", "单号", "返货单号", "出入库单号"],
  taxAmount: ["实际发货运费（含税）", "含税运费", "运费总计", "运费"],
  netAmount: ["实际发货运费(不含税）", "不含税运费", "运费不含税金额"],
  quantity: ["货品数量", "数量", "数量_1"],
  remark: ["备注", "备注_2", "特殊原因", "其他费用产生原因"]
};

const filterState = {
  sourceSheet: new Set(),
  month: new Set(),
  logistics: new Set(),
  category: new Set(),
  subject: new Set(),
  department: new Set(),
  shop: new Set(),
  productLine: new Set()
};

let normalizedRows = [];
let filteredRows = [];

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: "id" });
      }
      if (!db.objectStoreNames.contains(FACT_STORE_NAME)) {
        db.createObjectStore(FACT_STORE_NAME, { keyPath: "id" });
      }
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getAllFiles() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const request = tx.objectStore(STORE_NAME).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
  });
}

async function getSourceRecord() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readonly");
    const request = tx.objectStore(STORE_NAME).get(SOURCE_ID);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function getAppliedSourceRecord() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction([FACT_STORE_NAME, STORE_NAME], "readonly");
    const factRequest = tx.objectStore(FACT_STORE_NAME).get(SOURCE_SLOT_ID);
    let legacyRecord = null;

    factRequest.onsuccess = () => {
      const record = factRequest.result;
      if (record?.applied && record?.sheets?.length) {
        resolve(record);
        return;
      }
      const legacyRequest = tx.objectStore(STORE_NAME).get(SOURCE_ID);
      legacyRequest.onsuccess = () => {
        legacyRecord = legacyRequest.result || null;
      };
      legacyRequest.onerror = () => reject(legacyRequest.error);
    };
    factRequest.onerror = () => reject(factRequest.error);
    tx.oncomplete = () => {
      if (legacyRecord) resolve(legacyRecord);
      else resolve(null);
    };
    tx.onerror = () => reject(tx.error);
  });
}

async function saveSourceRecord(record) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).put(record);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function clearFiles() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, "readwrite");
    tx.objectStore(STORE_NAME).clear();
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      try {
        if (!window.XLSX) {
          throw new Error("XLSX 解析库未加载");
        }

        const workbook = XLSX.read(reader.result, { type: "array", cellDates: true });
        const sheets = workbook.SheetNames.map((sheetName) => {
          const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
            header: 1,
            defval: "",
            raw: false
          });
          return {
            name: sheetName,
            rows
          };
        });
        resolve(sheets);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatNumber(value, digits = 2) {
  return Number(value || 0).toLocaleString("zh-CN", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function formatInteger(value) {
  return Number(value || 0).toLocaleString("zh-CN", { maximumFractionDigits: 0 });
}

function parseAmount(value) {
  if (value === null || value === undefined || value === "") return 0;
  const cleaned = String(value).replace(/[,\s￥¥]/g, "");
  const number = Number(cleaned);
  return Number.isFinite(number) ? number : 0;
}

function formatDateTime(value) {
  if (!value) return "-";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "-";
  return date.toLocaleString("zh-CN", { hour12: false });
}

function formatSize(size) {
  if (size > 1024 * 1024) return `${(size / 1024 / 1024).toFixed(2)} MB`;
  if (size > 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${size} B`;
}

function isEmptyRow(row) {
  return row.every((value) => String(value ?? "").trim() === "");
}

function makeUniqueHeaders(headers) {
  const seen = new Map();
  return headers.map((header, index) => {
    const base = String(header || `Column${index + 1}`).trim();
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count === 0 ? base : `${base}_${count + 1}`;
  });
}

function rowsToObjects(rows) {
  if (!rows?.length) return { headers: [], rows: [] };
  const headers = makeUniqueHeaders(rows[0]);
  const dataRows = rows.slice(1).filter((row) => !isEmptyRow(row));
  return {
    headers,
    rows: dataRows.map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index] ?? "";
      });
      return item;
    })
  };
}

function firstValue(row, aliases) {
  for (const alias of aliases) {
    const value = row[alias];
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      return String(value).trim();
    }
  }
  return "";
}

function normalizeSource(record) {
  if (!record?.sheets) return [];
  return record.sheets.flatMap((sheet) => {
    const parsed = rowsToObjects(sheet.rows);
    return parsed.rows.map((row, index) => {
      const normalized = {
        id: `${sheet.name}:${index}`,
        sourceSheet: sheet.name,
        month: firstValue(row, FIELD_ALIASES.month),
        logistics: firstValue(row, FIELD_ALIASES.logistics),
        category: firstValue(row, FIELD_ALIASES.category),
        subject: firstValue(row, FIELD_ALIASES.subject),
        department: firstValue(row, FIELD_ALIASES.department),
        shop: firstValue(row, FIELD_ALIASES.shop),
        productLine: firstValue(row, FIELD_ALIASES.productLine),
        orderNo: firstValue(row, FIELD_ALIASES.orderNo),
        trackingNo: firstValue(row, FIELD_ALIASES.trackingNo),
        remark: firstValue(row, FIELD_ALIASES.remark),
        taxAmount: parseAmount(firstValue(row, FIELD_ALIASES.taxAmount)),
        netAmount: parseAmount(firstValue(row, FIELD_ALIASES.netAmount)),
        quantity: parseAmount(firstValue(row, FIELD_ALIASES.quantity)),
        raw: row
      };
      if (!normalized.netAmount && normalized.taxAmount) {
        normalized.netAmount = normalized.taxAmount / 1.06;
      }
      return normalized;
    });
  });
}

function setText(selector, value) {
  document.querySelectorAll(selector).forEach((node) => {
    node.textContent = value;
  });
}

function renderSourceMeta(records, source) {
  const totalFiles = source ? 1 : records.length;
  const totalRecords = source?.sheets?.reduce((sum, sheet) => sum + Math.max(sheet.rows.length - 1, 0), 0) || 0;
  setText("[data-total-files]", totalFiles);
  setText("[data-source-status]", source ? "已上传" : "未上传");
  setText("[data-total-records]", formatInteger(totalRecords));
  setText("[data-sheet-count]", source?.sheets?.length || 0);
  setText("[data-saved-time]", formatDateTime(source?.savedAt));
}

function renderSheetList(source) {
  const container = document.querySelector("[data-sheet-list]");
  if (!container) return;

  if (!source?.sheets?.length) {
    container.innerHTML = `<div class="file-item"><div><strong>暂无文件</strong><small>上传物流原表后显示 sheet 清单。</small></div></div>`;
    return;
  }

  container.innerHTML = source.sheets.map((sheet) => {
    const expected = EXPECTED_SHEETS.includes(sheet.name);
    return `
      <div class="file-item">
        <div>
          <strong>${escapeHtml(sheet.name)}</strong>
          <small>${escapeHtml(source.name)} · ${formatSize(source.size)} · ${formatDateTime(source.savedAt)}</small>
        </div>
        <span class="badge">${expected ? "业务表" : "附加表"} · ${formatInteger(Math.max(sheet.rows.length - 1, 0))} 行</span>
      </div>
    `;
  }).join("");
}

function renderPreview(source) {
  const table = document.querySelector("[data-preview-table]");
  if (!table) return;

  const sheet = source?.sheets?.find((item) => EXPECTED_SHEETS.includes(item.name)) || source?.sheets?.[0];
  if (!sheet?.rows?.length) {
    table.innerHTML = `<tbody><tr><td class="empty-state">暂无预览数据</td></tr></tbody>`;
    return;
  }

  table.innerHTML = sheet.rows.slice(0, 20).map((row, rowIndex) => {
    const tag = rowIndex === 0 ? "th" : "td";
    const cells = row.map((cell) => `<${tag}>${escapeHtml(cell)}</${tag}>`).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
}

function uniqueOptions(rows, key) {
  return Array.from(new Set(rows.map((row) => row[key]).filter(Boolean))).sort((a, b) => {
    return String(a).localeCompare(String(b), "zh-CN");
  });
}

function labelForFilter(container, selectedCount) {
  const key = container.dataset.filter;
  const allLabel = container.dataset.label || "全部";
  const selected = Array.from(filterState[key]);
  if (!selected.length) return allLabel;
  if (selected.length === 1) return selected[0];
  if (selected.length === 2) return selected.join("、");
  return `已选${selectedCount}项`;
}

function renderFilters() {
  document.querySelectorAll(".multi-filter").forEach((container) => {
    const key = container.dataset.filter;
    const values = uniqueOptions(normalizedRows, key);
    const selected = filterState[key];
    const allChecked = selected.size === 0 ? "checked" : "";
    container.innerHTML = `
      <button class="multi-filter-button" type="button">${escapeHtml(labelForFilter(container, selected.size))}</button>
      <div class="multi-filter-menu">
        <label class="multi-filter-option">
          <input type="checkbox" value="" data-all ${allChecked}>全部
        </label>
        ${values.map((value) => `
          <label class="multi-filter-option">
            <input type="checkbox" value="${escapeHtml(value)}" ${selected.has(value) ? "checked" : ""}>${escapeHtml(value)}
          </label>
        `).join("")}
      </div>
    `;
  });
}

function applyFilters() {
  const keyword = document.querySelector("[data-keyword]")?.value.trim().toLowerCase() || "";
  filteredRows = normalizedRows.filter((row) => {
    const matchesFilters = Object.entries(filterState).every(([key, selected]) => {
      return selected.size === 0 || selected.has(row[key]);
    });
    if (!matchesFilters) return false;
    if (!keyword) return true;
    const haystack = [
      row.orderNo,
      row.trackingNo,
      row.shop,
      row.logistics,
      row.category,
      row.remark,
      ...Object.values(row.raw || {})
    ].join(" ").toLowerCase();
    return haystack.includes(keyword);
  });
}

function renderKpis() {
  const taxAmount = filteredRows.reduce((sum, row) => sum + row.taxAmount, 0);
  const netAmount = filteredRows.reduce((sum, row) => sum + row.netAmount, 0);
  const logisticsCount = new Set(filteredRows.map((row) => row.logistics).filter(Boolean)).size;
  setText('[data-kpi="records"]', formatInteger(filteredRows.length));
  setText('[data-kpi="taxAmount"]', formatNumber(taxAmount));
  setText('[data-kpi="netAmount"]', formatNumber(netAmount));
  setText('[data-kpi="logisticsCount"]', formatInteger(logisticsCount));
}

function groupSummary(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const parts = [
      row.subject,
      row.department,
      row.shop,
      row.productLine,
      row.logistics,
      row.category
    ].map((value) => value || "未填");
    const key = parts.join("\u0001");
    if (!map.has(key)) {
      map.set(key, {
        subject: parts[0],
        department: parts[1],
        shop: parts[2],
        productLine: parts[3],
        logistics: parts[4],
        category: parts[5],
        records: 0,
        quantity: 0,
        taxAmount: 0,
        netAmount: 0
      });
    }
    const item = map.get(key);
    item.records += 1;
    item.quantity += row.quantity;
    item.taxAmount += row.taxAmount;
    item.netAmount += row.netAmount;
  });
  return Array.from(map.values()).sort((a, b) => b.taxAmount - a.taxAmount);
}

function renderTable(table, headers, rows, emptyText) {
  if (!table) return;
  if (!rows.length) {
    table.innerHTML = `<tbody><tr><td class="empty-state">${escapeHtml(emptyText)}</td></tr></tbody>`;
    return;
  }

  const head = `<thead><tr>${headers.map((header) => `<th>${escapeHtml(header.label)}</th>`).join("")}</tr></thead>`;
  const body = rows.map((row) => {
    const cells = headers.map((header) => {
      const value = header.format ? header.format(row[header.key], row) : row[header.key];
      const className = header.number ? ' class="number"' : "";
      return `<td${className}>${escapeHtml(value)}</td>`;
    }).join("");
    return `<tr>${cells}</tr>`;
  }).join("");
  table.innerHTML = `${head}<tbody>${body}</tbody>`;
}

function renderSummaryTable() {
  const rows = groupSummary(filteredRows).slice(0, 300);
  renderTable(document.querySelector("[data-summary-table]"), [
    { key: "subject", label: "主体" },
    { key: "department", label: "财务回传部门" },
    { key: "shop", label: "财务回传店铺" },
    { key: "productLine", label: "产品线" },
    { key: "logistics", label: "物流公司" },
    { key: "category", label: "分类" },
    { key: "records", label: "记录数", number: true, format: formatInteger },
    { key: "quantity", label: "数量", number: true, format: formatNumber },
    { key: "taxAmount", label: "含税运费", number: true, format: formatNumber },
    { key: "netAmount", label: "不含税运费", number: true, format: formatNumber }
  ], rows, "暂无可汇总数据");
}

function renderDetailTable() {
  const rows = filteredRows.slice(0, 300);
  renderTable(document.querySelector("[data-detail-table]"), [
    { key: "sourceSheet", label: "业务表" },
    { key: "month", label: "月份" },
    { key: "orderNo", label: "订单号/审批单号" },
    { key: "trackingNo", label: "物流单号" },
    { key: "logistics", label: "物流公司" },
    { key: "category", label: "分类" },
    { key: "subject", label: "主体" },
    { key: "department", label: "部门" },
    { key: "shop", label: "店铺" },
    { key: "productLine", label: "产品线" },
    { key: "taxAmount", label: "含税运费", number: true, format: formatNumber },
    { key: "netAmount", label: "不含税运费", number: true, format: formatNumber },
    { key: "remark", label: "备注" }
  ], rows, "暂无明细数据");
}

function rerenderDashboard() {
  applyFilters();
  renderFilters();
  renderKpis();
  renderSummaryTable();
  renderDetailTable();
}

async function handleUpload(event) {
  const input = event.currentTarget;
  const file = input.files?.[0];
  if (!file) return;

  const sheets = await readWorkbook(file);
  await saveSourceRecord({
    id: SOURCE_ID,
    kind: "fact",
    slot: "物流原表",
    name: file.name,
    size: file.size,
    sheets,
    savedAt: new Date().toISOString()
  });

  input.value = "";
  await refresh();
}

function exportRows() {
  if (!window.XLSX) {
    alert("XLSX 解析库未加载，无法导出。");
    return;
  }
  if (!filteredRows.length) {
    alert("当前筛选没有可导出的记录。");
    return;
  }

  const rows = filteredRows.map((row) => ({
    业务表: row.sourceSheet,
    月份: row.month,
    订单号或审批单号: row.orderNo,
    物流单号: row.trackingNo,
    物流公司: row.logistics,
    分类: row.category,
    主体: row.subject,
    财务回传部门: row.department,
    财务回传店铺: row.shop,
    产品线: row.productLine,
    数量: row.quantity,
    含税运费: row.taxAmount,
    不含税运费: row.netAmount,
    备注: row.remark
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "物流明细");
  const stamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
  XLSX.writeFile(workbook, `物流明细_${stamp}.xlsx`);
}

function bindFilterEvents() {
  document.addEventListener("click", (event) => {
    const button = event.target.closest(".multi-filter-button");
    if (button) {
      const filter = button.closest(".multi-filter");
      document.querySelectorAll(".multi-filter.open").forEach((node) => {
        if (node !== filter) node.classList.remove("open");
      });
      filter.classList.toggle("open");
      return;
    }
    if (!event.target.closest(".multi-filter")) {
      document.querySelectorAll(".multi-filter.open").forEach((node) => node.classList.remove("open"));
    }
  });

  document.addEventListener("change", (event) => {
    const input = event.target.closest(".multi-filter-menu input");
    if (!input) return;
    const filter = input.closest(".multi-filter");
    const key = filter.dataset.filter;
    if (input.dataset.all !== undefined) {
      filterState[key].clear();
    } else if (input.checked) {
      filterState[key].add(input.value);
    } else {
      filterState[key].delete(input.value);
    }
    rerenderDashboard();
    filter.classList.add("open");
  });

  document.querySelector("[data-keyword]")?.addEventListener("input", rerenderDashboard);
  document.querySelector("[data-clear-filters]")?.addEventListener("click", () => {
    Object.values(filterState).forEach((set) => set.clear());
    const keyword = document.querySelector("[data-keyword]");
    if (keyword) keyword.value = "";
    rerenderDashboard();
  });
  document.querySelector("[data-export]")?.addEventListener("click", exportRows);
}

function bindEvents() {
  document.querySelectorAll("[data-upload-source]").forEach((input) => {
    input.addEventListener("change", (event) => {
      handleUpload(event).catch((error) => {
        console.error("Upload failed.", error);
        alert(`上传失败：${error.message}`);
      });
    });
  });

  document.querySelectorAll("[data-clear-local]").forEach((button) => {
    button.addEventListener("click", async () => {
      await clearFiles();
      await refresh();
    });
  });

  bindFilterEvents();
}

async function refresh() {
  const records = await getAllFiles();
  const source = await getAppliedSourceRecord();
  normalizedRows = normalizeSource(source);
  renderSourceMeta(records, source);
  renderSheetList(source);
  renderPreview(source);
  rerenderDashboard();
}

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  refresh().catch((error) => {
    console.error("Failed to initialize logistics info library.", error);
  });
});
