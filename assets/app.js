const DB_NAME = "logistics-info-library";
const DB_VERSION = 2;
const LEGACY_STORE_NAME = "files";
const FACT_STORE_NAME = "fact-files";
const LEGACY_SOURCE_ID = "fact:logistics-source";
const LOGISTICS_SLOT_ID = "fact-1";
const DIMENSION_SLOT_ID = "fact-2";

const BUSINESS_SHEETS = ["医疗-发货", "医疗-退货", "医疗-单独发货", "医疗-专线"];

const FIELD_ALIASES = {
  month: ["取走数据月份", "取数月份"],
  logistics: ["物流", "物流公司", "发货物流"],
  subject: ["公司主体", "主体", "主体_1", "付款主体"],
  department: ["财务回传部门1"],
  shop: ["财务回传店铺1"],
  productLine: ["产品线", "产品线_1"],
  orderNo: ["有效订单号", "订单号", "订单号/审批单号（财务需求必填）", "OA审批单号", "查询编码"],
  trackingNo: ["有效物流单号", "物流单号", "单号", "返货单号", "出入库单号"],
  remark: ["备注", "备注_2", "特殊原因", "其他费用产生原因"],
  quantity: ["货品数量", "数量", "数量_1"]
};

const AMOUNT_RULES = {
  "医疗-发货": {
    tax: ["实际发货运费（含税）"],
    net: ["实际发货运费(不含税）"]
  },
  "医疗-退货": {
    tax: ["含税运费"],
    net: ["不含税运费"]
  },
  "医疗-单独发货": {
    tax: ["含税运费"],
    net: ["不含税运费"]
  },
  "医疗-专线": {
    tax: ["运费总计"],
    netFromTaxDivisor: 1.09
  }
};

const filterState = {
  month: new Set(),
  category: new Set(),
  logistics: new Set(),
  subject: new Set(),
  department: new Set(),
  shop: new Set(),
  productLine: new Set()
};

let normalizedRows = [];
let filteredRows = [];
let dimensionMaps = {
  department: new Map(),
  shop: new Map()
};

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      [LEGACY_STORE_NAME, FACT_STORE_NAME].forEach((storeName) => {
        if (!db.objectStoreNames.contains(storeName)) {
          db.createObjectStore(storeName, { keyPath: "id" });
        }
      });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function getAllLegacyFiles() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(LEGACY_STORE_NAME, "readonly");
    const request = tx.objectStore(LEGACY_STORE_NAME).getAll();
    request.onsuccess = () => resolve(request.result || []);
    request.onerror = () => reject(request.error);
    tx.oncomplete = () => db.close();
  });
}

async function getStoreRecord(storeName, key) {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, "readonly");
    const request = tx.objectStore(storeName).get(key);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
    tx.oncomplete = () => db.close();
  });
}

async function getAppliedRecord(slotId, legacyId = null) {
  const record = await getStoreRecord(FACT_STORE_NAME, slotId);
  const normalizedRecord = normalizeLibraryRecord(record);
  if (normalizedRecord?.sheets?.length) return normalizedRecord;
  if (!legacyId) return null;
  return getStoreRecord(LEGACY_STORE_NAME, legacyId);
}

function normalizeLibraryRecord(record) {
  if (!record) return null;
  if (record.sheets?.length) return record;
  if (record.pendingSheets?.length) {
    return {
      ...record,
      name: record.pendingName || record.name,
      size: record.pendingSize || record.size,
      savedAt: record.pendingSavedAt || record.savedAt,
      sheets: record.pendingSheets
    };
  }
  return record;
}

async function clearFiles() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction([LEGACY_STORE_NAME, FACT_STORE_NAME], "readwrite");
    tx.objectStore(LEGACY_STORE_NAME).clear();
    tx.objectStore(FACT_STORE_NAME).clear();
    tx.oncomplete = () => {
      db.close();
      resolve();
    };
    tx.onerror = () => reject(tx.error);
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

function normalizeKey(value) {
  return String(value ?? "").trim();
}

function isEmptyRow(row) {
  return row.every((value) => normalizeKey(value) === "");
}

function makeUniqueHeaders(headers) {
  const seen = new Map();
  return headers.map((header, index) => {
    const base = normalizeKey(header) || `Column${index + 1}`;
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
    if (normalizeKey(value) !== "") return normalizeKey(value);
  }
  return "";
}

function getSheetAmount(row, sheetName, kind) {
  const rule = AMOUNT_RULES[sheetName];
  if (!rule) return 0;
  if (kind === "net" && rule.netFromTaxDivisor) {
    return getSheetAmount(row, sheetName, "tax") / rule.netFromTaxDivisor;
  }
  const aliases = kind === "tax" ? rule.tax : rule.net;
  return parseAmount(firstValue(row, aliases || []));
}

function buildCodeMap(record, sheetKeyword) {
  const map = new Map();
  const sheet = record?.sheets?.find((item) => item.name.includes(sheetKeyword));
  if (!sheet?.rows?.length) return map;

  sheet.rows.forEach((row) => {
    const code = normalizeKey(row[0]);
    const name = normalizeKey(row[1]);
    if (name && code && !map.has(name)) map.set(name, code);
  });
  return map;
}

function buildDimensionMaps(record) {
  return {
    department: buildCodeMap(record, "部门维度信息"),
    shop: buildCodeMap(record, "店铺维度信息")
  };
}

function normalizeSource(record, maps) {
  if (!record?.sheets) return [];
  return record.sheets
    .filter((sheet) => BUSINESS_SHEETS.includes(sheet.name))
    .flatMap((sheet) => {
      const parsed = rowsToObjects(sheet.rows);
      return parsed.rows.map((row, index) => {
        const department = firstValue(row, FIELD_ALIASES.department);
        const shop = firstValue(row, FIELD_ALIASES.shop);
        const taxAmount = getSheetAmount(row, sheet.name, "tax");
        const netAmount = getSheetAmount(row, sheet.name, "net");
        return {
          id: `${sheet.name}:${index}`,
          sourceSheet: sheet.name,
          month: firstValue(row, FIELD_ALIASES.month),
          category: sheet.name,
          logistics: firstValue(row, FIELD_ALIASES.logistics),
          subject: firstValue(row, FIELD_ALIASES.subject),
          department,
          departmentCode: maps.department.get(department) || "",
          shop,
          shopCode: maps.shop.get(shop) || "",
          productLine: firstValue(row, FIELD_ALIASES.productLine),
          orderNo: firstValue(row, FIELD_ALIASES.orderNo),
          trackingNo: firstValue(row, FIELD_ALIASES.trackingNo),
          remark: firstValue(row, FIELD_ALIASES.remark),
          quantity: parseAmount(firstValue(row, FIELD_ALIASES.quantity)),
          taxAmount,
          netAmount,
          raw: row
        };
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
  const totalRecords = source?.sheets?.reduce((sum, sheet) => {
    if (!BUSINESS_SHEETS.includes(sheet.name)) return sum;
    return sum + Math.max(sheet.rows.length - 1, 0);
  }, 0) || 0;
  setText("[data-total-files]", totalFiles);
  setText("[data-source-status]", source ? "已上传" : "未上传");
  setText("[data-total-records]", formatInteger(totalRecords));
  setText("[data-sheet-count]", source?.sheets?.length || 0);
  setText("[data-saved-time]", formatDateTime(source?.savedAt));
}

function uniqueOptions(rows, key) {
  return Array.from(new Set(rows.map((row) => row[key]).filter(Boolean))).sort((a, b) => {
    return String(a).localeCompare(String(b), "zh-CN");
  });
}

function labelForFilter(container, selectedCount) {
  const key = container.dataset.filter;
  const allLabel = container.dataset.label || "全部";
  const selected = Array.from(filterState[key] || []);
  if (!selected.length) return allLabel;
  if (selected.length === 1) return selected[0];
  if (selected.length === 2) return selected.join("、");
  return `已选${selectedCount}项`;
}

function renderFilters() {
  document.querySelectorAll(".multi-filter").forEach((container) => {
    const key = container.dataset.filter;
    const values = uniqueOptions(normalizedRows, key);
    const selected = filterState[key] || new Set();
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
      row.shopCode,
      row.department,
      row.departmentCode,
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
      row.category,
      row.logistics,
      row.subject,
      row.department,
      row.departmentCode,
      row.shop,
      row.shopCode,
      row.productLine
    ].map((value) => value || "未填");
    const key = parts.join("\u0001");
    if (!map.has(key)) {
      map.set(key, {
        category: parts[0],
        logistics: parts[1],
        subject: parts[2],
        department: parts[3],
        departmentCode: parts[4],
        shop: parts[5],
        shopCode: parts[6],
        productLine: parts[7],
        records: 0,
        taxAmount: 0,
        netAmount: 0
      });
    }
    const item = map.get(key);
    item.records += 1;
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
  const rows = groupSummary(filteredRows).slice(0, 500);
  renderTable(document.querySelector("[data-summary-table]"), [
    { key: "category", label: "分类" },
    { key: "logistics", label: "物流" },
    { key: "subject", label: "主体" },
    { key: "department", label: "财务回传部门" },
    { key: "departmentCode", label: "部门编码" },
    { key: "shop", label: "财务回传店铺" },
    { key: "shopCode", label: "店铺客户编码" },
    { key: "productLine", label: "产品线" },
    { key: "taxAmount", label: "运费合计(含税)", number: true, format: formatNumber },
    { key: "netAmount", label: "运费合计不含税", number: true, format: formatNumber }
  ], rows, "暂无可汇总数据");
}

function renderDirectoryTable() {
  const rows = groupSummary(filteredRows);
  const visibleRows = rows.slice(0, 500);
  renderTable(document.querySelector("[data-directory-table]"), [
    { key: "category", label: "分类" },
    { key: "logistics", label: "物流" },
    { key: "subject", label: "主体" },
    { key: "department", label: "财务回传部门" },
    { key: "departmentCode", label: "部门编码" },
    { key: "shop", label: "财务回传店铺" },
    { key: "shopCode", label: "店铺客户编码" },
    { key: "productLine", label: "产品线" },
    { key: "taxAmount", label: "运费合计(含税)", number: true, format: formatNumber },
    { key: "netAmount", label: "运费合计不含税", number: true, format: formatNumber }
  ], visibleRows, "暂无检索结果");

  const state = document.querySelector("[data-directory-state]");
  if (state) {
    state.textContent = rows.length ? `当前 ${formatInteger(rows.length)} 条` : "等待数据";
  }
  const downloadButton = document.querySelector("[data-directory-download]");
  if (downloadButton) {
    downloadButton.disabled = rows.length === 0;
  }
}

function renderDetailTable() {
  const rows = filteredRows.slice(0, 300);
  renderTable(document.querySelector("[data-detail-table]"), [
    { key: "month", label: "月份" },
    { key: "category", label: "分类" },
    { key: "logistics", label: "物流" },
    { key: "subject", label: "主体" },
    { key: "department", label: "财务回传部门" },
    { key: "departmentCode", label: "部门编码" },
    { key: "shop", label: "财务回传店铺" },
    { key: "shopCode", label: "店铺客户编码" },
    { key: "productLine", label: "产品线" },
    { key: "orderNo", label: "订单号/审批单号" },
    { key: "trackingNo", label: "物流单号" },
    { key: "taxAmount", label: "含税运费", number: true, format: formatNumber },
    { key: "netAmount", label: "不含税运费", number: true, format: formatNumber },
    { key: "remark", label: "备注" }
  ], rows, "暂无明细数据");
}

function rerenderDashboard() {
  applyFilters();
  renderFilters();
  renderKpis();
  renderDirectoryTable();
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

  const rows = groupSummary(filteredRows).map((row) => ({
    分类: row.category,
    物流: row.logistics,
    主体: row.subject,
    财务回传部门: row.department,
    部门编码: row.departmentCode,
    财务回传店铺: row.shop,
    店铺客户编码: row.shopCode,
    产品线: row.productLine,
    "运费合计(含税)": row.taxAmount,
    运费合计不含税: row.netAmount
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "物流汇总");
  const stamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
  XLSX.writeFile(workbook, `物流汇总_${stamp}.xlsx`);
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
    if (!filterState[key]) return;
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
  document.querySelector("[data-directory-download]")?.addEventListener("click", exportRows);
}

function bindEvents() {
  document.querySelectorAll("[data-clear-local]").forEach((button) => {
    button.addEventListener("click", async () => {
      await clearFiles();
      await refresh();
    });
  });
  bindFilterEvents();
}

async function refresh() {
  const records = await getAllLegacyFiles();
  const [source, dimensionRecord] = await Promise.all([
    getAppliedRecord(LOGISTICS_SLOT_ID, LEGACY_SOURCE_ID),
    getAppliedRecord(DIMENSION_SLOT_ID)
  ]);
  dimensionMaps = buildDimensionMaps(dimensionRecord);
  normalizedRows = normalizeSource(source, dimensionMaps);
  renderSourceMeta(records, source);
  rerenderDashboard();
}

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  refresh().catch((error) => {
    console.error("Failed to initialize logistics info library.", error);
  });
});
