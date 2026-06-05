const DB_NAME = "logistics-info-library";
const REDIRECT_URL = "logistics.html";

const clearButton = document.querySelector("[data-clear-cache-button]");
const statusEl = document.querySelector("[data-clear-cache-status]");

function setStatus(text, state = "") {
  if (!statusEl) return;
  statusEl.textContent = text;
  statusEl.dataset.state = state;
}

function deleteDatabase(name) {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      resolve();
      return;
    }

    const request = indexedDB.deleteDatabase(name);
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error || new Error("IndexedDB 清除失败"));
    request.onblocked = () => reject(new Error("文件库正在被其它页面占用，请关闭其它物流信息库页面后重试"));
  });
}

async function clearCacheStorage() {
  if (!window.caches) return;
  const keys = await caches.keys();
  await Promise.all(keys.map((key) => caches.delete(key)));
}

async function unregisterServiceWorkers() {
  if (!navigator.serviceWorker) return;
  const registrations = await navigator.serviceWorker.getRegistrations();
  await Promise.all(registrations.map((registration) => registration.unregister()));
}

async function clearBrowserState() {
  localStorage.clear();
  sessionStorage.clear();
  await Promise.all([
    deleteDatabase(DB_NAME),
    clearCacheStorage(),
    unregisterServiceWorkers()
  ]);
}

async function clearAndEnter() {
  if (clearButton) clearButton.disabled = true;
  setStatus("正在清除缓存...", "working");

  try {
    await clearBrowserState();
    setStatus("已清除，正在进入物流大表...", "done");
    window.location.replace(`${REDIRECT_URL}?cacheCleared=${Date.now()}`);
  } catch (error) {
    console.warn("clear cache failed", error);
    setStatus(error.message || "清除失败，请刷新后重试", "error");
    if (clearButton) clearButton.disabled = false;
  }
}

clearButton?.addEventListener("click", clearAndEnter);

if (new URLSearchParams(window.location.search).get("auto") === "1") {
  clearAndEnter();
}
