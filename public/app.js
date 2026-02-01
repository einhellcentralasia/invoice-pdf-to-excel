const STRINGS = {
  en: {
    title: "Invoice PDF → Excel",
    subtitle: "Upload → Convert → Download",
    stepTitle: "Step 1. Upload PDF",
    stepHint: "We extract items, quantities, and prices into Excel.",
    previewTitle: "Step 2. Choose columns",
    previewHint: "Select the columns you want to export. Page is always included.",
    selectAll: "Select all",
    clearAll: "Clear",
    blockedInfo: "Amount/Total columns are not extractable.",
    blockedNote: "Not extractable",
    processBtn: "Convert",
    downloadOutput: "Download Excel",
    statusReady: "Ready.",
    statusPreview: "Reading headers...",
    statusProcessing: "Processing...",
    statusError: "Error:",
    statusNoFile: "Please select a .pdf file.",
    statusNoColumns: "Select at least one column.",
    statusDone: "Done. File is ready.",
    statusChecking: "API: checking…",
    statusOk: "API: online",
    statusBad: "API: offline",
  },
  ru: {
    title: "PDF счет → Excel",
    subtitle: "Загрузка → Конвертация → Скачивание",
    stepTitle: "Шаг 1. Загрузите PDF",
    stepHint: "Мы извлекаем позиции, количества и цены в Excel.",
    previewTitle: "Шаг 2. Выберите колонки",
    previewHint: "Выберите колонки для выгрузки. Страница всегда включена.",
    selectAll: "Выбрать все",
    clearAll: "Снять все",
    blockedInfo: "Колонки Amount/Total недоступны для выгрузки.",
    blockedNote: "Недоступно",
    processBtn: "Конвертировать",
    downloadOutput: "Скачать Excel",
    statusReady: "Готово.",
    statusPreview: "Чтение заголовков...",
    statusProcessing: "Обработка...",
    statusError: "Ошибка:",
    statusNoFile: "Выберите файл .pdf.",
    statusNoColumns: "Выберите хотя бы одну колонку.",
    statusDone: "Готово. Файл готов к скачиванию.",
    statusChecking: "API: проверка…",
    statusOk: "API: онлайн",
    statusBad: "API: офлайн",
  },
};

const storageKeys = {
  lang: "invoicePdfLang",
  theme: "invoicePdfTheme",
};

const API_BASE = (window.API_BASE || "").replace(/\/+$/, "");

const els = {
  fileInput: document.getElementById("fileInput"),
  processBtn: document.getElementById("processBtn"),
  statusMsg: document.getElementById("statusMsg"),
  downloadBtn: document.getElementById("downloadBtn"),
  themeToggle: document.getElementById("themeToggle"),
  apiStatus: document.getElementById("apiStatus"),
  headersList: document.getElementById("headersList"),
  previewTable: document.getElementById("previewTable"),
  selectAllBtn: document.getElementById("selectAllBtn"),
  clearAllBtn: document.getElementById("clearAllBtn"),
  blockedInfo: document.getElementById("blockedInfo"),
};

let currentLang = "en";
let currentTheme = "dark";
let latestBlobUrl = "";
let latestFilename = "";
let availableHeaders = [];
let blockedHeaders = [];
let sampleRows = [];
let selectedHeaders = new Set();
let alwaysHeaders = ["Page"];

function setTheme(theme) {
  currentTheme = theme;
  document.documentElement.setAttribute("data-theme", theme);
  localStorage.setItem(storageKeys.theme, theme);
  const icon = theme === "dark" ? "☀" : "☾";
  els.themeToggle.textContent = icon;
}

function setLang(lang) {
  currentLang = lang;
  localStorage.setItem(storageKeys.lang, lang);
  const dict = STRINGS[lang];
  document.querySelectorAll("[data-i18n]").forEach((el) => {
    const key = el.dataset.i18n;
    const value = dict[key];
    if (typeof value === "string") {
      el.textContent = value;
    }
  });
  document.querySelectorAll("[data-lang]").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.lang === lang);
  });
  if (els.apiStatus) {
    els.apiStatus.textContent = dict.statusChecking;
  }
  if (els.blockedInfo) {
    els.blockedInfo.textContent = dict.blockedInfo;
  }
  if (availableHeaders.length) {
    renderHeadersList();
  }
  if (els.statusMsg.textContent) {
    els.statusMsg.textContent = dict.statusReady;
    els.statusMsg.className = "msg";
  }
}

function setStatus(message, type = "") {
  els.statusMsg.textContent = message;
  els.statusMsg.className = `msg ${type}`.trim();
}

function toggleLang() {
  const next = currentLang === "en" ? "ru" : "en";
  setLang(next);
}

function setHeaders(headers, blocked, always) {
  availableHeaders = headers || [];
  blockedHeaders = blocked || [];
  alwaysHeaders = always && always.length ? always : ["Page"];
  selectedHeaders = new Set(availableHeaders.filter((h) => !blockedHeaders.includes(h)));
  renderHeadersList();
  renderPreviewTable();
  updateProcessAvailability();
}

function renderHeadersList() {
  if (!els.headersList) return;
  els.headersList.innerHTML = "";

  if (!availableHeaders.length) {
    return;
  }

  availableHeaders.forEach((header) => {
    const item = document.createElement("label");
    item.className = "header-item";
    const isBlocked = blockedHeaders.includes(header);
    if (isBlocked) item.classList.add("disabled");

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.disabled = isBlocked;
    checkbox.checked = selectedHeaders.has(header);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) {
        selectedHeaders.add(header);
      } else {
        selectedHeaders.delete(header);
      }
      updateProcessAvailability();
    });

    const label = document.createElement("span");
    label.textContent = header;

    item.appendChild(checkbox);
    item.appendChild(label);

    if (isBlocked) {
      const note = document.createElement("span");
      note.className = "header-note";
      note.textContent = STRINGS[currentLang].blockedNote;
      item.appendChild(note);
    }

    els.headersList.appendChild(item);
  });

  const pageBadge = document.createElement("span");
  pageBadge.className = "badge info";
  pageBadge.textContent = `${alwaysHeaders.join(", ")} ${currentLang === "ru" ? "всегда включены" : "always included"}`;
  els.headersList.appendChild(pageBadge);
}

function renderPreviewTable() {
  if (!els.previewTable) return;
  els.previewTable.innerHTML = "";
  if (!availableHeaders.length) return;

  const columns = [...availableHeaders];
  alwaysHeaders.forEach((h) => {
    if (!columns.includes(h)) columns.push(h);
  });

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  columns.forEach((col) => {
    const th = document.createElement("th");
    th.textContent = col;
    headRow.appendChild(th);
  });
  thead.appendChild(headRow);
  els.previewTable.appendChild(thead);

  const tbody = document.createElement("tbody");
  sampleRows.forEach((row) => {
    const tr = document.createElement("tr");
    columns.forEach((col) => {
      const td = document.createElement("td");
      td.textContent = row[col] || "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  els.previewTable.appendChild(tbody);
}

function updateProcessAvailability() {
  const hasSelection = selectedHeaders.size > 0;
  els.processBtn.disabled = !els.fileInput.files.length || !hasSelection;
}

function revokeBlob() {
  if (latestBlobUrl) {
    URL.revokeObjectURL(latestBlobUrl);
    latestBlobUrl = "";
  }
}

function downloadOutput() {
  if (!latestBlobUrl) return;
  const link = document.createElement("a");
  link.href = latestBlobUrl;
  link.download = latestFilename || "output.xlsx";
  link.click();
}

async function handleProcess() {
  const dict = STRINGS[currentLang];
  const file = els.fileInput.files[0];
  if (!file) {
    setStatus(dict.statusNoFile, "error");
    return;
  }
  if (!selectedHeaders.size) {
    setStatus(dict.statusNoColumns, "error");
    return;
  }

  setStatus(dict.statusProcessing);
  els.processBtn.disabled = true;
  els.downloadBtn.disabled = true;

  try {
    const formData = new FormData();
    formData.append("file", file, file.name);
    formData.append("columns", JSON.stringify([...selectedHeaders]));
    const res = await fetch(`${API_BASE}/upload`, {
      method: "POST",
      body: formData,
    });
    if (!res.ok) {
      let message = res.statusText;
      try {
        const payload = await res.json();
        if (payload && payload.error === "NO_COLUMNS") {
          message = dict.statusNoColumns;
        } else if (payload && payload.error) {
          message = payload.error;
        }
      } catch (err) {
        const text = await res.text();
        if (text) message = text;
      }
      setStatus(`${dict.statusError} ${message}`, "error");
      return;
    }

    const blob = await res.blob();
    revokeBlob();
    latestBlobUrl = URL.createObjectURL(blob);
    latestFilename = file.name.toLowerCase().endsWith(".pdf")
      ? file.name.replace(/\.pdf$/i, ".xlsx")
      : "output.xlsx";

    els.downloadBtn.disabled = false;
    setStatus(dict.statusDone, "success");
    downloadOutput();
  } catch (err) {
    setStatus(`${dict.statusError} ${err.message}`, "error");
  } finally {
    els.processBtn.disabled = false;
  }
}

async function checkApiStatus() {
  if (!els.apiStatus) return;
  const dict = STRINGS[currentLang];
  els.apiStatus.textContent = dict.statusChecking;
  els.apiStatus.classList.remove("ok", "bad");
  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 5000);
    const res = await fetch(`${API_BASE}/health`, { signal: controller.signal });
    clearTimeout(timeout);
    if (res.ok) {
      els.apiStatus.textContent = dict.statusOk;
      els.apiStatus.classList.add("ok");
    } else {
      els.apiStatus.textContent = dict.statusBad;
      els.apiStatus.classList.add("bad");
    }
  } catch (err) {
    els.apiStatus.textContent = dict.statusBad;
    els.apiStatus.classList.add("bad");
  }
}

async function handlePreview(file) {
  const dict = STRINGS[currentLang];
  setStatus(dict.statusPreview);
  els.downloadBtn.disabled = true;

  try {
    const formData = new FormData();
    formData.append("file", file, file.name);
    const res = await fetch(`${API_BASE}/preview`, {
      method: "POST",
      body: formData,
    });
    const payload = await res.json();
    if (!payload.ok) {
      setStatus(`${dict.statusError} ${payload.error || res.statusText}`, "error");
      return;
    }
    sampleRows = payload.sample || [];
    setHeaders(payload.headers || [], payload.blocked || [], payload.always || []);
    setStatus(dict.statusReady);
  } catch (err) {
    setStatus(`${dict.statusError} ${err.message}`, "error");
  }
}

function init() {
  const savedLang = localStorage.getItem(storageKeys.lang);
  const savedTheme = localStorage.getItem(storageKeys.theme);
  setTheme(savedTheme || "dark");
  setLang(savedLang || "en");
  checkApiStatus();

  const langToggle = document.querySelector(".lang-toggle");
  if (langToggle) {
    langToggle.addEventListener("click", (e) => {
      e.preventDefault();
      toggleLang();
    });
  }

  els.themeToggle.addEventListener("click", () => {
    setTheme(currentTheme === "dark" ? "light" : "dark");
  });

  els.fileInput.addEventListener("change", () => {
    const file = els.fileInput.files[0];
    if (!file) {
      els.processBtn.disabled = true;
      return;
    }
    handlePreview(file);
  });

  els.processBtn.addEventListener("click", handleProcess);
  els.downloadBtn.addEventListener("click", downloadOutput);
  if (els.selectAllBtn) {
    els.selectAllBtn.addEventListener("click", () => {
      selectedHeaders = new Set(availableHeaders.filter((h) => !blockedHeaders.includes(h)));
      renderHeadersList();
      updateProcessAvailability();
    });
  }
  if (els.clearAllBtn) {
    els.clearAllBtn.addEventListener("click", () => {
      selectedHeaders = new Set();
      renderHeadersList();
      updateProcessAvailability();
    });
  }
}

window.addEventListener("DOMContentLoaded", init);
