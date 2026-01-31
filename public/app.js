const STRINGS = {
  en: {
    title: "Invoice PDF → Excel",
    subtitle: "Upload → Convert → Download",
    stepTitle: "Step 1. Upload PDF",
    stepHint: "We extract items, quantities, and prices into Excel.",
    uploadTitle: "Step 2. Download Excel",
    uploadHint: "Upload your PDF and get a structured .xlsx file.",
    processBtn: "Convert",
    downloadOutput: "Download Excel",
    statusReady: "Ready.",
    statusProcessing: "Processing...",
    statusError: "Error:",
    statusNoFile: "Please select a .pdf file.",
    statusDone: "Done. File is ready.",
  },
  ru: {
    title: "PDF счет → Excel",
    subtitle: "Загрузка → Конвертация → Скачивание",
    stepTitle: "Шаг 1. Загрузите PDF",
    stepHint: "Мы извлекаем позиции, количества и цены в Excel.",
    uploadTitle: "Шаг 2. Скачайте Excel",
    uploadHint: "Загрузите PDF и получите структурированный .xlsx файл.",
    processBtn: "Конвертировать",
    downloadOutput: "Скачать Excel",
    statusReady: "Готово.",
    statusProcessing: "Обработка...",
    statusError: "Ошибка:",
    statusNoFile: "Выберите файл .pdf.",
    statusDone: "Готово. Файл готов к скачиванию.",
  },
};

const storageKeys = {
  lang: "invoicePdfLang",
  theme: "invoicePdfTheme",
};

const els = {
  fileInput: document.getElementById("fileInput"),
  processBtn: document.getElementById("processBtn"),
  statusMsg: document.getElementById("statusMsg"),
  downloadBtn: document.getElementById("downloadBtn"),
  themeToggle: document.getElementById("themeToggle"),
};

let currentLang = "en";
let currentTheme = "dark";
let latestBlobUrl = "";
let latestFilename = "";

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

  setStatus(dict.statusProcessing);
  els.processBtn.disabled = true;
  els.downloadBtn.disabled = true;

  try {
    const formData = new FormData();
    formData.append("file", file, file.name);
    const res = await fetch("/upload", {
      method: "POST",
      body: formData,
    });
    if (!res.ok) {
      const text = await res.text();
      setStatus(`${dict.statusError} ${text || res.statusText}`, "error");
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

function init() {
  const savedLang = localStorage.getItem(storageKeys.lang);
  const savedTheme = localStorage.getItem(storageKeys.theme);
  setTheme(savedTheme || "dark");
  setLang(savedLang || "en");

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
    els.processBtn.disabled = !els.fileInput.files.length;
  });

  els.processBtn.addEventListener("click", handleProcess);
  els.downloadBtn.addEventListener("click", downloadOutput);
}

window.addEventListener("DOMContentLoaded", init);
