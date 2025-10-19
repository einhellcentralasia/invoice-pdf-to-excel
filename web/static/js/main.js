// main.js — handles upload + UI toggles + hidden debug console

document.addEventListener("DOMContentLoaded", () => {
  const form = document.getElementById("upload-form");
  const result = document.getElementById("result");
  const infoBtn = document.getElementById("info-btn");
  const infoModal = document.getElementById("info-modal");
  const closeInfo = document.getElementById("close-info");
  const langBtn = document.getElementById("lang-btn");

  // Load language preference
  let lang = localStorage.getItem("lang") || "RU";
  updateLangIcon();

  langBtn.onclick = () => {
    lang = lang === "RU" ? "EN" : "RU";
    localStorage.setItem("lang", lang);
    updateLangIcon();
  };

  function updateLangIcon() {
    langBtn.textContent = lang === "RU" ? "🇷🇺" : "🇬🇧";
  }

  infoBtn.onclick = () => infoModal.style.display = "grid";
  closeInfo.onclick = () => infoModal.style.display = "none";

  form.onsubmit = async (e) => {
    e.preventDefault();
    const fileInput = document.getElementById("pdf");
    if (!fileInput.files.length) return alert("Select a PDF first!");

    const data = new FormData();
    data.append("file", fileInput.files[0]);

    result.innerHTML = "Processing… please wait ⏳";

    const resp = await fetch("/upload", { method: "POST", body: data });
    if (!resp.ok) {
      result.innerHTML = "Error processing file.";
      return;
    }

    const blob = await resp.blob();
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileInput.files[0].name.replace(".pdf", ".xlsx");
    link.click();

    result.innerHTML = `<button class="btn" id="download-again">⬇️ Download Again</button>`;
    document.getElementById("download-again").onclick = () => link.click();
  };

  // Hidden debug toggle
  let debugVisible = false;
  document.addEventListener("keydown", (e) => {
    if (e.ctrlKey && e.shiftKey && e.key === "D") {
      debugVisible = !debugVisible;
      alert(debugVisible ? "Debug ON" : "Debug OFF");
    }
  });
});
