/* ============================================================
   محوّل — منطق الواجهة
   ============================================================ */
(() => {
  "use strict";

  const ALLOWED = ["pdf", "png", "jpg", "jpeg", "webp", "bmp", "tif", "tiff", "gif"];
  const MAX_FILES = 30;
  const MAX_TOTAL = 120 * 1024 * 1024;
  const LABEL = { docx: "Word", xlsx: "Excel", pptx: "PowerPoint" };

  const $ = (id) => document.getElementById(id);
  const dz = $("dropzone");
  const input = $("fileInput");
  const fileList = $("fileList");
  const convertBtn = $("convertBtn");
  const clearBtn = $("clearBtn");
  const progress = $("progress");
  const progressText = $("progressText");
  const results = $("results");
  const resultList = $("resultList");
  const ocrToggle = $("ocrToggle");
  const fidelityToggle = $("fidelityToggle");
  const engineLocal = $("engineLocal");
  const engineAi = $("engineAi");
  const aiKeyBox = $("aiKeyBox");
  const geminiApiKey = $("geminiApiKey");
  const toggleKeyVisibility = $("toggleKeyVisibility");
  const toastEl = $("toast");

  // استرجاع المفتاح المحفوظ
  const savedApiKey = localStorage.getItem("muhawwil-gemini-key") || "";
  if (geminiApiKey) geminiApiKey.value = savedApiKey;

  const getSelectedEngine = () => (engineAi && engineAi.checked ? "ai" : "local");

  function updateEngineUI() {
    const isAi = getSelectedEngine() === "ai";
    if (aiKeyBox) aiKeyBox.hidden = !isAi;
  }

  if (engineLocal && engineAi) {
    engineLocal.addEventListener("change", updateEngineUI);
    engineAi.addEventListener("change", updateEngineUI);
  }

  if (geminiApiKey) {
    geminiApiKey.addEventListener("input", (e) => {
      localStorage.setItem("muhawwil-gemini-key", e.target.value.trim());
    });
  }

  if (toggleKeyVisibility && geminiApiKey) {
    toggleKeyVisibility.addEventListener("click", () => {
      geminiApiKey.type = geminiApiKey.type === "password" ? "text" : "password";
    });
  }

  const state = { files: [], target: "docx", busy: false };

  /* ---------------- أدوات ---------------- */
  const extOf = (name) => (name.split(".").pop() || "").toLowerCase();
  const isPdf = (name) => extOf(name) === "pdf";

  function fmtSize(bytes) {
    if (bytes < 1024) return bytes + " بايت";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(0) + " كيلوبايت";
    return (bytes / (1024 * 1024)).toFixed(1) + " ميجابايت";
  }

  let toastTimer;
  function toast(msg, isError = false) {
    toastEl.textContent = msg;
    toastEl.classList.toggle("err", isError);
    toastEl.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toastEl.classList.remove("show"), isError ? 8000 : 3600);
  }

  const totalBytes = () => state.files.reduce((sum, f) => sum + f.size, 0);

  /* ---------------- الملفات ---------------- */
  function addFiles(list) {
    const incoming = Array.from(list || []);
    let rejected = 0;
    let duplicates = 0;

    for (const file of incoming) {
      if (!ALLOWED.includes(extOf(file.name))) {
        rejected++;
        continue;
      }
      const key = file.name + "|" + file.size;
      if (state.files.some((f) => f.name + "|" + f.size === key)) {
        duplicates++;
        continue;
      }
      if (state.files.length >= MAX_FILES) {
        toast(`الحد الأقصى ${MAX_FILES} ملفًا في المرة الواحدة.`, true);
        break;
      }
      state.files.push(file);
    }

    if (totalBytes() > MAX_TOTAL) {
      toast("الحجم الإجمالي يتجاوز 120 ميجابايت — أزل بعض الملفات.", true);
    }
    if (rejected) toast(`تم تجاهل ${rejected} ملف بصيغة غير مدعومة.`, true);
    if (duplicates) toast(`تم تجاهل ${duplicates} ملف مكرر.`);

    render();
  }

  function removeAt(index) {
    state.files.splice(index, 1);
    render();
  }

  function render() {
    fileList.hidden = state.files.length === 0;
    fileList.innerHTML = "";

    state.files.forEach((file, i) => {
      const pdf = isPdf(file.name);
      const li = document.createElement("li");
      li.className = "file-item";
      li.innerHTML = `
        <span class="f-badge ${pdf ? "pdf" : "img"}">${pdf ? "PDF" : extOf(file.name).slice(0, 4).toUpperCase()}</span>
        <span class="f-meta">
          <span class="f-name"></span>
          <span class="f-size">${fmtSize(file.size)}</span>
        </span>
        <button class="f-del" type="button" title="إزالة" aria-label="إزالة">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M18 6 6 18M6 6l12 12"/></svg>
        </button>`;
      li.querySelector(".f-name").textContent = file.name;
      li.querySelector(".f-del").addEventListener("click", () => removeAt(i));
      fileList.appendChild(li);
    });

    const ready = state.files.length > 0 && !state.busy;
    convertBtn.disabled = !ready;
    clearBtn.disabled = state.files.length === 0 || state.busy;
    convertBtn.querySelector(".btn-label").textContent =
      state.files.length > 1
        ? `تحويل الملفات (${state.files.length}) إلى ${LABEL[state.target]}`
        : `تحويل إلى ${LABEL[state.target]}`;
  }

  /* ---------------- التحويل ---------------- */
  function setBusy(on, message = "") {
    state.busy = on;
    convertBtn.classList.toggle("busy", on);
    progress.hidden = !on;
    if (on) progressText.textContent = message;
    render();
  }

  function convert() {
    if (state.busy || !state.files.length) return;
    if (totalBytes() > MAX_TOTAL) {
      toast("الحجم الإجمالي يتجاوز 120 ميجابايت.", true);
      return;
    }

    const useOcr = ocrToggle.checked;
    const engineMode = getSelectedEngine();
    const apiKeyVal = geminiApiKey ? geminiApiKey.value.trim() : "";

    const form = new FormData();
    form.append("target", state.target);
    form.append("ocr", useOcr ? "1" : "0");
    form.append("fid", fidelityToggle.checked ? "1" : "0");
    form.append("engine", engineMode);
    form.append("apiKey", apiKeyVal);
    state.files.forEach((f) => form.append("files", f, f.name));

    const slowHint = engineMode === "ai" ? " جاري المعالجة بالذكاء الاصطناعي..." : (useOcr ? " قد يستغرق الملف الممسوح دقيقة تقريبًا." : "");
    setBusy(true, "جارٍ رفع الملفات…");

    const xhr = new XMLHttpRequest();
    xhr.open("POST", "/api/convert");
    xhr.responseType = "blob";

    xhr.upload.onprogress = (e) => {
      if (!e.lengthComputable) return;
      const pct = Math.round((e.loaded / e.total) * 100);
      progressText.textContent = pct >= 100 ? "جارٍ التحويل…" : `جارٍ الرفع ${pct}%`;
    };
    xhr.upload.onload = () => {
      progressText.textContent = "جارٍ التحويل…" + slowHint;
    };

    xhr.onload = async () => {
      setBusy(false);
      if (xhr.status === 200) {
        const encoded = xhr.getResponseHeader("X-Output-Name") || "";
        let name = "";
        try { name = decodeURIComponent(encoded); } catch { name = ""; }
        if (!name) name = `محوّل_${state.target}.${state.target}`;
        showResult(xhr.response, name);
        toast("تم التحويل بنجاح ✓");
        return;
      }
      let message = "تعذّر إكمال التحويل.";
      try {
        const text = await xhr.response.text();
        const data = JSON.parse(text);
        if (data && data.detail) message = data.detail;
      } catch { /* تجاهل */ }
      showError(message);
      toast(message, true);
    };

    xhr.onerror = () => {
      setBusy(false);
      const message = "انقطع الاتصال بالخدمة — تأكد أن التطبيق يعمل.";
      showError(message);
      toast(message, true);
    };

    xhr.send(form);
  }

  function showError(message) {
    results.hidden = false;
    const li = document.createElement("li");
    li.className = "res-item error";
    li.innerHTML = `
      <span class="f-badge err">!</span>
      <span class="f-meta">
        <span class="f-name"></span>
        <span class="f-size">لم يتم إنشاء ملف — راجع السبب وحاول مرة أخرى</span>
      </span>`;
    li.querySelector(".f-name").textContent = message || "تعذّر إكمال التحويل.";
    resultList.prepend(li);
  }

  function showResult(blob, name) {
    const url = URL.createObjectURL(blob);
    const ext = extOf(name);
    const badge = ["docx", "xlsx", "pptx", "zip"].includes(ext) ? ext : "zip";

    results.hidden = false;
    const li = document.createElement("li");
    li.className = "res-item";
    li.innerHTML = `
      <span class="f-badge ${badge}">${ext.toUpperCase().slice(0, 4)}</span>
      <span class="f-meta">
        <span class="f-name"></span>
        <span class="f-size">${fmtSize(blob.size)} · جاهز للتنزيل</span>
      </span>
      <a class="dl-btn" download>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v12"/><path d="m6 11 6 6 6-6"/><path d="M5 20h14"/></svg>
        تنزيل
      </a>`;
    li.querySelector(".f-name").textContent = name;
    const link = li.querySelector(".dl-btn");
    link.href = url;
    link.download = name;
    resultList.prepend(li);

    // تنزيل تلقائي مباشر
    const auto = document.createElement("a");
    auto.href = url;
    auto.download = name;
    document.body.appendChild(auto);
    auto.click();
    auto.remove();
  }

  /* ---------------- الأحداث ---------------- */
  dz.addEventListener("click", () => input.click());
  dz.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") { e.preventDefault(); input.click(); }
  });
  input.addEventListener("change", () => { addFiles(input.files); input.value = ""; });

  ["dragenter", "dragover"].forEach((evt) =>
    dz.addEventListener(evt, (e) => { e.preventDefault(); dz.classList.add("drag"); })
  );
  ["dragleave", "dragend"].forEach((evt) =>
    dz.addEventListener(evt, () => dz.classList.remove("drag"))
  );
  dz.addEventListener("drop", (e) => {
    e.preventDefault();
    dz.classList.remove("drag");
    addFiles(e.dataTransfer.files);
  });

  // منع المتصفح من فتح الملف عند إفلاته خارج المنطقة
  window.addEventListener("dragover", (e) => e.preventDefault());
  window.addEventListener("drop", (e) => e.preventDefault());

  // لصق صور من الحافظة
  window.addEventListener("paste", (e) => {
    const items = e.clipboardData?.files;
    if (items && items.length) { addFiles(items); toast("تم لصق الملف من الحافظة."); }
  });

  document.querySelectorAll(".target").forEach((btn) => {
    btn.addEventListener("click", () => {
      document.querySelectorAll(".target").forEach((b) => {
        b.classList.remove("active");
        b.setAttribute("aria-selected", "false");
      });
      btn.classList.add("active");
      btn.setAttribute("aria-selected", "true");
      state.target = btn.dataset.target;
      render();
    });
  });

  convertBtn.addEventListener("click", convert);
  clearBtn.addEventListener("click", () => {
    state.files = [];
    resultList.innerHTML = "";
    results.hidden = true;
    render();
  });

  // المظهر
  const themeBtn = $("themeBtn");
  const savedTheme = localStorage.getItem("muhawwil-theme");
  const prefersDark = window.matchMedia?.("(prefers-color-scheme: dark)").matches;
  document.documentElement.dataset.theme = savedTheme || (prefersDark ? "dark" : "light");
  themeBtn.addEventListener("click", () => {
    const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
    document.documentElement.dataset.theme = next;
    localStorage.setItem("muhawwil-theme", next);
  });

  render();
})();
