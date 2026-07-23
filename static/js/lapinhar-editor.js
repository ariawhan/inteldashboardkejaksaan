const editorInstances = [];
let paginationFrame;

function updateAutomaticReportNumber() {
  const field = document.querySelector('[name="report_number"]');
  const sequence = document.querySelector("[data-letter-sequence]")?.value || "—";
  const institution = document.querySelector("[data-institution-code]")?.value || "N.1.11";
  const issue = document.querySelector("[data-issue-code]")?.value || "—";
  const selectedDate = document.querySelector('[name="report_date"]')?.value;
  let month = "—";
  let year = "—";
  if (selectedDate) {
    [year, month] = selectedDate.split("-");
  }
  document.querySelectorAll('[data-output="issue-code"]').forEach((node) => { node.textContent = issue; });
  document.querySelectorAll('[data-output="month-year"]').forEach((node) => { node.textContent = `${month}/${year}`; });
  if (field) {
    const prefix = lapinharForm?.dataset.documentPrefix || "R.LIH";
    field.value = `${prefix}-${sequence}/${institution}/${issue}/${month}/${year}`;
    field.dispatchEvent(new Event("input", { bubbles: true }));
  }
}

document.querySelector("[data-issue-code]")?.addEventListener("change", updateAutomaticReportNumber);
document.querySelector('[name="report_date"]')?.addEventListener("input", updateAutomaticReportNumber);
requestAnimationFrame(updateAutomaticReportNumber);

const lapinharForm = document.querySelector("#lapinhar-form");
const numberStatus = document.querySelector("[data-number-status]");
const reloadNumberButton = document.querySelector("[data-reload-letter-number]");

function updateSipedePreviewNumber(value) {
  const number = String(value || "").trim() || "—";
  document.querySelectorAll('[data-output="sipede-number"]').forEach(node => {
    node.textContent = number === "-" ? "—" : number;
  });
}

const sipedeNumberField = document.querySelector('input[name="sipede_number"]');
updateSipedePreviewNumber(sipedeNumberField?.value);

async function checkLetterNumber() {
  const response = await fetch(lapinharForm?.dataset.checkNumberUrl || "/lapinhar/check-number", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      report_number: document.querySelector('[name="report_number"]')?.value || "",
      reservation_token: document.querySelector('[name="number_reservation_token"]')?.value || "",
      report_id: lapinharForm?.dataset.reportId || ""
    })
  });
  let result;
  try { result = await response.json(); } catch (_error) { result = { available: false, message: "Gagal memeriksa nomor surat." }; }
  if (numberStatus) {
    numberStatus.textContent = result.available
      ? (result.message || "Nomor surat tersedia.")
      : `${result.message || "Nomor surat sudah digunakan."} Tekan tombol Reload Nomor untuk mengambil nomor baru.`;
    numberStatus.classList.toggle("number-safe", Boolean(result.available));
    numberStatus.classList.toggle("number-conflict", !result.available);
  }
  if (reloadNumberButton) reloadNumberButton.hidden = Boolean(result.available);
  if (!response.ok || !result.available) throw new Error(result.message || "Nomor surat sudah digunakan.");
  return true;
}

reloadNumberButton?.addEventListener("click", async (event) => {
  const button = event.currentTarget;
  button.disabled = true;
  const oldText = button.textContent;
  button.textContent = "Memuat...";
  try {
    const selectedYear = Number.parseInt(document.querySelector('[name="report_date"]')?.value?.slice(0, 4), 10) || new Date().getFullYear();
    const response = await fetch(lapinharForm?.dataset.reloadNumberUrl || "/lapinhar/reload-number", {
      method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ document_year: selectedYear })
    });
    if (!response.ok) throw new Error("Gagal mengambil nomor surat baru.");
    const result = await response.json();
    document.querySelector('[name="number_reservation_token"]').value = result.reservation_token;
    document.querySelector("[data-letter-sequence]").value = result.sequence_number;
    document.querySelector("[data-letter-year]").value = result.document_year;
    updateAutomaticReportNumber();
    button.hidden = true;
    if (numberStatus) {
      numberStatus.textContent = `Nomor baru ${result.sequence_number} berhasil diambil. Silakan simpan kembali.`;
      numberStatus.classList.add("number-safe");
      numberStatus.classList.remove("number-conflict");
    }
  } catch (error) {
    alert(error.message);
  } finally {
    button.disabled = false;
    button.textContent = oldText;
  }
});

document.querySelector('[name="report_date"]')?.addEventListener("change", () => {
  const selectedYear = document.querySelector('[name="report_date"]')?.value?.slice(0, 4);
  const reservedYear = document.querySelector("[data-letter-year]")?.value;
  if (selectedYear && selectedYear !== reservedYear) reloadNumberButton?.click();
});

let numberValidatedSubmit = false;
lapinharForm?.addEventListener("submit", async (event) => {
  if (!event.submitter?.matches("[data-save-report]") || numberValidatedSubmit) {
    numberValidatedSubmit = false;
    return;
  }
  event.preventDefault();
  try {
    await checkLetterNumber();
    numberValidatedSubmit = true;
    lapinharForm.requestSubmit(event.submitter);
  } catch (error) {
    alert(error.message);
  }
});

const saveActionModal = document.querySelector("[data-save-action-modal]");
let pendingSaveActionButton = null;

function closeSaveActionModal() {
  if (!saveActionModal) return;
  saveActionModal.hidden = true;
  document.body.classList.remove("save-action-modal-open");
  pendingSaveActionButton = null;
}

document.querySelectorAll("[data-save-before-action]").forEach((button) => {
  button.addEventListener("click", (event) => {
    if (button.dataset.afterSaveReady === "true") return;
    event.preventDefault();
    pendingSaveActionButton = button;
    const labels = { docx: "Ekspor DOCX", print: "Cetak Preview", pdf: "Ekspor PDF", sipede: "Upload ke Sipede" };
    const action = button.dataset.saveBeforeAction || "";
    const label = saveActionModal?.querySelector("[data-save-action-label]");
    if (label) label.textContent = labels[action] || "aksi yang dipilih";
    if (saveActionModal) {
      saveActionModal.hidden = false;
      document.body.classList.add("save-action-modal-open");
      saveActionModal.querySelector("[data-save-action-confirm]")?.focus();
    }
  });
});

saveActionModal?.querySelectorAll("[data-save-action-cancel]").forEach((button) =>
  button.addEventListener("click", closeSaveActionModal));

saveActionModal?.querySelector("[data-save-action-confirm]")?.addEventListener("click", () => {
  if (!pendingSaveActionButton) return;
  const action = pendingSaveActionButton.dataset.saveBeforeAction || "";
  const actionField = lapinharForm?.querySelector('[name="after_save"]');
  if (actionField) actionField.value = action;
  saveActionModal.hidden = true;
  document.body.classList.remove("save-action-modal-open");
  const saveButton = lapinharForm?.querySelector("[data-save-report]");
  pendingSaveActionButton = null;
  if (saveButton) lapinharForm.requestSubmit(saveButton);
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && saveActionModal && !saveActionModal.hidden) closeSaveActionModal();
});

const categorySearch = document.querySelector("[data-category-search]");
const categorySelect = document.querySelector("[data-category-select]");
const categoryDropdown = document.querySelector("[data-category-dropdown]");
const categoryCombobox = document.querySelector("[data-category-combobox]");
if (categorySearch && categorySelect) {
  const categoryOptions = Array.from(categorySelect.options).slice(1).map((option) => ({
    value: option.value,
    label: option.textContent.trim()
  }));
  const closeCategories = () => {
    categoryCombobox.classList.remove("open");
    categorySearch.setAttribute("aria-expanded", "false");
  };
  const renderCategories = (showAll = false) => {
    const query = showAll ? "" : categorySearch.value.trim().toLocaleLowerCase("id-ID");
    const matches = categoryOptions.filter((option) => option.label.toLocaleLowerCase("id-ID").includes(query));
    categoryDropdown.innerHTML = "";
    matches.forEach((item) => {
      const option = document.createElement("button");
      option.type = "button";
      option.setAttribute("role", "option");
      option.dataset.value = item.value;
      option.textContent = item.label;
      option.addEventListener("mousedown", (event) => event.preventDefault());
      option.addEventListener("click", () => {
        categorySelect.value = item.value;
        categorySearch.value = item.label;
        categorySearch.setCustomValidity("");
        closeCategories();
      });
      categoryDropdown.appendChild(option);
    });
    if (!matches.length) categoryDropdown.innerHTML = '<em>Kategori tidak ditemukan</em>';
  };
  const openCategories = (showAll = false) => {
    renderCategories(showAll);
    categoryCombobox.classList.add("open");
    categorySearch.setAttribute("aria-expanded", "true");
  };
  categorySearch.addEventListener("focus", openCategories);
  categorySearch.addEventListener("click", openCategories);
  categorySearch.addEventListener("input", () => {
    categorySelect.value = "";
    categorySearch.setCustomValidity("Pilih kategori dari daftar.");
    openCategories();
  });
  categorySearch.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeCategories();
    if (event.key === "Enter" && categoryCombobox.classList.contains("open")) {
      const first = categoryDropdown.querySelector("button");
      if (first) { event.preventDefault(); first.click(); }
    }
  });
  document.querySelector("[data-category-toggle]")?.addEventListener("click", () => {
    if (categoryCombobox.classList.contains("open")) closeCategories();
    else { categorySearch.focus(); openCategories(true); }
  });
  document.addEventListener("click", (event) => {
    if (!categoryCombobox.contains(event.target)) closeCategories();
  });
  categorySearch.setCustomValidity("Pilih kategori dari daftar.");
  const currentCategory = categorySelect.dataset.currentValue;
  const selectedCategory = categoryOptions.find((option) => option.value === currentCategory);
  if (selectedCategory) {
    categorySelect.value = selectedCategory.value;
    categorySearch.value = selectedCategory.label;
    categorySearch.setCustomValidity("");
  }
}

const issueSearch = document.querySelector("[data-issue-search]");
const issueSelect = document.querySelector("[data-issue-code]");
const issueDropdown = document.querySelector("[data-issue-dropdown]");
const issueCombobox = document.querySelector("[data-issue-combobox]");
if (issueSearch && issueSelect && issueDropdown && issueCombobox) {
  const issueOptions = Array.from(issueSelect.options).slice(1).map((option) => ({
    value: option.value,
    label: option.textContent.trim()
  }));
  const closeIssues = () => {
    issueCombobox.classList.remove("open");
    issueSearch.setAttribute("aria-expanded", "false");
  };
  const renderIssues = (showAll = false) => {
    const query = showAll ? "" : issueSearch.value.trim().toLocaleLowerCase("id-ID");
    const matches = issueOptions.filter((option) => option.label.toLocaleLowerCase("id-ID").includes(query));
    issueDropdown.innerHTML = "";
    matches.forEach((item) => {
      const option = document.createElement("button");
      option.type = "button";
      option.setAttribute("role", "option");
      option.textContent = item.label;
      option.addEventListener("mousedown", (event) => event.preventDefault());
      option.addEventListener("click", () => {
        issueSelect.value = item.value;
        issueSearch.value = item.label;
        issueSearch.setCustomValidity("");
        issueSelect.dispatchEvent(new Event("change", { bubbles: true }));
        closeIssues();
      });
      issueDropdown.appendChild(option);
    });
    if (!matches.length) issueDropdown.innerHTML = '<em>Nomor permasalahan tidak ditemukan</em>';
  };
  const openIssues = (showAll = false) => {
    renderIssues(showAll);
    issueCombobox.classList.add("open");
    issueSearch.setAttribute("aria-expanded", "true");
  };
  issueSearch.addEventListener("focus", openIssues);
  issueSearch.addEventListener("click", openIssues);
  issueSearch.addEventListener("input", () => {
    issueSelect.value = "";
    issueSearch.setCustomValidity("Pilih nomor permasalahan dari daftar.");
    updateAutomaticReportNumber();
    openIssues();
  });
  issueSearch.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeIssues();
    if (event.key === "Enter" && issueCombobox.classList.contains("open")) {
      const first = issueDropdown.querySelector("button");
      if (first) { event.preventDefault(); first.click(); }
    }
  });
  document.querySelector("[data-issue-toggle]")?.addEventListener("click", () => {
    if (issueCombobox.classList.contains("open")) closeIssues();
    else { issueSearch.focus(); openIssues(true); }
  });
  document.addEventListener("click", (event) => {
    if (!issueCombobox.contains(event.target)) closeIssues();
  });
  issueSearch.setCustomValidity("Pilih nomor permasalahan dari daftar.");
  const currentIssue = issueSelect.dataset.currentValue;
  const selectedIssue = issueOptions.find((option) => option.value === currentIssue);
  if (selectedIssue) {
    issueSelect.value = selectedIssue.value;
    issueSearch.value = selectedIssue.label;
    issueSearch.setCustomValidity("");
  }
}

document.querySelectorAll("[data-rich-editor]").forEach((container) => {
  const name = container.dataset.richEditor;
  const hiddenInput = document.querySelector(`[data-rich-input="${name}"]`);
  const preview = document.querySelector(`[data-output-list="${name}"]`);
  const isLarge = container.classList.contains("large");

  const editor = Jodit.make(container, {
    height: isLarge ? 320 : 220,
    minHeight: isLarge ? 260 : 170,
    maxHeight: 520,
    placeholder: container.dataset.placeholder,
    toolbarAdaptive: true,
    toolbarSticky: false,
    statusbar: true,
    showCharsCounter: true,
    showWordsCounter: true,
    showXPathInStatusbar: false,
    askBeforePasteHTML: false,
    askBeforePasteFromWord: false,
    defaultActionOnPaste: "insert_clear_html",
    indentMargin: 30,
    buttons: [
      "undo", "redo", "|", "bold", "italic", "underline", "fontsize", "|",
      "ol", "ul", "|", "outdent", "indent", "|",
      "left", "center", "right", "justify", "|", "eraser"
    ]
  });

  editorInstances.push(editor);
  if (hiddenInput.value) editor.value = hiddenInput.value;
  const syncContent = () => {
    const html = editor.value.trim();
    hiddenInput.value = html;
    preview.innerHTML = html || "<p>Isi bagian ini akan tampil di sini.</p>";
    scheduleReportPagination();
  };
  editor.events.on("change", syncContent);
  syncContent();
  editor.editor.style.fontFamily = '"Times New Roman", Times, serif';
  editor.editor.style.fontSize = "16px";
  editor.editor.style.lineHeight = "1.5";
  preview.style.fontFamily = '"Times New Roman", Times, serif';
  preview.style.lineHeight = "1.5";
});

const fixedSource = "Intelijen Kejaksaan Negeri Buleleng";
const sourceContainer = document.querySelector("[data-additional-sources]");
const sourceHidden = document.querySelector("[data-source-hidden]");

function syncSources() {
  if (!sourceHidden) return;
  const extraSources = Array.from(sourceContainer.querySelectorAll("input"))
    .map((input) => input.value.trim()).filter(Boolean);
  if (!extraSources.length) {
    sourceHidden.value = `<p>${fixedSource}</p>`;
  } else {
    sourceHidden.value = `<ol><li>${fixedSource}</li>${extraSources.map((source) => `<li>${escapeSource(source)}</li>`).join("")}</ol>`;
  }
  const preview = document.querySelector('[data-output-list="sources"]');
  if (preview) preview.innerHTML = sourceHidden.value;
  scheduleReportPagination();
}

function escapeSource(value) {
  const element = document.createElement("div");
  element.textContent = value;
  return element.innerHTML;
}

function addSourceInput(value = "", focus = true) {
  const row = document.createElement("div");
  row.className = "additional-source-row";
  row.innerHTML = '<input type="text" placeholder="Masukkan sumber tambahan"><button type="button" aria-label="Hapus sumber">×</button>';
  row.querySelector("input").addEventListener("input", syncSources);
  row.querySelector("button").addEventListener("click", () => { row.remove(); syncSources(); });
  row.querySelector("input").value = value;
  sourceContainer.appendChild(row);
  if (focus) row.querySelector("input").focus();
  syncSources();
}

document.querySelector("[data-add-source]")?.addEventListener("click", () => addSourceInput());
try {
  const existingSources = JSON.parse(document.querySelector("[data-existing-sources]")?.textContent || "[]");
  existingSources.forEach((source) => addSourceInput(source, false));
} catch (_error) {}
syncSources();

const attachmentInputList = document.querySelector("[data-attachment-input-list]");
const attachmentPreviewGroup = document.querySelector("[data-attachment-preview-group]");
let attachmentSequence = 0;

function attachmentPageMarkup(id, title, imageUrls) {
  const number = document.querySelector('[name="report_number"]')?.value || "—";
  const dateInput = document.querySelector('[name="report_date"]')?.value;
  const date = dateInput ? new Intl.DateTimeFormat("id-ID", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" }).format(new Date(`${dateInput}T00:00:00Z`)) : "—";
  const images = imageUrls.length
    ? imageUrls.map((url) => `<figure><img src="${url}" alt="Foto lampiran"><figcaption>${escapeSource(title || "Dokumentasi")}</figcaption></figure>`).join("")
    : '<div class="attachment-empty"><span>▧</span><strong>Belum ada foto</strong><small>Pilih maksimal dua foto pada form lampiran.</small></div>';
  return `<article class="paper-preview folio-page attachment-preview" data-attachment-page="${id}">
    <div class="classification-header">RAHASIA</div>
    <div class="attachment-heading"><p><b>LAMPIRAN ${escapeSource(lapinharForm?.dataset.reportHeading || "LAPORAN INFORMASI HARIAN")}</b></p><p>NOMOR : <span data-output="number">${escapeSource(number)}</span></p><p>TANGGAL : <span data-output="date">${date}</span></p></div>
    <div class="attachment-photo-grid count-${Math.max(imageUrls.length, 1)}">${images}</div>
    <div class="attachment-caption"><b>KETERANGAN:</b><p>${escapeSource(title || "—")}</p></div>
    <footer class="classification-footer">RAHASIA</footer>
  </article>`;
}

function syncAttachment(row) {
  const input = row.querySelector('input[type="file"]');
  const selected = Array.from(input.files || []).slice(0, 2);
  const title = row.querySelector('input[type="text"]').value.trim();
  const urls = selected.map((file) => URL.createObjectURL(file));
  const oldPage = attachmentPreviewGroup.querySelector(`[data-attachment-page="${row.dataset.attachmentId}"]`);
  const holder = document.createElement("div");
  holder.innerHTML = attachmentPageMarkup(row.dataset.attachmentId, title, urls);
  if (oldPage) oldPage.replaceWith(holder.firstElementChild); else attachmentPreviewGroup.appendChild(holder.firstElementChild);
  row.querySelector(".selected-file-count").textContent = `${selected.length}/2 foto dipilih`;
}

function addAttachment() {
  attachmentSequence += 1;
  const id = attachmentSequence;
  const row = document.createElement("section");
  row.className = "attachment-input-card";
  row.dataset.attachmentId = id;
  row.innerHTML = `<div class="attachment-card-head"><strong>Lampiran ${id}</strong><button type="button" aria-label="Hapus lampiran">×</button></div>
    <label>Judul / keterangan<input type="text" name="attachment_title_${id}" placeholder="Keterangan dokumentasi"></label>
    <label class="photo-picker"><input type="file" name="attachment_images_${id}" accept="image/jpeg,image/png,image/webp" multiple><span>Pilih foto</span><small class="selected-file-count">0/2 foto dipilih</small></label>`;
  const fileInput = row.querySelector('input[type="file"]');
  fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 2) {
      alert("Setiap lampiran maksimal dua foto. Silakan pilih ulang.");
      fileInput.value = "";
    }
    syncAttachment(row);
  });
  row.querySelector('input[type="text"]').addEventListener("input", () => syncAttachment(row));
  row.querySelector(".attachment-card-head button").addEventListener("click", () => {
    if (attachmentInputList.children.length === 1) return;
    attachmentPreviewGroup.querySelector(`[data-attachment-page="${id}"]`)?.remove();
    row.remove();
  });
  attachmentInputList.appendChild(row);
  syncAttachment(row);
}

function loadAttachmentItem(file, index) {
  return new Promise((resolve) => {
    const url = URL.createObjectURL(file);
    const image = new window.Image();
    image.onload = () => resolve({ file, url, index, portrait: image.naturalHeight > image.naturalWidth, scale: 100 });
    image.onerror = () => resolve({ file, url, index, portrait: false, scale: 100 });
    image.src = url;
  });
}

function resolvedAttachmentLayout(row) {
  const selected = row.querySelector(".attachment-layout").value;
  if (selected !== "auto") return selected;
  return row._attachmentItems?.length === 2 && row._attachmentItems.every((item) => item.portrait) ? "side" : "stack";
}

function updateAutomaticAttachmentCount() {
  const count = Array.from(attachmentInputList.querySelectorAll(".attachment-input-card"))
    .filter((row) => (row._attachmentItems || []).length > 0).length;
  const words = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan", "sepuluh"];
  const value = count ? `${count} (${words[count] || count}) Lembar` : "-";
  const field = document.querySelector("[data-automatic-attachment]");
  if (field) field.value = value;
  document.querySelectorAll('[data-output="attachment"]').forEach((node) => { node.textContent = value; });
}

function renderResponsiveAttachment(row) {
  const id = row.dataset.attachmentId;
  const items = row._attachmentItems || [];
  const oldPage = attachmentPreviewGroup.querySelector(`[data-attachment-page="${id}"]`);
  if (!items.length) {
    oldPage?.remove();
    row.querySelector(".selected-file-count").textContent = "0/2 foto dipilih";
    updateAutomaticAttachmentCount();
    return;
  }
  const layout = resolvedAttachmentLayout(row);
  const number = document.querySelector('[name="report_number"]')?.value || "—";
  const dateInput = document.querySelector('[name="report_date"]')?.value;
  const date = dateInput ? new Intl.DateTimeFormat("id-ID", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" }).format(new Date(`${dateInput}T00:00:00Z`)) : "—";
  const photos = items.length ? items.map((item) => `<figure class="photo-frame ${item.portrait ? "portrait" : "landscape"}"><div><img src="${item.url}" alt="Foto lampiran" style="width:${item.scale}%;height:${item.scale}%"></div></figure>`).join("") : '<div class="attachment-empty"><span>▧</span><strong>Belum ada foto</strong><small>Pilih maksimal dua foto.</small></div>';
  const markup = `<article class="paper-preview folio-page attachment-preview" data-attachment-page="${id}"><div class="classification-header">RAHASIA</div><div class="attachment-heading"><p><b>LAMPIRAN ${escapeSource(lapinharForm?.dataset.reportHeading || "LAPORAN INFORMASI HARIAN")}</b></p><p>NOMOR : <span data-output="number">${escapeSource(number)}</span></p><p>TANGGAL : <span data-output="date">${date}</span></p></div><div class="attachment-photo-grid layout-${layout} count-${Math.max(items.length, 1)}">${photos}</div><footer class="classification-footer">RAHASIA</footer></article>`;
  const holder = document.createElement("div");
  holder.innerHTML = markup;
  if (oldPage) oldPage.replaceWith(holder.firstElementChild); else attachmentPreviewGroup.appendChild(holder.firstElementChild);
  row.querySelector(".selected-file-count").textContent = `${items.length}/2 foto dipilih`;
  updateAutomaticAttachmentCount();
}

function buildPhotoControls(row) {
  const controls = row.querySelector(".photo-size-controls");
  controls.innerHTML = "";
  (row._attachmentItems || []).forEach((item, index) => {
    const control = document.createElement("label");
    control.className = "photo-size-control";
    control.innerHTML = `<span>Ukuran foto ${index + 1}</span><input type="range" name="attachment_scale_${row.dataset.attachmentId}_${index + 1}" min="40" max="100" value="100"><output>100%</output>`;
    const range = control.querySelector("input");
    range.addEventListener("input", () => { item.scale = Number(range.value); control.querySelector("output").value = `${range.value}%`; renderResponsiveAttachment(row); });
    controls.appendChild(control);
  });
}

async function addResponsiveAttachment(existing = null) {
  const requestedId = Number(existing?.group || 0);
  const id = requestedId || attachmentSequence + 1;
  attachmentSequence = Math.max(attachmentSequence, id);
  const row = document.createElement("section");
  row.className = "attachment-input-card";
  row.dataset.attachmentId = id;
  row._attachmentItems = [];
  row.innerHTML = `<div class="attachment-card-head"><strong>Lampiran ${id}</strong><button type="button" aria-label="Hapus lampiran">×</button></div><label class="layout-picker">Tata letak<select class="attachment-layout" name="attachment_layout_${id}"><option value="auto">Otomatis</option><option value="stack">Atas – Bawah</option><option value="side">Berdampingan</option></select></label><div class="photo-upload-slots"><label class="photo-picker"><input type="file" name="attachment_images_${id}" accept="image/jpeg,image/png,image/webp"><span>Pilih Foto 1</span><small>Belum ada foto</small></label><label class="photo-picker"><input type="file" name="attachment_images_${id}" accept="image/jpeg,image/png,image/webp"><span>Pilih Foto 2 (opsional)</span><small>Belum ada foto</small></label></div><small class="selected-file-count">0/2 foto dipilih</small><div class="photo-size-controls"></div>`;
  const fileInputs = Array.from(row.querySelectorAll('input[type="file"]'));
  const refreshPhotos = async () => {
    const chosen = fileInputs.map((input, index) => ({ input, index, file: input.files?.[0] })).filter((entry) => entry.file);
    row._attachmentItems.forEach((item) => URL.revokeObjectURL(item.url));
    row._attachmentItems = await Promise.all(chosen.map((entry) => loadAttachmentItem(entry.file, entry.index)));
    fileInputs.forEach((input) => {
      const fileName = input.files?.[0]?.name || "Belum ada foto";
      const fileLabel = input.closest(".photo-picker").querySelector("small");
      fileLabel.textContent = fileName;
      fileLabel.title = fileName;
    });
    buildPhotoControls(row);
    renderResponsiveAttachment(row);
  };
  fileInputs.forEach((input) => input.addEventListener("change", refreshPhotos));
  row.querySelector(".attachment-layout").addEventListener("change", () => renderResponsiveAttachment(row));
  row.querySelector(".attachment-card-head button").addEventListener("click", () => { if (attachmentInputList.children.length === 1) return; attachmentPreviewGroup.querySelector(`[data-attachment-page="${id}"]`)?.remove(); row.remove(); updateAutomaticAttachmentCount(); });
  attachmentInputList.appendChild(row);
  if (existing?.images?.length) {
    row._attachmentItems = await Promise.all(existing.images.slice(0, 2).map((imageData, index) => new Promise((resolve) => {
      const image = new window.Image();
      image.onload = () => resolve({ url: imageData.url, index, portrait: image.naturalHeight > image.naturalWidth, scale: 100, existing: true });
      image.onerror = () => resolve({ url: imageData.url, index, portrait: false, scale: 100, existing: true });
      image.src = imageData.url;
    })));
    existing.images.slice(0, 2).forEach((imageData, index) => {
      const fileName = imageData.filename || `Foto tersimpan ${index + 1}`;
      const fileLabel = fileInputs[index].closest(".photo-picker").querySelector("small");
      fileLabel.textContent = fileName;
      fileLabel.title = fileName;
    });
    buildPhotoControls(row);
  }
  renderResponsiveAttachment(row);
}

document.querySelector("[data-add-attachment]")?.addEventListener("click", addResponsiveAttachment);
let existingAttachmentGroups = [];
try {
  existingAttachmentGroups = JSON.parse(document.querySelector("[data-existing-attachments]")?.textContent || "[]");
} catch (_error) {}
const attachmentInitializationPromise = existingAttachmentGroups.length
  ? Promise.all(existingAttachmentGroups.map((group) => addResponsiveAttachment(group)))
  : addResponsiveAttachment();

function refreshAttachmentMetadata() {
  attachmentInputList.querySelectorAll(".attachment-input-card").forEach((row) => renderResponsiveAttachment(row));
}

document.querySelector('[name="report_number"]')?.addEventListener("input", refreshAttachmentMetadata);
document.querySelector('[name="report_date"]')?.addEventListener("input", refreshAttachmentMetadata);
document.querySelector('[name="report_date"]')?.addEventListener("change", refreshAttachmentMetadata);

function scheduleReportPagination() {
  cancelAnimationFrame(paginationFrame);
  paginationFrame = requestAnimationFrame(() => requestAnimationFrame(paginateReportPreview));
}
window.scheduleReportPagination = scheduleReportPagination;

function reportContentSections() {
  const definitions = [
    ["I.   INFORMASI YANG DIPEROLEH", "information"], ["II.  SUMBER INFORMASI", "sources"],
    ["III. TREN PERKEMBANGAN / PERKIRAAN", "trends"], ["IV.  PENDAPAT / SARAN", "suggestions"]
  ];
  return definitions.map(([title, name]) => {
    const holder = document.createElement("div");
    holder.innerHTML = document.querySelector(`[data-rich-input="${name}"]`)?.value || "";
    const blocks = [];
    Array.from(holder.children).forEach((element) => {
      if (["OL", "UL"].includes(element.tagName)) {
        const originalStart = Number.parseInt(element.getAttribute("start") || "1", 10);
        Array.from(element.children).forEach((item, itemIndex) => {
          const list = document.createElement(element.tagName.toLowerCase());
          Array.from(element.attributes).forEach((attribute) => list.setAttribute(attribute.name, attribute.value));
          if (element.tagName === "OL") list.setAttribute("start", originalStart + itemIndex);
          list.appendChild(item.cloneNode(true));
          blocks.push(list);
        });
      } else {
        blocks.push(element.cloneNode(true));
      }
    });
    return { title, blocks };
  });
}

function paginateReportPreview() {
  const group = document.querySelector("[data-report-preview-group]");
  if (!group) return;
  const measurementMode = !group.classList.contains("active");
  if (measurementMode) group.classList.add("pagination-measuring");
  const existingPages = group.querySelectorAll("[data-report-page]");
  if (!existingPages.length) {
    if (measurementMode) group.classList.remove("pagination-measuring");
    return;
  }
  const firstPage = existingPages[0];
  if (!firstPage.clientHeight) {
    if (measurementMode) group.classList.remove("pagination-measuring");
    return;
  }
  const originalSignature = group.querySelector(".signature-grid")?.cloneNode(true);
  existingPages.forEach((page, index) => { if (index) page.remove(); });
  const body = firstPage.querySelector(".official-body");
  body.innerHTML = "";
  firstPage.querySelector(".signature-grid")?.remove();
  const skeleton = firstPage.cloneNode(true);
  const pages = [firstPage];

  const addPage = () => {
    const page = skeleton.cloneNode(true);
    page.classList.add("continuation-page");
    page.querySelector(".report-code")?.remove();
    page.querySelector(".document-header")?.remove();
    group.appendChild(page);
    pages.push(page);
    return page;
  };
  const overflows = (page) => page.scrollHeight > page.clientHeight + 2;
  let currentPage = firstPage;

  reportContentSections().forEach((sectionData) => {
    if (!sectionData.blocks.length) return;
    let section = document.createElement("section");
    section.className = "paginated-section";
    const heading = document.createElement("h3");
    heading.textContent = sectionData.title;
    section.appendChild(heading);
    currentPage.querySelector(".official-body").appendChild(section);

    sectionData.blocks.forEach((block, blockIndex) => {
      section.appendChild(block);
      if (!overflows(currentPage)) return;
      section.removeChild(block);
      if (!section.querySelector("p,ol,ul")) section.remove();
      currentPage = addPage();
      section = document.createElement("section");
      section.className = "paginated-section";
      if (blockIndex === 0) {
        const movedHeading = document.createElement("h3");
        movedHeading.textContent = sectionData.title;
        section.appendChild(movedHeading);
      }
      section.appendChild(block);
      currentPage.querySelector(".official-body").appendChild(section);
    });
  });

  if (originalSignature) {
    currentPage.appendChild(originalSignature);
    if (overflows(currentPage)) {
      originalSignature.remove();
      currentPage = addPage();
      currentPage.appendChild(originalSignature);
    }
  }
  pages.forEach((page, index) => page.dataset.pageNumber = index + 1);
  if (measurementMode) group.classList.remove("pagination-measuring");
}

async function preparePreviewForOutput() {
  await attachmentInitializationPromise;
  if (document.fonts?.ready) await document.fonts.ready;
  const pendingImages = Array.from(document.querySelectorAll(".preview-pages img")).filter((image) => !image.complete);
  await Promise.all(pendingImages.map((image) => new Promise((resolve) => {
    image.addEventListener("load", resolve, { once: true });
    image.addEventListener("error", resolve, { once: true });
  })));
  paginateReportPreview();
  await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
}

attachmentInitializationPromise.then(() => {
  if (document.fonts?.ready) document.fonts.ready.then(scheduleReportPagination);
  else scheduleReportPagination();
});

document.querySelector("#lapinhar-form")?.addEventListener("submit", (event) => {
  if (issueSelect && !issueSelect.value) {
    event.preventDefault();
    issueSearch.reportValidity();
    issueSearch.focus();
    return;
  }
  if (categorySelect && !categorySelect.value) {
    event.preventDefault();
    categorySearch.reportValidity();
    categorySearch.focus();
    return;
  }
  const information = document.querySelector('[data-rich-input="information"]');
  if (!information?.value) {
    event.preventDefault();
    document.querySelector('[data-rich-editor="information"]')?.scrollIntoView({ behavior: "smooth", block: "center" });
    editorInstances[0]?.focus();
  }
});

document.querySelectorAll('[name="subsection_signer"]').forEach((choice) => {
  choice.addEventListener("change", () => {
    if (!choice.checked) return;
    const noAuthentication = choice.dataset.signerNoAuth === "true";
    document.querySelectorAll(".signature-grid").forEach((signature) => signature.classList.toggle("auth-hidden", noAuthentication));
    document.querySelectorAll('[data-output="creator_position"]').forEach((item) => item.textContent = choice.dataset.signerPosition || "—");
    document.querySelectorAll('[data-output="creator_name"]').forEach((item) => item.textContent = choice.dataset.signerName || "Belum dikonfigurasi");
    document.querySelectorAll('[data-output="creator_rank"]').forEach((item) => item.textContent = choice.dataset.signerRank || "—");
    scheduleReportPagination();
  });
});

// Saat membuka halaman edit, terapkan pilihan yang sudah tersimpan ke pratinjau.
// Sebelumnya pratinjau tetap memakai data Kasubsi I sampai radio diklik ulang.
document.querySelector('[name="subsection_signer"]:checked')?.dispatchEvent(
  new Event("change", { bubbles: true })
);

async function printLivePreview() {
  try {
    await checkLetterNumber();
    await preparePreviewForOutput();
    document.body.classList.add("printing-preview");
    requestAnimationFrame(() => requestAnimationFrame(() => window.print()));
  } catch (error) {
    alert(error.message);
  }
}

document.querySelectorAll("[data-print-preview]").forEach((button) => {
  button.addEventListener("click", (event) => {
    if (event.currentTarget.dataset.afterSaveReady !== "true") return;
    printLivePreview();
  });
});

function imageAsDataUrl(image) {
  if (image.src.startsWith("data:")) return Promise.resolve(image.src);
  return new Promise((resolve, reject) => {
    const source = new Image();
    source.onload = () => {
      const maximum = 2000;
      const scale = Math.min(1, maximum / Math.max(source.naturalWidth, source.naturalHeight));
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(source.naturalWidth * scale));
      canvas.height = Math.max(1, Math.round(source.naturalHeight * scale));
      canvas.getContext("2d").drawImage(source, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL("image/jpeg", .92));
    };
    source.onerror = reject;
    source.src = image.src;
  });
}

async function exportPreviewPdf(button) {
  const originalText = button.textContent;
  button.disabled = true;
  button.textContent = "Membuat PDF…";
  try {
    await checkLetterNumber();
    await preparePreviewForOutput();
    const source = document.querySelector(".preview-pages");
    const clone = source.cloneNode(true);
    const sourceImages = Array.from(source.querySelectorAll("img"));
    const cloneImages = Array.from(clone.querySelectorAll("img"));
    const imageData = await Promise.all(sourceImages.map(imageAsDataUrl));
    cloneImages.forEach((image, index) => { image.src = imageData[index]; });

    const response = await fetch("/lapinhar/export-preview-pdf", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        html: clone.innerHTML,
        report_number: document.querySelector('[name="report_number"]')?.value || "lapinhar",
        report_date: document.querySelector('[name="report_date"]')?.value || "",
        subject: document.querySelector('[name="subject"]')?.value || "",
        reservation_token: document.querySelector('[name="number_reservation_token"]')?.value || "",
        report_id: lapinharForm?.dataset.reportId || "",
        document_kind: lapinharForm?.dataset.documentKind || "lapinhar"
      })
    });
    if (!response.ok) throw new Error(await response.text() || "Ekspor PDF gagal.");
    const blob = await response.blob();
    const disposition = response.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename\*?=(?:UTF-8''|\")?([^\";]+)/i);
    const filename = match ? decodeURIComponent(match[1].replace(/\"/g, "")) : "LAPINHAR.pdf";
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  } catch (error) {
    alert(error.message || "Gagal membuat PDF dari pratinjau.");
  } finally {
    button.disabled = false;
    button.textContent = originalText;
  }
}

document.querySelector("[data-export-preview-pdf]")?.addEventListener("click", (event) => {
  if (event.currentTarget.dataset.afterSaveReady !== "true") return;
  exportPreviewPdf(event.currentTarget);
});

const sipedeConfirmModal = document.querySelector("[data-sipede-confirm-modal]");
let resolveSipedeConfirmation = null;

function closeSipedeConfirmation(confirmed = false) {
  if (!sipedeConfirmModal || sipedeConfirmModal.hidden) return;
  sipedeConfirmModal.hidden = true;
  document.body.classList.remove("sipede-confirm-open");
  const resolve = resolveSipedeConfirmation;
  resolveSipedeConfirmation = null;
  resolve?.(confirmed);
}

function confirmSipedeUpload() {
  if (!sipedeConfirmModal) return Promise.resolve(true);
  sipedeConfirmModal.hidden = false;
  document.body.classList.add("sipede-confirm-open");
  sipedeConfirmModal.querySelector("[data-sipede-confirm-submit]")?.focus();
  return new Promise((resolve) => { resolveSipedeConfirmation = resolve; });
}

sipedeConfirmModal?.querySelectorAll("[data-sipede-confirm-cancel]").forEach((button) =>
  button.addEventListener("click", () => closeSipedeConfirmation(false)));
sipedeConfirmModal?.querySelector("[data-sipede-confirm-submit]")?.addEventListener("click", () =>
  closeSipedeConfirmation(true));

const sipedeDestinationModal = document.querySelector("[data-sipede-destination-modal]");
const sipedeDestinationList = sipedeDestinationModal?.querySelector("[data-sipede-destination-list]");
const sipedeDestinationSearch = sipedeDestinationModal?.querySelector("[data-sipede-destination-search]");
const sipedeDestinationAll = sipedeDestinationModal?.querySelector("[data-sipede-destination-all]");
const sipedeDestinationCount = sipedeDestinationModal?.querySelector("[data-sipede-destination-count]");
const sipedeDestinationSend = sipedeDestinationModal?.querySelector("[data-sipede-destination-send]");
let sipedeDestinations = [];
let sipedeUploadButton = null;

function selectedSipedeDestinationIds() {
  return Array.from(sipedeDestinationList?.querySelectorAll('input[type="checkbox"]:checked') || [])
    .map(input => input.value);
}

function updateSipedeDestinationSelection() {
  const selected = selectedSipedeDestinationIds();
  if (sipedeDestinationCount) sipedeDestinationCount.textContent = `${selected.length} tujuan dipilih`;
  if (sipedeDestinationSend) sipedeDestinationSend.disabled = selected.length === 0;
  const visible = Array.from(sipedeDestinationList?.querySelectorAll('.sipede-destination-item:not([hidden]) input') || []);
  if (sipedeDestinationAll) {
    sipedeDestinationAll.checked = visible.length > 0 && visible.every(input => input.checked);
    sipedeDestinationAll.indeterminate = visible.some(input => input.checked) && !sipedeDestinationAll.checked;
  }
}

function filterSipedeDestinations() {
  const query = (sipedeDestinationSearch?.value || "").trim().toLocaleLowerCase("id");
  sipedeDestinationList?.querySelectorAll(".sipede-destination-item").forEach(item => {
    item.hidden = Boolean(query) && !item.dataset.search.includes(query);
  });
  updateSipedeDestinationSelection();
}

function openSipedeDestinationModal(destinations, button) {
  if (!sipedeDestinationModal || !sipedeDestinationList) return;
  sipedeDestinations = Array.isArray(destinations) ? destinations : [];
  sipedeUploadButton = button;
  sipedeDestinationList.replaceChildren();
  sipedeDestinations.forEach(destination => {
    const label = document.createElement("label");
    label.className = "sipede-destination-item";
    label.dataset.search = `${destination.position || ""} ${destination.user || ""}`.toLocaleLowerCase("id");
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.value = destination.id;
    checkbox.addEventListener("change", updateSipedeDestinationSelection);
    const text = document.createElement("span");
    const position = document.createElement("strong");
    position.textContent = destination.position || "Tujuan SIPede";
    const user = document.createElement("small");
    user.textContent = destination.user || "NO USER";
    text.append(position, user);
    label.append(checkbox, text);
    sipedeDestinationList.append(label);
  });
  if (!sipedeDestinations.length) {
    const empty = document.createElement("p");
    empty.className = "sipede-destination-empty";
    empty.textContent = "Tujuan SIPede belum tersedia.";
    sipedeDestinationList.append(empty);
  }
  if (sipedeDestinationSearch) sipedeDestinationSearch.value = "";
  if (sipedeDestinationAll) sipedeDestinationAll.checked = false;
  updateSipedeDestinationSelection();
  sipedeDestinationModal.hidden = false;
  document.body.classList.add("sipede-confirm-open");
  sipedeDestinationSearch?.focus();
}

function closeSipedeDestinationModal() {
  if (!sipedeDestinationModal) return;
  sipedeDestinationModal.hidden = true;
  document.body.classList.remove("sipede-confirm-open");
  sipedeUploadButton = null;
}

async function createSipedePdf() {
  await preparePreviewForOutput();
  const source = document.querySelector(".preview-pages");
  const clone = source.cloneNode(true);
  const sourceImages = Array.from(source.querySelectorAll("img"));
  const cloneImages = Array.from(clone.querySelectorAll("img"));
  const imageData = await Promise.all(sourceImages.map(imageAsDataUrl));
  cloneImages.forEach((image, index) => { image.src = imageData[index]; });
  const response = await fetch("/lapinhar/export-preview-pdf", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      html: clone.innerHTML,
      report_number: document.querySelector('[name="report_number"]')?.value || "lapinsus",
      report_date: document.querySelector('[name="report_date"]')?.value || "",
      subject: document.querySelector('[name="subject"]')?.value || "",
      reservation_token: document.querySelector('[name="number_reservation_token"]')?.value || "",
      report_id: lapinharForm?.dataset.reportId || "",
      document_kind: "lapinsus"
    })
  });
  if (!response.ok) throw new Error(await response.text() || "PDF LAPINSUS gagal dibuat.");
  const disposition = response.headers.get("Content-Disposition") || "";
  const match = disposition.match(/filename\*?=(?:UTF-8''|\")?([^\";]+)/i);
  return {
    blob: await response.blob(),
    filename: match ? decodeURIComponent(match[1].replace(/\"/g, "")) : "LAPINSUS.pdf"
  };
}

sipedeDestinationSearch?.addEventListener("input", filterSipedeDestinations);
sipedeDestinationAll?.addEventListener("change", () => {
  sipedeDestinationList?.querySelectorAll('.sipede-destination-item:not([hidden]) input').forEach(input => {
    input.checked = sipedeDestinationAll.checked;
  });
  updateSipedeDestinationSelection();
});
sipedeDestinationModal?.querySelectorAll("[data-sipede-destination-close]").forEach(button =>
  button.addEventListener("click", closeSipedeDestinationModal));
sipedeDestinationSend?.addEventListener("click", async () => {
  const selected = selectedSipedeDestinationIds();
  if (!selected.length || !sipedeUploadButton) return;
  const original = sipedeDestinationSend.textContent;
  sipedeDestinationSend.disabled = true;
  sipedeDestinationSend.textContent = "Membuat PDF…";
  try {
    const pdf = await createSipedePdf();
    sipedeDestinationSend.textContent = "Mengirim…";
    const formData = new FormData();
    formData.append("destinations", JSON.stringify(selected));
    formData.append("sipede_number", document.querySelector('input[name="sipede_number"]')?.value || "");
    formData.append("document", pdf.blob, pdf.filename);
    const response = await fetch(sipedeUploadButton.dataset.url, {
      method: "POST", headers: { "Accept": "application/json" }, body: formData
    });
    const result = await response.json();
    if (result.requires_sipede_login) {
      closeSipedeDestinationModal();
      window.dispatchEvent(new CustomEvent("sipede-auth-start"));
      return;
    }
    if (!response.ok || !result.uploaded) throw new Error(result.message || "Upload SIPede gagal.");
    const completedUploadButton = sipedeUploadButton;
    closeSipedeDestinationModal();
    if (completedUploadButton) {
      completedUploadButton.disabled = true;
      completedUploadButton.textContent = "Sudah di SIPede";
    }
    document.querySelector("[data-sipede-upload-notice]")?.remove();
    window.location.href = result.redirect_url || "/lapinsus";
  } catch (error) {
    alert(error.message || "LAPINSUS gagal dikirim ke SIPede.");
    sipedeDestinationSend.disabled = false;
    sipedeDestinationSend.textContent = original;
  }
});

document.querySelector("[data-sipede-upload]")?.addEventListener("click", async (event) => {
  const button = event.currentTarget;
  if (button.dataset.afterSaveReady !== "true") return;
  if (!button.dataset.url) return alert("Simpan LAPINSUS terlebih dahulu sebelum upload ke Sipede.");
  if (!await confirmSipedeUpload()) return;
  const original = button.textContent;
  button.disabled = true;
  button.textContent = "Menghubungkan Sipede…";
  try {
    const response = await fetch(button.dataset.url, { method: "POST", headers: { "Accept": "application/json" } });
    const result = await response.json();
    if (result.requires_sipede_login) {
      button.disabled = false;
      button.textContent = original;
      window.dispatchEvent(new CustomEvent("sipede-auth-start"));
      return;
    }
    if (!response.ok) throw new Error(result.message || "Upload Sipede gagal.");
    if (result.sipede_connected) {
      const sipedeNumberInput = document.querySelector('input[name="sipede_number"]');
      if (sipedeNumberInput && result.sipede_number) {
        sipedeNumberInput.value = result.sipede_number;
        updateSipedePreviewNumber(result.sipede_number);
        await preparePreviewForOutput();
      }
      button.disabled = false;
      button.textContent = original;
      openSipedeDestinationModal(result.destinations, button);
      return;
    }
    button.textContent = "Sudah di Sipede";
    document.querySelector("[data-sipede-upload-notice]")?.remove();
  } catch (error) {
    alert(error.message);
    button.disabled = false;
    button.textContent = original;
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && sipedeConfirmModal && !sipedeConfirmModal.hidden) {
    closeSipedeConfirmation(false);
  }
  if (event.key === "Escape" && sipedeDestinationModal && !sipedeDestinationModal.hidden) {
    closeSipedeDestinationModal();
  }
});
window.addEventListener("beforeprint", paginateReportPreview);
window.addEventListener("afterprint", () => document.body.classList.remove("printing-preview"));

window.addEventListener("resize", scheduleReportPagination);

window.addEventListener("load", () => {
  const params = new URLSearchParams(window.location.search);
  const action = params.get("after_save");
  if (!new Set(["docx", "print", "pdf", "sipede"]).has(action)) return;
  params.delete("after_save");
  const cleanQuery = params.toString();
  history.replaceState({}, "", `${window.location.pathname}${cleanQuery ? `?${cleanQuery}` : ""}${window.location.hash}`);
  window.setTimeout(() => {
    if (action === "docx") {
      const sourceButton = document.querySelector('[data-save-before-action="docx"]');
      const submitButton = document.createElement("button");
      submitButton.type = "submit";
      submitButton.hidden = true;
      submitButton.setAttribute("formaction", sourceButton?.dataset.exportUrl || "/lapinhar/export/docx");
      lapinharForm.appendChild(submitButton);
      lapinharForm.requestSubmit(submitButton);
      return;
    }
    const selector = action === "print" ? '[data-save-before-action="print"]' :
      (action === "sipede" ? '[data-save-before-action="sipede"]' : '[data-save-before-action="pdf"]');
    const button = document.querySelector(selector);
    if (!button) return;
    button.dataset.afterSaveReady = "true";
    button.click();
  }, 1400);
});
