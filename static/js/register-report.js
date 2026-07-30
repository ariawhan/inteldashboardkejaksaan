const registerReportModal = document.querySelector("[data-register-report-modal]");
const registerReportForm = document.querySelector("[data-register-report-form]");
const registerReportMessage = document.querySelector("[data-register-report-message]");
let activeRegisterReportButton = null;

function registerCurrentTime() {
  const now = new Date();
  return `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
}

function setRegisterReportMessage(message = "", type = "error") {
  if (!registerReportMessage) return;
  registerReportMessage.textContent = message;
  registerReportMessage.className = `register-report-message ${type}`;
  registerReportMessage.hidden = !message;
}

function openRegisterReportModal(button) {
  if (!registerReportModal || !registerReportForm) return;
  activeRegisterReportButton = button;
  registerReportForm.reset();
  registerReportForm.elements.received_time.value = registerCurrentTime();
  registerReportForm.elements.information_value.value = "A1";
  registerReportForm.elements.disposition.value = "TL KE KEJATI";
  registerReportForm.elements.follow_up.value = "-SEGERA TL TERUSKAN KE KEJATI -ARSIPKAN";
  registerReportForm.elements.remarks.value = "Arsip";
  registerReportModal.querySelector("[data-register-report-label]").textContent =
    button.dataset.reportLabel || "Laporan";
  registerReportModal.querySelector("[data-register-report-number]").textContent =
    button.dataset.reportNumber || "—";
  registerReportModal.querySelector("[data-register-report-subject]").textContent =
    button.dataset.reportSubject || "Perihal belum diisi";
  setRegisterReportMessage();
  registerReportModal.hidden = false;
  document.body.classList.add("register-report-open");
  registerReportForm.elements.disposition.focus();
}

function closeRegisterReportModal() {
  if (!registerReportModal) return;
  registerReportModal.hidden = true;
  document.body.classList.remove("register-report-open");
  activeRegisterReportButton = null;
  setRegisterReportMessage();
}

document.querySelectorAll("[data-register-report]").forEach((button) => {
  button.addEventListener("click", () => openRegisterReportModal(button));
});

registerReportModal?.querySelectorAll("[data-register-report-close]").forEach((button) => {
  button.addEventListener("click", closeRegisterReportModal);
});

registerReportForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!activeRegisterReportButton?.dataset.url) return;
  const submitButton = registerReportForm.querySelector(".register-report-submit");
  const originalText = submitButton.textContent;
  submitButton.disabled = true;
  submitButton.textContent = "Menyimpan…";
  setRegisterReportMessage();
  try {
    const response = await fetch(activeRegisterReportButton.dataset.url, {
      method: "POST",
      headers: {
        "Accept": "application/json",
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        received_time: registerReportForm.elements.received_time.value,
        information_value: registerReportForm.elements.information_value.value,
        disposition: registerReportForm.elements.disposition.value,
        follow_up: registerReportForm.elements.follow_up.value,
        remarks: registerReportForm.elements.remarks.value
      })
    });
    const responseText = await response.text();
    let result = {};
    try {
      result = responseText ? JSON.parse(responseText) : {};
    } catch (_parseError) {
      result = { message: responseText.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim() };
    }
    if (!response.ok || !result.registered) {
      if (result.already_registered) {
        setRegisterReportMessage(result.message || "Laporan sudah terdaftar.", "success");
        setTimeout(() => window.location.reload(), 800);
        return;
      }
      throw new Error(result.message || `Register gagal dengan status ${response.status}.`);
    }
    setRegisterReportMessage(result.message || "Laporan berhasil diregister.", "success");
    setTimeout(() => window.location.reload(), 900);
  } catch (error) {
    setRegisterReportMessage(error.message || "Laporan gagal diregister.");
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = originalText;
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && registerReportModal && !registerReportModal.hidden) {
    closeRegisterReportModal();
  }
});
