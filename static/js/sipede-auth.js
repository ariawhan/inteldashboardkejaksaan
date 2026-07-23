(() => {
  const connectButton = document.querySelector("[data-sipede-connect]");
  const modal = document.querySelector("[data-sipede-modal]");
  if (!modal) return;

  const message = modal.querySelector("[data-sipede-message]");
  const loading = modal.querySelector("[data-sipede-loading]");
  const captchaForm = modal.querySelector("[data-sipede-captcha-form]");
  const captchaImage = modal.querySelector("[data-sipede-captcha-image]");
  const successPanel = modal.querySelector("[data-sipede-success]");
  let authId = null;
  let statusTimer = null;
  let finished = false;

  function showStep(status, result = {}) {
    message.textContent = result.message || "Memproses login SIPede…";
    loading.hidden = !["starting", "loading"].includes(status);
    captchaForm.hidden = status !== "captcha";
    successPanel.hidden = status !== "success";
    if (status === "captcha" && result.captcha) {
      captchaImage.src = result.captcha;
      captchaForm.elements.captcha.focus();
    }
    if (["success", "error"].includes(status)) {
      finished = true;
      clearTimeout(statusTimer);
      if (status === "success") setTimeout(() => {
        window.dispatchEvent(new CustomEvent("sipede-auth-success"));
        if (modal.dataset.reload === "true") window.location.reload();
        else {
          modal.hidden = true;
          document.body.classList.remove("modal-open");
        }
      }, 1400);
    }
  }

  async function pollStatus() {
    if (!authId || finished) return;
    try {
      const response = await fetch(`/settings/sipede/connect/${authId}/status`);
      if (!response.ok) throw new Error("Sesi login SIPede tidak ditemukan.");
      const result = await response.json();
      showStep(result.status, result);
      if (!finished) statusTimer = setTimeout(pollStatus, 900);
    } catch (error) {
      showStep("error", { message: error.message });
    }
  }

  async function start() {
    modal.hidden = false;
    document.body.classList.add("modal-open");
    finished = false;
    showStep("starting", { message: "Menyiapkan browser di komputer server…" });
    try {
      const response = await fetch("/settings/sipede/connect/start", { method: "POST" });
      const result = await response.json();
      if (!response.ok) throw new Error(result.message || "Tidak dapat memulai login SIPede.");
      authId = result.auth_id;
      pollStatus();
    } catch (error) {
      showStep("error", { message: error.message });
    }
  }

  connectButton?.addEventListener("click", start);
  window.addEventListener("sipede-auth-start", start);
  captchaForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    const input = captchaForm.elements.captcha;
    const response = await fetch(`/settings/sipede/connect/${authId}/captcha`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ captcha: input.value }),
    });
    const result = await response.json();
    if (!response.ok) return alert(result.message || "CAPTCHA tidak dapat dikirim.");
    input.value = "";
    showStep("loading", { message: result.message });
  });
  modal.querySelector("[data-sipede-close]").addEventListener("click", async () => {
    clearTimeout(statusTimer);
    if (authId && !finished) await fetch(`/settings/sipede/connect/${authId}/cancel`, { method: "POST" });
    modal.hidden = true;
    document.body.classList.remove("modal-open");
    authId = null;
  });
})();
