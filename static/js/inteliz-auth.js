const connectButton = document.querySelector("[data-inteliz-connect]");
const authModal = document.querySelector("[data-inteliz-modal]");
const authMessage = document.querySelector("[data-inteliz-message]");
const loading = document.querySelector("[data-inteliz-loading]");
const captchaForm = document.querySelector("[data-inteliz-captcha-form]");
const captchaImage = document.querySelector("[data-inteliz-captcha-image]");
const otpForm = document.querySelector("[data-inteliz-otp-form]");
const successPanel = document.querySelector("[data-inteliz-success]");
let authId = null;
let statusTimer = null;
let finished = false;

function showAuthStep(status, result) {
  authMessage.textContent = result.message || "Memproses login Inteliz…";
  loading.hidden = !["starting", "loading"].includes(status);
  captchaForm.hidden = status !== "captcha";
  otpForm.hidden = status !== "otp";
  successPanel.hidden = status !== "success";
  if (status === "captcha" && result.captcha) {
    captchaImage.src = result.captcha;
    captchaForm.querySelector("input").focus();
  }
  if (status === "otp") otpForm.querySelector("input").focus();
  if (status === "success" || status === "error") {
    finished = true;
    clearTimeout(statusTimer);
    if (status === "success") setTimeout(() => {
      window.location.href = authModal?.dataset.nextUrl || window.location.href.split("?")[0];
    }, 1400);
  }
}

async function pollAuthStatus() {
  if (!authId || finished) return;
  try {
    const response = await fetch(`/settings/inteliz/connect/${authId}/status`);
    if (!response.ok) throw new Error("Sesi login tidak ditemukan.");
    const result = await response.json();
    showAuthStep(result.status, result);
    if (!finished) statusTimer = setTimeout(pollAuthStatus, 900);
  } catch (error) {
    showAuthStep("error", { message: error.message });
  }
}

async function startIntelizAuth() {
  authModal.hidden = false;
  document.body.classList.add("modal-open");
  finished = false;
  showAuthStep("starting", { message: "Menyiapkan browser di komputer server…" });
  try {
    const response = await fetch("/settings/inteliz/connect/start", { method: "POST" });
    const result = await response.json();
    if (!response.ok) throw new Error(result.message || "Tidak dapat memulai login Inteliz.");
    authId = result.auth_id;
    pollAuthStatus();
  } catch (error) {
    showAuthStep("error", { message: error.message });
  }
}

connectButton?.addEventListener("click", startIntelizAuth);

captchaForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const input = captchaForm.elements.captcha;
  const response = await fetch(`/settings/inteliz/connect/${authId}/captcha`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ captcha: input.value }) });
  const result = await response.json();
  if (!response.ok) return alert(result.message || "CAPTCHA tidak dapat dikirim.");
  input.value = "";
  showAuthStep("loading", { message: result.message });
});

otpForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const input = otpForm.elements.otp;
  const response = await fetch(`/settings/inteliz/connect/${authId}/otp`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ otp: input.value }) });
  const result = await response.json();
  if (!response.ok) return alert(result.message || "Kode autentikator tidak dapat dikirim.");
  input.value = "";
  showAuthStep("loading", { message: result.message });
});

document.querySelector("[data-inteliz-close]")?.addEventListener("click", async () => {
  clearTimeout(statusTimer);
  if (authId && !finished) await fetch(`/settings/inteliz/connect/${authId}/cancel`, { method: "POST" });
  authModal.hidden = true;
  document.body.classList.remove("modal-open");
  authId = null;
});

if (authModal?.dataset.autoConnect === "true") {
  requestAnimationFrame(startIntelizAuth);
}
