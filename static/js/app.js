const toggle = document.querySelector("[data-password-toggle]");
if (toggle) {
  toggle.addEventListener("click", () => {
    const input = document.querySelector("#password");
    const visible = input.type === "text";
    input.type = visible ? "password" : "text";
    toggle.setAttribute("aria-label", visible ? "Tampilkan kata sandi" : "Sembunyikan kata sandi");
    toggle.classList.toggle("active", !visible);
  });
}

const forgotPasswordModal = document.querySelector("[data-forgot-password-modal]");
document.querySelector("[data-forgot-password-open]")?.addEventListener("click", () => {
  forgotPasswordModal.hidden = false;
  document.body.classList.add("forgot-password-open");
  forgotPasswordModal.querySelector("[data-forgot-password-close]")?.focus();
});
forgotPasswordModal?.querySelectorAll("[data-forgot-password-close]").forEach((button) => button.addEventListener("click", () => {
  forgotPasswordModal.hidden = true;
  document.body.classList.remove("forgot-password-open");
}));

const profileToggle = document.querySelector("[data-profile-toggle]");
const profileDropdown = document.querySelector("[data-profile-dropdown]");
if (profileToggle && profileDropdown) {
  profileToggle.addEventListener("click", (event) => {
    event.stopPropagation();
    profileDropdown.classList.toggle("open");
  });
  document.addEventListener("click", () => profileDropdown.classList.remove("open"));
}

const sidebarToggle = document.querySelector("[data-sidebar-toggle]");
if (sidebarToggle) {
  sidebarToggle.addEventListener("click", () => document.querySelector("#sidebar").classList.toggle("open"));
}

document.querySelectorAll("[data-preview]").forEach((input) => {
  const outputs = document.querySelectorAll(`[data-output="${input.dataset.preview}"]`);
  const update = () => {
    let value = input.value || "—";
    if (input.dataset.preview === "date" && input.value) {
      value = new Intl.DateTimeFormat("id-ID", { day: "numeric", month: "long", year: "numeric", timeZone: "UTC" }).format(new Date(`${input.value}T00:00:00Z`));
    }
    outputs.forEach((output) => output.textContent = value);
  };
  input.addEventListener("input", update);
  update();
});

document.querySelectorAll("[data-preview-list]").forEach((input) => {
  const output = document.querySelector(`[data-output-list="${input.dataset.previewList}"]`);
  const update = () => {
    const lines = input.value.split("\n").map((line) => line.trim()).filter(Boolean);
    output.innerHTML = "";
    (lines.length ? lines : ["Isi bagian ini akan tampil di sini."]).forEach((line) => {
      const item = document.createElement("p"); item.textContent = line; output.appendChild(item);
    });
  };
  input.addEventListener("input", update);
  update();
});

const previewTabs = document.querySelectorAll("[data-preview-tab]");
const previewPages = document.querySelectorAll("[data-preview-page]");
previewTabs.forEach((tab) => {
  tab.addEventListener("click", () => {
    previewTabs.forEach((item) => item.classList.toggle("active", item === tab));
    previewPages.forEach((page) => page.classList.toggle("active", page.dataset.previewPage === tab.dataset.previewTab));
    document.querySelector(".preview-pages")?.scrollTo({ top: 0, behavior: "smooth" });
    if (tab.dataset.previewTab === "laporan") window.scheduleReportPagination?.();
  });
});
