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

const themeToggle = document.querySelector("[data-theme-toggle]");
const applyTheme = (theme) => {
  const dark = theme === "dark";
  document.documentElement.dataset.theme = dark ? "dark" : "light";
  document.body.classList.toggle("dark-mode", dark);
  document.querySelector('meta[name="theme-color"]')?.setAttribute("content", dark ? "#07111f" : "#f4f7f9");
  themeToggle?.setAttribute("aria-label", dark ? "Aktifkan mode terang" : "Aktifkan mode gelap");
  themeToggle?.setAttribute("title", dark ? "Mode terang" : "Mode gelap");
  themeToggle?.setAttribute("aria-pressed", dark ? "true" : "false");
};
applyTheme(localStorage.getItem("indraone-theme") === "dark" ? "dark" : "light");
themeToggle?.addEventListener("click", () => {
  const nextTheme = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
  localStorage.setItem("indraone-theme", nextTheme);
  applyTheme(nextTheme);
});

const sidebarToggle = document.querySelector("[data-sidebar-toggle]");
if (sidebarToggle) {
  sidebarToggle.addEventListener("click", () => document.querySelector("#sidebar").classList.toggle("open"));
}

const sidebarCollapseToggle = document.querySelector("[data-sidebar-collapse-toggle]");
const sidebarCollapseIcon = document.querySelector("[data-sidebar-collapse-icon]");
const setSidebarCollapsed = (collapsed) => {
  document.body.classList.toggle("sidebar-collapsed", collapsed);
  if (sidebarCollapseIcon) sidebarCollapseIcon.textContent = collapsed ? "›" : "‹";
  sidebarCollapseToggle?.setAttribute("aria-label", collapsed ? "Tampilkan menu" : "Sembunyikan menu");
  sidebarCollapseToggle?.setAttribute("aria-expanded", collapsed ? "false" : "true");
};
if (sidebarCollapseToggle) {
  setSidebarCollapsed(document.body.classList.contains("sidebar-collapsed"));
  sidebarCollapseToggle.addEventListener("click", () => {
    setSidebarCollapsed(!document.body.classList.contains("sidebar-collapsed"));
  });
}

const reportActionMenus = document.querySelectorAll("[data-report-action-menu]");
function closeReportActionMenus(exceptMenu = null) {
  reportActionMenus.forEach((menu) => {
    if (menu === exceptMenu) return;
    menu.classList.remove("open");
    menu.querySelector("[data-report-action-toggle]")?.setAttribute("aria-expanded", "false");
  });
}
reportActionMenus.forEach((menu) => {
  const button = menu.querySelector("[data-report-action-toggle]");
  button?.addEventListener("click", (event) => {
    event.stopPropagation();
    const willOpen = !menu.classList.contains("open");
    closeReportActionMenus(menu);
    menu.classList.toggle("open", willOpen);
    button.setAttribute("aria-expanded", willOpen ? "true" : "false");
  });
  menu.addEventListener("click", (event) => event.stopPropagation());
});
document.addEventListener("click", () => closeReportActionMenus());
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeReportActionMenus();
});

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
