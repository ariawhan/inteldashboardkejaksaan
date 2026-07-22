const passwordAdminModal = document.querySelector("[data-password-admin-modal]");
function closePasswordAdmin() { if (passwordAdminModal) { passwordAdminModal.hidden = true; document.body.classList.remove("password-admin-open"); } }
document.querySelectorAll("[data-password-user]").forEach((button) => button.addEventListener("click", () => {
  passwordAdminModal.querySelector("[data-password-admin-name]").textContent = button.dataset.userName;
  passwordAdminModal.querySelector("[data-password-admin-form]").action = `/users/${button.dataset.userId}/change-password`;
  passwordAdminModal.hidden = false; document.body.classList.add("password-admin-open");
  passwordAdminModal.querySelector('[name="new_password"]').focus();
}));
passwordAdminModal?.querySelectorAll("[data-password-admin-close]").forEach((button) => button.addEventListener("click", closePasswordAdmin));
passwordAdminModal?.querySelector("form")?.addEventListener("submit", (event) => {
  const form = event.currentTarget;
  if (form.new_password.value !== form.confirm_password.value) { event.preventDefault(); form.confirm_password.setCustomValidity("Konfirmasi password tidak sama."); form.confirm_password.reportValidity(); }
});
