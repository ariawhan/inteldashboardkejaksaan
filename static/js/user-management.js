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

const userEditModal = document.querySelector("[data-user-edit-modal]");
function closeUserEdit() {
  if (!userEditModal) return;
  userEditModal.hidden = true;
  document.body.classList.remove("password-admin-open");
}
document.querySelectorAll("[data-edit-user]").forEach((button) => button.addEventListener("click", () => {
  const form = userEditModal.querySelector("[data-user-edit-form]");
  const image = userEditModal.querySelector("[data-user-edit-image]");
  const initial = userEditModal.querySelector("[data-user-edit-initial]");
  form.action = `/users/${button.dataset.userId}/update-profile`;
  form.full_name.value = button.dataset.userName || "";
  form.new_password.value = "";
  form.confirm_password.value = "";
  form.profile_photo.value = "";
  if (button.dataset.userAvatar) {
    image.src = button.dataset.userAvatar;
    image.hidden = false;
    initial.hidden = true;
  } else {
    image.hidden = true;
    initial.hidden = false;
    initial.textContent = (button.dataset.userName || "U").slice(0, 1).toUpperCase();
  }
  image.style.transform = "scale(1)";
  const zoom = userEditModal.querySelector("[data-user-edit-zoom]");
  if (zoom) zoom.value = "1";
  userEditModal.hidden = false;
  document.body.classList.add("password-admin-open");
  form.full_name.focus();
}));
userEditModal?.querySelectorAll("[data-user-edit-close]").forEach((button) => button.addEventListener("click", closeUserEdit));
userEditModal?.querySelector("form")?.addEventListener("submit", (event) => {
  const form = event.currentTarget;
  form.confirm_password.setCustomValidity("");
  if (form.new_password.value !== form.confirm_password.value) {
    event.preventDefault();
    form.confirm_password.setCustomValidity("Konfirmasi password tidak sama.");
    form.confirm_password.reportValidity();
  }
});

const userEditPhotoInput = userEditModal?.querySelector("[data-user-edit-photo-input]");
const userEditZoom = userEditModal?.querySelector("[data-user-edit-zoom]");
function updateUserEditZoom() {
  const image = userEditModal?.querySelector("[data-user-edit-image]");
  if (image) image.style.transform = `scale(${userEditZoom?.value || 1})`;
}
userEditPhotoInput?.addEventListener("change", () => {
  const file = userEditPhotoInput.files?.[0];
  if (!file) return;
  const image = userEditModal.querySelector("[data-user-edit-image]");
  const initial = userEditModal.querySelector("[data-user-edit-initial]");
  image.src = URL.createObjectURL(file);
  image.hidden = false;
  initial.hidden = true;
  if (userEditZoom) userEditZoom.value = "1";
  updateUserEditZoom();
});
userEditZoom?.addEventListener("input", updateUserEditZoom);
