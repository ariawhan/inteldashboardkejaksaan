const lapinsusConvertModal = document.querySelector("[data-lapinsus-convert-modal]");
let lapinsusConvertForm = null;

function closeLapinsusConvertModal() {
  if (!lapinsusConvertModal) return;
  lapinsusConvertModal.hidden = true;
  document.body.classList.remove("lapinsus-convert-open");
  lapinsusConvertForm = null;
}

document.querySelectorAll("[data-lapinsus-convert]").forEach((button) => {
  button.addEventListener("click", () => {
    lapinsusConvertForm = document.getElementById(button.dataset.formId);
    lapinsusConvertModal.querySelector("[data-lapinsus-convert-number]").textContent = button.dataset.number || "—";
    lapinsusConvertModal.querySelector("[data-lapinsus-convert-subject]").textContent = button.dataset.subject || "—";
    lapinsusConvertModal.hidden = false;
    document.body.classList.add("lapinsus-convert-open");
    lapinsusConvertModal.querySelector("[data-lapinsus-convert-confirm]").focus();
  });
});

lapinsusConvertModal?.querySelectorAll("[data-lapinsus-convert-close]").forEach((button) =>
  button.addEventListener("click", closeLapinsusConvertModal));

lapinsusConvertModal?.querySelector("[data-lapinsus-convert-confirm]")?.addEventListener("click", (event) => {
  if (!lapinsusConvertForm) return;
  event.currentTarget.disabled = true;
  event.currentTarget.textContent = "Membuat Draft…";
  lapinsusConvertForm.submit();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && lapinsusConvertModal && !lapinsusConvertModal.hidden) closeLapinsusConvertModal();
});
