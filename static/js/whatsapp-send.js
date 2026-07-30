(() => {
  const modal = document.querySelector('[data-wa-choice-modal]');
  if (!modal) return;
  let reportUrl = '';
  const closeModal = () => { modal.classList.remove('open'); modal.setAttribute('aria-hidden', 'true'); };

  document.querySelectorAll('[data-whatsapp-send]').forEach((button) => {
    button.addEventListener('click', () => {
      reportUrl = button.dataset.url || '';
      const reportLabel = button.dataset.reportLabel || 'Laporan';
      const count = Number(button.dataset.attachmentCount || 0);
      const title = modal.querySelector('#wa-choice-title');
      if (title) title.textContent = `Kirim ${reportLabel} ke WhatsApp`;
      modal.querySelector('[data-wa-attachment-note]').textContent = count
        ? `Semua ${count} foto lampiran akan diunduh dalam satu berkas ZIP.`
        : `${reportLabel} ini belum mempunyai foto lampiran.`;
      modal.querySelector('[data-wa-choice="with"]').disabled = count === 0;
      modal.classList.add('open');
      modal.setAttribute('aria-hidden', 'false');
    });
  });
  modal.querySelectorAll('[data-wa-choice-close]').forEach((button) => button.addEventListener('click', closeModal));
  modal.querySelectorAll('[data-wa-choice]').forEach((button) => button.addEventListener('click', () => {
    if (!reportUrl || button.disabled) return;
    const attachments = button.dataset.waChoice === 'with' ? '1' : '0';
    window.open(`${reportUrl}?attachments=${attachments}`, '_blank', 'noopener');
    closeModal();
  }));
  document.addEventListener('keydown', (event) => { if (event.key === 'Escape') closeModal(); });
})();
