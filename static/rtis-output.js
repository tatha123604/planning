(() => {
  const buttons = document.querySelectorAll('[data-rtis-export]');
  const status = document.getElementById('rtis-export-status');
  buttons.forEach((button) => button.addEventListener('click', async () => {
    const extension = button.dataset.rtisExport;
    const format = extension === 'pdf' ? 'PDF' : 'Excel';
    buttons.forEach((item) => { item.disabled = true; });
    status.classList.remove('is-error');
    status.textContent = `Preparing ${format} for all divisions. Please wait…`;
    try {
      const response = await fetch(`/rtis/output.${extension}`, {
        credentials: 'same-origin', headers: { 'X-Requested-With': 'fetch' },
      });
      if (!response.ok) {
        if (response.status === 401) throw new Error('Please sign in again, then retry the download.');
        throw new Error(`${format} export failed. Please retry.`);
      }
      const blob = await response.blob();
      const expectedType = extension === 'pdf' ? 'application/pdf'
        : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      const signature = new Uint8Array(await blob.slice(0, 5).arrayBuffer());
      const expectedSignature = extension === 'pdf' ? [37, 80, 68, 70, 45] : [80, 75, 3, 4];
      if (!blob.type.includes(expectedType)
          || !expectedSignature.every((value, index) => signature[index] === value)) {
        throw new Error('The download was not a valid file. Please sign in again and retry.');
      }
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `rtis_merged_output.${extension}`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.setTimeout(() => URL.revokeObjectURL(url), 60000);
      status.textContent = `${format} download ready. All divisions and pages are included.`;
    } catch (error) {
      status.classList.add('is-error');
      status.textContent = error.message;
    } finally {
      buttons.forEach((item) => { item.disabled = false; });
    }
  }));
})();
