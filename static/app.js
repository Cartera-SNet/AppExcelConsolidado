const input = document.getElementById('files');
const fileCount = document.getElementById('file-count');

if (input && fileCount) {
  input.addEventListener('change', () => {
    const files = Array.from(input.files || []);
    if (!files.length) {
      fileCount.textContent = 'Ningún archivo seleccionado';
      return;
    }
    fileCount.textContent = `${files.length} archivo(s): ${files.map(f => f.name).join(', ')}`;
  });
}
