(() => {
  const form = document.getElementById('fsd-search-form');
  if (!form) return;
  const input = document.getElementById('fsd-station-search');
  const section = document.getElementById('fsd-coordinates');
  const rows = Array.from(section.querySelectorAll('[data-fsd-station]'));
  const update = () => {
    const query = input.value.trim().toUpperCase();
    let count = 0;
    rows.forEach(row => {
      row.hidden = !row.dataset.fsdStation.toUpperCase().includes(query);
      if (!row.hidden) count += 1;
    });
    document.getElementById('fsd-match-count').textContent = count;
    document.getElementById('fsd-no-matches').hidden = count > 0;
    document.getElementById('fsd-search-status').textContent = `${count} of ${rows.length} signals shown`;
    const download = new URL(document.getElementById('fsd-signal-export').href);
    download.searchParams.set('home_station', query);
    document.getElementById('fsd-signal-export').href = download.href;
    const url = new URL(window.location.href);
    if (query) url.searchParams.set('home_station', query);
    else url.searchParams.delete('home_station');
    window.history.replaceState(null, '', url);
  };
  input.addEventListener('input', update);
  form.addEventListener('submit', event => {
    event.preventDefault();
    update();
  });
  document.getElementById('fsd-search-reset').addEventListener('click', event => {
    event.preventDefault();
    input.value = '';
    update();
    input.focus();
  });
  update();
})();
