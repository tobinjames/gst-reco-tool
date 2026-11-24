// A search box above every long table that has none of its own: type to show only the rows that contain the text.
export function initTableSearch() {
  const main = document.querySelector('main');
  if (!main) return;
  let timer = 0;
  const run = () => {
    main.querySelectorAll<HTMLElement>('.tablewrap:not([data-tq])').forEach(w => {
      if (w.closest('#rv') || w.classList.contains('rv-tw')) return; // the review tables have their own filters
      const t = w.querySelector('table');
      if (!t) return;
      const rows = Array.from(t.rows).filter(r => !r.querySelector('th'));
      if (rows.length < 15) return;
      w.setAttribute('data-tq', '1');
      const box = document.createElement('div');
      box.className = 'tq';
      box.innerHTML =
        '<input type="search" placeholder="Search this table" aria-label="Search this table"/><span class="muted"></span>';
      const input = box.querySelector('input')!,
        count = box.querySelector('span')!;
      input.oninput = () => {
        const q = input.value.trim().toLowerCase();
        let n = 0;
        for (const r of rows) {
          const show = !q || (r.textContent || '').toLowerCase().includes(q);
          r.style.display = show ? '' : 'none';
          if (show) n++;
        }
        count.textContent = q ? `${n} of ${rows.length} rows` : '';
      };
      w.parentElement?.insertBefore(box, w);
    });
  };
  new MutationObserver(() => {
    clearTimeout(timer);
    timer = window.setTimeout(run, 150);
  }).observe(main, { childList: true, subtree: true });
}
