// The page locks on to one item at a time: a wheel or trackpad flick moves to the next (or previous) section head,
// role, school or quote. An item taller than the window scrolls normally until its end is on screen.
(() => {
  const wide = matchMedia("(min-width:900px)");
  const calm = matchMedia("(prefers-reduced-motion:reduce)");
  let busy = 0;
  const AT = 72; // where a locked item sits below the top of the window
  const items = () => {
    const page = document.querySelector(".page:not([hidden])");
    if (!page) return [];
    const q = ".summary, aside, .education h2, .education .item, .experience h2, .experience .item, .testimonials h2, .testimonials .item";
    const y = (el) => el.getBoundingClientRect().top;
    return [...page.querySelectorAll(q)].filter((el) => el.offsetParent && el.getBoundingClientRect().height).sort((a, b) => y(a) - y(b));
  };
  addEventListener("wheel", (e) => {
    if (!wide.matches || calm.matches || e.ctrlKey || Math.abs(e.deltaY) < 4 || Math.abs(e.deltaX) > Math.abs(e.deltaY)) return;
    if (e.target.closest?.("textarea, select, [contenteditable]")) return;
    const list = items();
    const top = (el) => el.getBoundingClientRect().top;
    const down = e.deltaY > 0;
    const cur = [...list].reverse().find((el) => top(el) <= AT + 12);
    if (cur) {
      const r = cur.getBoundingClientRect();
      if (down && r.bottom > innerHeight + 12) return; // taller than the window: read on
      if (!down && r.top < AT - 12) return;
    }
    const dest = down ? list.find((el) => top(el) > AT + 12) : [...list].reverse().find((el) => top(el) < AT - 12);
    if (!dest) return;
    e.preventDefault();
    if (Date.now() < busy) return;
    busy = Date.now() + 650;
    scrollBy({ top: top(dest) - AT, behavior: "smooth" });
  }, { passive: false });
})();
