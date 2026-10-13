// The item in view is marked as you scroll: the role, school or quote nearest the middle of the window gets the
// card highlight a mouse-over gives, and the rest recede. Scrolling itself is left to the browser, so wheel and
// trackpad momentum are untouched.
(() => {
  const wide = matchMedia("(min-width:900px)");
  const calm = matchMedia("(prefers-reduced-motion:reduce)");
  let items = [];
  let raf = 0;
  const collect = () => {
    const page = document.querySelector(".page:not([hidden])");
    const q = ".education .item, .experience .item, .testimonials .item";
    items = page ? [...page.querySelectorAll(q)].filter((el) => el.offsetParent && el.offsetHeight) : [];
  };
  const mark = () => {
    raf = 0;
    const on = wide.matches;
    document.documentElement.classList.toggle("snapping", on);
    if (!on) return;
    const mid = innerHeight * 0.45;
    let here = null;
    let best = Infinity;
    for (const el of items) {
      const r = el.getBoundingClientRect();
      const d = r.top <= mid && r.bottom >= mid ? 0 : Math.min(Math.abs(r.top - mid), Math.abs(r.bottom - mid));
      if (d < best) [here, best] = [el, d];
    }
    for (const el of items) el.classList.toggle("here", el === here);
  };
  const queue = () => raf || (raf = requestAnimationFrame(mark));
  addEventListener("scroll", queue, { passive: true });
  addEventListener("resize", () => { collect(); queue(); });
  addEventListener("langchange", () => { collect(); queue(); });
  document.addEventListener("click", (e) => {
    const item = e.target.closest?.(".education .item, .experience .item, .testimonials .item");
    // a click on another item brings it to the middle, without cancelling what the click itself does
    if (item && !item.classList.contains("here") && wide.matches && !e.target.closest("a, button, .tools li")) window.splitGlide?.(item);
    setTimeout(() => { collect(); queue(); }, 400);
  });
  collect();
  queue();
})();
