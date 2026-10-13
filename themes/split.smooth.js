// Mouse-wheel notches glide instead of stepping: each notch moves a target and the page eases toward it, so fast
// scrolling stays fluid. Trackpads and other fine-grained input already scroll smoothly and are left to the browser.
(() => {
  const wide = matchMedia("(min-width:900px)");
  const calm = matchMedia("(prefers-reduced-motion:reduce)");
  const EASE = 0.11; // share of the remaining distance covered per 60 Hz frame
  let target = 0;
  let id = 0;
  let last = 0;
  const room = () => document.documentElement.scrollHeight - innerHeight;
  const notch = (e) => e.deltaMode === 1 || (e.wheelDeltaY !== undefined && e.wheelDeltaY !== 0 && e.wheelDeltaY % 120 === 0);
  const step = (now) => {
    const dt = Math.min(64, now - last);
    last = now;
    const gap = target - scrollY;
    if (Math.abs(gap) < 1) { scrollTo(0, target); id = 0; return; }
    const move = gap * (1 - Math.pow(1 - EASE, dt / 16.7));
    scrollTo(0, scrollY + (Math.abs(move) < 1 ? Math.sign(gap) : move)); // whole pixels at least, so it never stalls
    id = requestAnimationFrame(step);
  };
  addEventListener("wheel", (e) => {
    if (!wide.matches || calm.matches || e.ctrlKey || e.defaultPrevented || Math.abs(e.deltaX) > Math.abs(e.deltaY) || !notch(e)) return;
    if (e.target.closest?.("textarea, select, [contenteditable]")) return;
    e.preventDefault();
    if (!id) { target = scrollY; last = performance.now(); id = requestAnimationFrame(step); }
    target = Math.max(0, Math.min(room(), target + (e.deltaMode === 1 ? e.deltaY * 40 : e.deltaY)));
  }, { passive: false });
  // anything else that moves the page (keys, scrollbar, a click glide) takes over
  addEventListener("keydown", () => { cancelAnimationFrame(id); id = 0; });
  addEventListener("touchstart", () => { cancelAnimationFrame(id); id = 0; }, { passive: true });
})();
