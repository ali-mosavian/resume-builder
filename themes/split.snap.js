// The page glides from one item to the next, the way Apple's product pages do: each flick of the wheel or trackpad
// is one eased move to the next (or previous) section head, role, school or quote, centered in the window with the
// rest dimmed. Trackpad momentum is swallowed, so a flick never double-steps. An item taller than the window scrolls
// normally until its end is on screen. Off on narrow screens and with reduced motion.
(() => {
  const wide = matchMedia("(min-width:900px)");
  const calm = matchMedia("(prefers-reduced-motion:reduce)");
  const AT = 72; // the least room above a locked item
  const GLIDE = 900; // ms
  const GAP = 90; // ms of wheel silence that ends a gesture
  let gliding = false;
  let lastWheel = 0;
  let spent = false; // this gesture already moved the page
  let items = [];
  let raf = 0;

  const on = () => wide.matches && !calm.matches;
  const abs = (el) => el.getBoundingClientRect().top + scrollY;
  const collect = () => {
    const page = document.querySelector(".page:not([hidden])");
    const q = ".summary, aside, .education h2, .education .item, .experience h2, .experience .item, .testimonials h2, .testimonials .item";
    items = page ? [...page.querySelectorAll(q)].filter((el) => el.offsetParent && el.offsetHeight) : [];
    items.sort((a, b) => abs(a) - abs(b));
  };
  const room = () => document.documentElement.scrollHeight - innerHeight;
  const spot = (el) => {
    const h = el.offsetHeight;
    const lead = h + 2 * AT < innerHeight ? (innerHeight - h) / 2 : AT; // short items are centered
    return Math.max(0, Math.min(room(), abs(el) - lead));
  };
  const mark = () => {
    raf = 0;
    document.documentElement.classList.toggle("snapping", on());
    if (!on()) return;
    let here = null;
    for (const el of items) if (spot(el) <= scrollY + 12) here = el;
    for (const el of items) el.classList.toggle("here", el === here);
  };
  const queue = () => raf || (raf = requestAnimationFrame(mark));
  const ease = (t) => (t < 0.5 ? 4 * t * t * t : 1 - (-2 * t + 2) ** 3 / 2);
  const glide = (to) => {
    const from = scrollY;
    const t0 = performance.now();
    gliding = true;
    const step = (now) => {
      const t = Math.min(1, (now - t0) / GLIDE);
      scrollTo(0, from + (to - from) * ease(t));
      if (t < 1) requestAnimationFrame(step);
      else gliding = false;
    };
    requestAnimationFrame(step);
  };

  addEventListener("wheel", (e) => {
    if (!on() || e.ctrlKey || Math.abs(e.deltaX) > Math.abs(e.deltaY)) return;
    const now = performance.now();
    const fresh = now - lastWheel > GAP; // a new gesture, not the tail of the last one's momentum
    lastWheel = now;
    if (gliding || (spent && !fresh)) return e.preventDefault();
    if (fresh) spent = false;
    if (Math.abs(e.deltaY) < 4) return;
    collect();
    const down = e.deltaY > 0;
    const cur = [...items].reverse().find((el) => spot(el) <= scrollY + 12);
    if (cur) {
      const r = cur.getBoundingClientRect();
      if (down && r.bottom > innerHeight + 12 && r.height + 2 * AT >= innerHeight) return; // taller than the window: read on
      if (!down && r.top < AT - 12 && r.height + 2 * AT >= innerHeight) return;
    }
    const dest = down ? items.find((el) => spot(el) > scrollY + 12) : [...items].reverse().find((el) => spot(el) < scrollY - 12);
    if (!dest) return;
    e.preventDefault();
    spent = true;
    glide(spot(dest));
  }, { passive: false });
  addEventListener("scroll", queue, { passive: true });
  addEventListener("resize", () => { collect(); queue(); });
  addEventListener("langchange", () => { collect(); queue(); });
  collect();
  queue();
})();
