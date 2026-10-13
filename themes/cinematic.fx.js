// Cinematic effects. Narrow screens get the film: reveals, the role in focus, and a scroll-progress fallback for
// browsers without CSS scroll-driven animations (scene timing lives in cinematic.css; the vh ranges below mirror it).
// Wide screens get A4 sheets (cinematic.sheets.js): sheets rise into view and the role at the reading line lights up.
(() => {
  const calm = matchMedia("(prefers-reduced-motion: reduce)").matches;
  if (calm) return;  // static cut: nothing moves, nothing hides
  if (matchMedia("(min-width: 860px)").matches) return sheets();
  const root = document.documentElement;
  root.classList.add("cine");
  const pages = [...document.querySelectorAll(".page")];
  for (const p of pages) p.querySelector("header").style.setProperty("--n", p.querySelectorAll(".highlights div").length);

  // Reveal once, on first sight.
  const io = new IntersectionObserver((entries) => {
    for (const e of entries) if (e.isIntersecting) { e.target.classList.add("in"); io.unobserve(e.target); }
  }, { rootMargin: "0px 0px -10% 0px" });
  const revealable = ".experience .item, .education .item, .testimonials .item, main h2, .summary p, aside > *";
  for (const el of document.querySelectorAll(revealable)) io.observe(el);

  // The role nearest the reading line (40% down the viewport) is in focus.
  const focus = () => {
    const page = pages.find((p) => !p.hidden);
    if (!page) return;
    const line = innerHeight * .4;
    let best = null, dist = Infinity;
    for (const it of page.querySelectorAll(".experience .item")) {
      const r = it.getBoundingClientRect();
      const d = r.top > line ? r.top - line : r.bottom < line ? line - r.bottom : 0;
      if (d < dist) { dist = d; best = it; }
    }
    for (const it of page.querySelectorAll(".experience .item.focus")) if (it !== best) it.classList.remove("focus");
    if (best && dist < innerHeight * .5) best.classList.add("focus");
  };

  // Fallback scrubber: --p per scene, same ranges as the animation-range rules in cinematic.css.
  const native = CSS.supports("animation-timeline: scroll()");
  const clamp = (x) => Math.max(0, Math.min(1, x));
  const scrub = () => {
    const page = pages.find((p) => !p.hidden);
    if (!page) return;
    const vh = innerHeight, y = scrollY;
    const set = (el, p) => el && el.style.setProperty("--p", p.toFixed(4));
    const hero = clamp(y / (.7 * vh));
    for (const sel of ["h1", ".role", ".where", "header", ".contact"]) set(page.querySelector(sel), hero);
    const stats = page.querySelectorAll(".highlights div");
    stats.forEach((d, k) => set(d, clamp((y - (.9 + .55 * k) * vh) / (.65 * vh))));
    const end = (.9 + .55 * stats.length + .1) * vh;
    set(page, clamp((y - end) / (.4 * vh)));
    const exp = page.querySelector(".experience"), r = exp.getBoundingClientRect();
    set(exp, clamp((vh / 2 - r.top) / r.height));
    set(document.body, clamp(y / (root.scrollHeight - vh)));
  };

  let queued = false;
  const onScroll = () => {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => { queued = false; focus(); if (!native) scrub(); });
  };
  addEventListener("scroll", onScroll, { passive: true });
  addEventListener("resize", onScroll);
  for (const b of document.querySelectorAll(".lang-switch button")) b.addEventListener("click", onScroll);
  onScroll();

  function sheets() {
    const seen = new IntersectionObserver((entries) => {
      for (const e of entries) if (e.isIntersecting) { e.target.classList.add("in"); seen.unobserve(e.target); }
    }, { threshold: .02 });
    new MutationObserver((muts) => {
      for (const m of muts) for (const n of m.addedNodes) if (n.classList?.contains("sheet")) { n.classList.add("pre"); seen.observe(n); }
    }).observe(document.body, { childList: true, subtree: true });

    let queued = false;
    const focus = () => {
      queued = false;
      const page = document.querySelector(".page:not([hidden])"), line = innerHeight * .4;
      let best = null, dist = Infinity;
      for (const it of page?.querySelectorAll(".sheet .experience .item:not(.cont)") ?? []) {
        const r = it.getBoundingClientRect(), d = r.top > line ? r.top - line : r.bottom < line ? line - r.bottom : 0;
        if (d < dist) { dist = d; best = it; }
      }
      for (const it of document.querySelectorAll(".experience .item.focus")) if (it !== best) it.classList.remove("focus");
      if (best && dist < innerHeight * .3) best.classList.add("focus");
    };
    addEventListener("scroll", () => queued || ((queued = true), requestAnimationFrame(focus)), { passive: true });
  }
})();
