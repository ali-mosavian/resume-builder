// Section menu under the headline numbers, as on Brittany Chiang's page: a line that grows and a label that
// brightens for the section being read. Built from each language's own headings.
(() => {
  // An eased glide for clicks: it speeds up, then settles softly. Any wheel, touch or key press by the reader stops it.
  let stop = () => {};
  const glide = (to) => {
    stop();
    to = Math.max(0, Math.min(document.documentElement.scrollHeight - innerHeight, to));
    const from = scrollY;
    if (Math.abs(to - from) < 2 || matchMedia("(prefers-reduced-motion:reduce)").matches) return scrollTo(0, to);
    const ms = Math.min(1400, 450 + Math.abs(to - from) * 0.45);
    const t0 = performance.now();
    let id = 0;
    const end = () => { cancelAnimationFrame(id); for (const ev of ["wheel", "touchstart", "keydown"]) removeEventListener(ev, end); };
    const step = (now) => {
      const t = Math.min(1, (now - t0) / ms);
      scrollTo(0, from + (to - from) * (t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2));
      if (t < 1) id = requestAnimationFrame(step);
      else end();
    };
    stop = end;
    for (const ev of ["wheel", "touchstart", "keydown"]) addEventListener(ev, end, { passive: true });
    id = requestAnimationFrame(step);
  };
  window.splitGlide = (el) => {
    const r = el.getBoundingClientRect();
    const lead = r.height + 144 < innerHeight ? (innerHeight - r.height) / 2 : 96; // short items are centered
    glide(r.top + scrollY - lead);
  };
  for (const page of document.querySelectorAll(".page")) {
    const targets = [
      [".summary", page.querySelector(".summary h2")],
      [".education", page.querySelector(".education h2")],
      [".experience", page.querySelector(".experience h2")],
      [".testimonials", page.querySelector(".testimonials h2")],
    ].filter(([, h]) => h);
    const skills = page.querySelector("aside h2:nth-of-type(3)");
    if (skills) targets.splice(1, 0, ["aside", skills]);
    const nav = document.createElement("nav");
    nav.className = "sections";
    const items = targets.map(([sel, h]) => {
      const a = document.createElement("a");
      a.href = "#";
      a.innerHTML = `<i></i><span>${h.textContent}</span>`;
      const el = page.querySelector(sel);
      a.addEventListener("click", (e) => {
        e.preventDefault();
        glide((sel === "aside" ? skills : h).getBoundingClientRect().top + scrollY - 48);
      });
      nav.append(a);
      return [a, el];
    });
    page.querySelector(".highlights").after(nav);
    const mark = () => {
      if (page.hidden) return;
      const line = innerHeight * 0.35;
      let on = items[0];
      for (const it of items) if (it[1].getBoundingClientRect().top <= line) on = it;
      for (const [a] of items) a.classList.toggle("on", a === on[0]);
    };
    addEventListener("scroll", mark, { passive: true });
    addEventListener("resize", mark);
    addEventListener("langchange", mark);
    mark();
  }
})();
