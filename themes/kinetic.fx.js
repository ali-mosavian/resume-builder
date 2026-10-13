// Kinetic effects: letters of the name that lift toward the cursor, a tool ticker, magnetic download buttons.
// Runs after kinetic.js (the shared interactions). Transform and colour only.
(() => {
  const calm = matchMedia("(prefers-reduced-motion: reduce)").matches;
  document.documentElement.classList.add("js");
  const pages = [...document.querySelectorAll(".page")];
  const visible = () => pages.find((p) => !p.hidden) || pages[0];
  const pointer = { x: -1e4, y: -1e4, live: false };
  addEventListener("pointermove", (e) => { pointer.x = e.clientX; pointer.y = e.clientY; pointer.live = true; }, { passive: true });
  addEventListener("pointerleave", () => { pointer.live = false; });
  document.addEventListener("mouseleave", () => { pointer.live = false; });

  // --- 1. The name: one span per letter, weight from cursor proximity, width from scroll, a slow idle wave.
  const letters = [];
  for (const h1 of document.querySelectorAll(".page > header > h1")) {
    const name = h1.textContent.trim();
    h1.replaceChildren();
    const sr = document.createElement("span");
    sr.className = "sr";
    sr.textContent = name;
    h1.append(sr);
    name.split(/\s+/).forEach((word, wi) => {
      const w = document.createElement("span");
      w.className = "w";
      w.setAttribute("aria-hidden", "true");
      for (const ch of word) {
        const l = document.createElement("span");
        l.className = "l";
        l.textContent = ch;
        w.append(l);
        letters.push({ el: l, i: letters.length, y0: 0, n: 0, x: 0, y: 0 });
      }
      h1.append(w);
    });
  }

  // Letters lift and warm toward the cursor (a gaussian dent that travels with it), over a slow idle wave. Only
  // transform and colour move: the name never reflows, and the A4 sheet it sits on never scrolls.
  let running = false, idleFrames = 0, heroVisible = false;
  function frame(t) {
    const page = visible();
    const live = letters.filter((l) => page.contains(l.el));
    // Read all rects first, then write: one layout per frame.
    for (const l of live) { const r = l.el.offsetParent ? l.el.getBoundingClientRect() : { left: 0, width: 0, top: 0, height: 0 }; l.x = r.left + r.width / 2; l.y = r.top + r.height / 2; }
    const radius = 140;
    let moving = false;
    for (const l of live) {
      const dx = (pointer.x - l.x) / radius, dy = (pointer.y - l.y) / radius;
      const near = pointer.live ? Math.exp(-(dx * dx + dy * dy) * 2.2) : 0;
      const lift = -14 * near + 2.5 * Math.sin(t / 1500 + l.i * 0.6);
      const k = 0.18;  // ease toward target
      const ny = l.y0 + (lift - l.y0) * k, nn = l.n + (near - l.n) * k;
      if (Math.abs(ny - l.y0) > 0.05 || Math.abs(nn - l.n) > 0.005) moving = true;
      l.y0 = ny; l.n = nn;
      l.el.style.transform = `translateY(${ny.toFixed(2)}px)`;
      l.el.style.color = nn > 0.02 ? `color-mix(in srgb, #f2c230 ${(nn * 100).toFixed(0)}%, #fff)` : "";
    }
    idleFrames = moving ? 0 : idleFrames + 1;
    if (heroVisible || idleFrames < 30) requestAnimationFrame(frame); else running = false;
  }
  const start = () => { if (!running && !calm) { running = true; requestAnimationFrame(frame); } };
  if (letters.length && !calm) {
    const io = new IntersectionObserver((es) => { heroVisible = es.some((e) => e.isIntersecting); if (heroVisible) start(); });
    document.querySelectorAll(".page > header, .sheet > header").forEach((h) => io.observe(h));
    addEventListener("pointermove", start, { passive: true });
    addEventListener("langchange", () => setTimeout(start, 100));
    start();
  }

  // --- 2. Tool ticker: every tool from the roles in one amber strip at the foot of the hero. Click to filter, as a chip would.
  //        The header moves into the first A4 sheet with its ticker (the strip is reserved in the layout, so print is unchanged).
  for (const page of pages) {
    const seen = new Map();
    for (const li of page.querySelectorAll(".tools li")) {
      const k = li.textContent.trim().toLowerCase();
      if (!seen.has(k)) seen.set(k, li);
    }
    const all = [...seen.values()];
    if (all.length < 6) continue;
    const band = document.createElement("div");
    band.className = "band";
    band.setAttribute("aria-hidden", "true");
    const ul = document.createElement("ul");
    for (const src of [...all, ...all]) {  // twice, for a seamless loop
      const li = document.createElement("li");
      li.textContent = src.textContent;
      if (src.dataset.hits) { li.dataset.hits = src.dataset.hits; li.addEventListener("click", () => src.click()); }
      ul.append(li);
    }
    ul.style.setProperty("--dur", `${all.length * 2.4}s`);
    band.append(ul);
    page.querySelector("header").append(band);
  }

  // --- 4. Magnetic buttons in the hero.
  if (!calm && matchMedia("(hover: hover)").matches) {
    const btns = [...document.querySelectorAll(".contact a.download, .contact button")];
    let near = new Set();
    addEventListener("pointermove", () => {
      const page = visible();
      for (const b of btns) {
        if (!page.contains(b)) continue;
        const r = b.getBoundingClientRect();
        const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
        const dx = pointer.x - cx, dy = pointer.y - cy, d = Math.hypot(dx, dy);
        if (d < 90) { b.style.transform = `translate(${dx * 0.3}px, ${dy * 0.3}px)`; near.add(b); }
        else if (near.delete(b)) b.style.transform = "";
      }
    }, { passive: true });
  }
})();
