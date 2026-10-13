// Kinetic effects: letter-by-letter variable type on the name, tool bands, era colour fields, magnetic buttons.
// Runs after kinetic.js (the shared interactions). Transform / opacity / font-variation-settings only.
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
        letters.push({ el: l, i: letters.length, wght: 700, wdth: 100, x: 0, y: 0 });
      }
      h1.append(w);
    });
  }

  let heroH = 1, scrollP = 0;
  const measureHero = () => { const h = visible().querySelector("header"); heroH = Math.max(1, h.offsetHeight); };
  measureHero();
  addEventListener("resize", measureHero);

  let running = false, idleFrames = 0;
  function frame(t) {
    scrollP = Math.min(1, Math.max(0, scrollY / (heroH * 0.9)));
    // Read all rects first, then write: one layout per frame.
    const page = visible();
    const live = letters.filter((l) => page.contains(l.el));
    for (const l of live) { const r = l.el.getBoundingClientRect(); l.x = r.left + r.width / 2; l.y = r.top + r.height / 2; }
    const radius = Math.max(140, innerWidth * 0.16);
    let moving = false;
    for (const l of live) {
      const dx = (pointer.x - l.x) / radius, dy = (pointer.y - l.y) / radius;
      const near = pointer.live ? Math.exp(-(dx * dx + dy * dy) * 2.2) : 0;  // gaussian falloff
      const wave = Math.sin(t / 1500 + l.i * 0.6);
      // Rest: heavy and wide, breathing slowly; the cursor presses letters thin and narrow (a dent that
      // travels with it); scrolling away thins and condenses the whole name.
      const base = 720 - 480 * scrollP + 50 * wave * (1 - scrollP), baseW = 100 - 25 * scrollP + 3 * wave * (1 - scrollP);
      const wght = base + (240 - base) * near;
      const wdth = baseW + (75 - baseW) * near;
      const k = 0.18;  // ease toward target
      const nw = l.wght + (wght - l.wght) * k, nd = Math.min(100, l.wdth + (wdth - l.wdth) * k);
      if (Math.abs(nw - l.wght) > 0.2 || Math.abs(nd - l.wdth) > 0.05) moving = true;
      l.wght = nw; l.wdth = nd;
      l.el.style.fontVariationSettings = `"opsz" 96, "wdth" ${nd.toFixed(2)}, "wght" ${nw.toFixed(1)}`;
    }
    idleFrames = moving ? 0 : idleFrames + 1;
    if (heroVisible || idleFrames < 30) requestAnimationFrame(frame); else running = false;
  }
  let heroVisible = false;
  const start = () => { if (!running && !calm) { running = true; requestAnimationFrame(frame); } };
  if (letters.length && !calm) {
    const io = new IntersectionObserver((es) => { heroVisible = es.some((e) => e.isIntersecting); if (heroVisible) start(); });
    document.querySelectorAll(".page > header").forEach((h) => io.observe(h));
    addEventListener("pointermove", start, { passive: true });
    addEventListener("scroll", start, { passive: true });
  }

  // --- 2. Tool bands: every tool from the roles, two rows in opposite directions. Click to filter, as a chip would.
  for (const page of pages) {
    const seen = new Map();
    for (const li of page.querySelectorAll(".tools li")) {
      const k = li.textContent.trim().toLowerCase();
      if (!seen.has(k)) seen.set(k, li);
    }
    const all = [...seen.values()];
    if (all.length < 6) continue;
    const half = Math.ceil(all.length / 2);
    const rows = [all.slice(0, half), all.slice(half)];
    rows.forEach((row, i) => {
      const band = document.createElement("div");
      band.className = "band" + (i ? " rev" : "");
      band.setAttribute("aria-hidden", "true");
      const ul = document.createElement("ul");
      for (const src of [...row, ...row]) {  // twice, for a seamless loop
        const li = document.createElement("li");
        li.textContent = src.textContent;
        if (src.dataset.hits) { li.dataset.hits = src.dataset.hits; li.addEventListener("click", () => src.click()); }
        ul.append(li);
      }
      ul.style.setProperty("--dur", `${row.length * 2.4}s`);
      band.append(ul);
      page.querySelector("header").after(band);
    });
  }

  // --- 3. Era colour fields: group roles by start year, head each group with a poster-size range,
  //        and recolour the page to the era under the viewport's upper third.
  const eraOf = (y) => (y >= 2025 ? "now" : y >= 2018 ? "mid" : "early");
  for (const page of pages) {
    const exp = page.querySelector(".experience");
    const present = (JSON.parse(page.dataset.ui || "{}").present) || page.querySelector(".experience .when")?.textContent.split(" - ")[1]?.trim() || "";
    const kids = [...exp.children].filter((el) => el.tagName !== "H2");
    const groups = [];
    let cur = null, pending = [];
    for (const el of kids) {
      if (!el.classList.contains("item")) { pending.push(el); continue; }  // e.g. the details toggle
      const years = (el.querySelector(".when")?.textContent.match(/\d{4}/g) || []).map(Number);
      const era = eraOf(years[0] || 0);
      if (!cur || cur.era !== era) { cur = { era, els: [], years: [], open: era === "now" }; groups.push(cur); }
      cur.els.push(...pending, el); pending = [];
      cur.years.push(...years);  // start and end years: the range an era spans ("present" adds none)
    }
    for (const g of groups) {
      const wrap = document.createElement("div");
      wrap.className = "era"; wrap.dataset.era = g.era;
      const lab = document.createElement("div");
      lab.className = "era-label";
      const lo = Math.min(...g.years), hi = Math.max(...g.years);
      lab.textContent = g.open ? `${lo}→` : lo === hi ? `${lo}` : `${lo}-${hi}`;
      const n = g.els.filter((e) => e.classList.contains("item")).length;
      const small = document.createElement("small");
      small.textContent = g.open ? `${lo} - ${present}`.toUpperCase() : `${n} ${n === 1 ? "role" : "roles"}`;
      lab.append(small);
      wrap.append(lab, ...g.els);
      exp.append(wrap);
    }
  }
  const html = document.documentElement;
  let eraTick = false;
  function eraUpdate() {
    eraTick = false;
    const line = innerHeight * 0.38;
    let era = "";
    for (const g of visible().querySelectorAll(".era")) {
      const r = g.getBoundingClientRect();
      if (r.top <= line && r.bottom > line) { era = g.dataset.era; break; }
    }
    if (html.dataset.era !== era) { if (era) html.dataset.era = era; else delete html.dataset.era; }
  }
  const queueEra = () => { if (!eraTick) { eraTick = true; requestAnimationFrame(eraUpdate); } };
  addEventListener("scroll", queueEra, { passive: true });
  addEventListener("resize", queueEra);
  addEventListener("click", () => setTimeout(() => { measureHero(); queueEra(); }, 50));  // language switch, filters, toggles
  addEventListener("beforeprint", () => delete html.dataset.era);
  queueEra();

  // --- 4. Magnetic buttons in the hero.
  if (!calm && matchMedia("(hover: hover)").matches) {
    const btns = [...document.querySelectorAll(".contact a, .contact button")];
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
