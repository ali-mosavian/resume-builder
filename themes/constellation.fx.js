// Constellation: a star chart of the career drawn from the page. Roles on a time axis, skills as stars above,
// linked by the roles whose text mentions them. Hover lights a star's links and the roles below; click filters.
(() => {
  const calm = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const MONTHS = { jan:0, feb:1, mar:2, apr:3, may:4, maj:4, jun:5, jul:6, aug:7, sep:8, oct:9, okt:9, nov:10, dec:11 };  // ponytail: EN + SV only
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const matcher = (term, aka = "") => new RegExp(  // same rule as the filter in constellation.js
    [...term.split("/"), ...aka.split("|").filter(Boolean)].map((t) => `(?<![\\w-])${esc(t.trim())}(?![\\w-])`).join("|"), "i");
  const MAX_STARS = 32, LABELS = 22;
  const NIGHT = { bg:null, star:"#cfd8ff", gold:"#e8bd63", edge:"#8fa0e6", ink:"#eef1fb", dim:"#8f99bd", halo:"#070a17" };
  const PAPER = { bg:"#fff", star:"#3b47c9", gold:"#b8892a", edge:"#3b47c9", ink:"#141827", dim:"#5b6378", halo:"#fff" };

  const dateOf = (s) => { const m = s.match(/([a-zåäö]{3})\w*\.?\s+(\d{4})/i); return m ? +m[2] + (MONTHS[m[1].toLowerCase()] ?? 0) / 12 : null; };
  const rnd = (seed) => () => (seed = (seed * 16807) % 2147483647) / 2147483647;  // deterministic: same sky every visit

  function collect(page) {
    const roles = [...page.querySelectorAll(".experience .item")].map((el) => {
      const [a, b] = el.querySelector(".when").textContent.split(/\s-\s/);
      const t0 = dateOf(a), t1 = dateOf(b || "") ?? new Date().getFullYear() + new Date().getMonth() / 12;
      const name = el.querySelector(".company").textContent.split(",")[0].trim(), brief = el.classList.contains("brief");
      return { el, name, label: brief ? name.split(" ").slice(0, 2).join(" ") : name, t0, t1, brief,
               text: [...el.querySelectorAll("h3, li, p")].map((e) => e.textContent).join(" \n ") };
    }).filter((r) => r.t0 !== null);
    const stars = [];
    const add = (name, aka, level, chip) => {
      const re = matcher(name, aka), links = roles.filter((r) => re.test(r.text));
      if (links.length) stars.push({ name, re, level, chip, links, w: links.length });
    };
    for (const kv of page.querySelectorAll("aside .kv[style]")) add(kv.firstElementChild.textContent.trim(), kv.dataset.aka || "", +(kv.style.getPropertyValue("--level") || 3), kv);
    // Tools that recur across roles but have no sidebar skill, most used first.
    const seen = new Map();
    for (const li of page.querySelectorAll(".tools li")) {
      const name = li.textContent.trim();
      if (stars.some((s) => s.re.test(name))) continue;
      const e = seen.get(name.toLowerCase()) || seen.set(name.toLowerCase(), { name, chip: li, n: 0 }).get(name.toLowerCase());
      e.n++;
    }
    for (const t of [...seen.values()].filter((t) => t.n > 1).sort((a, b) => b.n - a.n)) {
      if (stars.length >= MAX_STARS) break;
      add(t.name, "", 2, t.chip);
    }
    return { roles, stars };
  }

  function build(page) {
    const { roles, stars } = collect(page);
    if (!roles.length || !stars.length) return;
    const header = page.querySelector("header");
    const sky = document.createElement("div");
    sky.className = "sky";
    sky.setAttribute("role", "img");
    sky.setAttribute("aria-label", `${roles.length} roles, ${stars.length} technologies`);
    const canvas = document.createElement("canvas");
    sky.append(canvas);
    header.append(sky);
    sky.stars = stars; sky.roles = roles;  // for tests and debugging
    const hint = document.createElement("div");
    hint.className = "sky-hint";
    hint.textContent = page.lang === "sv" ? "Peka på en stjärna" : "Hover a star";  // ponytail: two languages
    header.append(hint);
    const kicker = document.createElement("div");
    kicker.className = "kicker";
    const years = [Math.floor(Math.min(...roles.map((r) => r.t0))), Math.floor(Math.max(...roles.map((r) => r.t1)))];
    kicker.textContent = `${years[0]} - ${years[1]}`;
    header.prepend(kicker);

    const ctx = canvas.getContext("2d");
    let W = 0, H = 0, dpr = 1, hot = null, palette = NIGHT, running = false, t0 = performance.now();
    const rand = rnd(7);
    const dust = Array.from({ length: 140 }, () => ({ x: rand(), y: rand(), r: .4 + rand() * 1.1, p: rand() * 6.28, s: .5 + rand() * 1.5 }));
    const wmax = Math.max(...stars.map((s) => s.w));
    let labelled = new Set();
    for (const s of stars) { s.phase = rand() * 6.28; s.r = 2.2 + 2.6 * Math.sqrt(s.w / wmax) + (s.level >= 5 ? .8 : 0); }

    function layout() {
      const small = W < 500, short = H < 360, pad = { l: 36, r: 56, t: short ? 18 : 44 };
      const tmin = years[0], tmax = years[1] + 1;
      // Perspective: recent years get more room (ticks stay labelled, so the warp is visible).
      const xOf = (t) => pad.l + (1 - Math.pow((tmax - t) / (tmax - tmin), .7)) * (W - pad.l - pad.r);
      for (const r of roles) { r.x = xOf(r.t0); r.r = r.brief ? 2 : 3.6; }
      // Role labels below the axis: greedy rows so neighbours do not overlap.
      ctx.font = `500 10px ${mono()}`;
      const rows = [];
      // Main roles first, compact; brief roles (labelled only when lit) take leftover slots, one spare row.
      const place = (list) => {
        for (const r of list.sort((a, b) => a.x - b.x)) {
          const w = ctx.measureText(r.label).width, x0 = Math.min(Math.max(r.x - w / 2, 2), W - w - 2);
          let row = rows.findIndex((end) => end < x0 - 8);
          if (row < 0) row = rows.push(0) - 1;
          rows[row] = x0 + w;
          r.lx = x0; r.row = row;
        }
      };
      place(roles.filter((r) => !r.brief));
      const maxRows = rows.length + 1;
      place(roles.filter((r) => r.brief));
      const axisY = H - 22 - maxRows * 13;
      for (const r of roles) r.y = axisY;
      // Stars: spring to their roles in x, rise by weight in y, keep apart. Seeded, so the sky is stable.
      const top = pad.t, bottom = axisY - (short ? 28 : 46), rand = rnd(11);
      const byWeight = [...stars].sort((a, b) => b.w - a.w);  // heaviest stars highest, the rest spread down to the axis
      for (const s of stars) {
        const cx = s.links.reduce((a, r) => a + r.x, 0) / s.w;
        s.x = cx + (rand() - .5) * 60;
        s.ty = top + byWeight.indexOf(s) / (stars.length - 1) * (bottom - top) * .92 + (rand() - .5) * 40;
        s.y = s.ty + (rand() - .5) * 40;
        s.vx = s.vy = 0;
      }
      const minD = small ? 30 : 48;  // labels sit to the side, so x counts for more
      for (let i = 0; i < 360; i++) {
        for (const s of stars) {
          let fx = 0, fy = (s.ty - s.y) * .02;
          for (const r of s.links) fx += (r.x - s.x) * .004;
          for (const o of stars) {
            if (o === s) continue;
            const dx = s.x - o.x, dy = s.y - o.y, d = Math.hypot(dx * .5, dy) || .1;
            if (d < minD) { const f = (minD - d) / d * .12; fx += dx * f; fy += dy * f; }
          }
          s.vx = (s.vx + fx) * .8; s.vy = (s.vy + fy) * .8;
          s.x = Math.min(Math.max(s.x + s.vx, pad.l), W - pad.r); s.y = Math.min(Math.max(s.y + s.vy, top), bottom);
        }
      }
      // Labels: heaviest first, skipping any that would overlap one already placed; kept inside the box.
      ctx.font = `500 11px ${sans()}`;
      const want = Math.min(LABELS, Math.floor(W * H / (small || short ? 12000 : 9000))), placed = [];
      labelled = new Set();
      for (const s of [...stars].sort((a, b) => b.w - a.w || b.level - a.level)) {
        const w = ctx.measureText(s.name).width;
        s.labelLeft = s.x + 9 + w > W - 4;
        const box = s.labelLeft ? [s.x - 9 - w, s.y - 7, s.x + 4, s.y + 7] : [s.x - 4, s.y - 7, s.x + 9 + w, s.y + 7];
        if (labelled.size >= want || placed.some((b) => box[0] < b[2] && box[2] > b[0] && box[1] < b[3] && box[3] > b[1])) continue;
        placed.push(box); labelled.add(s);
      }
      roles.axisY = axisY; roles.xOf = xOf; roles.tmin = tmin; roles.tmax = tmax; roles.maxRows = maxRows;
    }
    const sans = () => getComputedStyle(document.body).fontFamily;
    const mono = () => getComputedStyle(document.body).getPropertyValue("--mono") || "monospace";

    function draw(now) {
      if (!roles.xOf) return;  // never laid out (hidden language)
      const t = calm ? 0 : (now - t0) / 1000;
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      ctx.clearRect(0, 0, W, H);
      if (palette.bg) { ctx.fillStyle = palette.bg; ctx.fillRect(0, 0, W, H); }
      const P = palette, lit = hot ? new Set(hot.links ? [hot, ...hot.links] : [hot, ...stars.filter((s) => s.links.includes(hot))]) : null;
      // dust
      if (!P.bg) for (const d of dust) {
        ctx.globalAlpha = .25 + .35 * (.5 + .5 * Math.sin(t * d.s + d.p));
        ctx.fillStyle = P.star; ctx.beginPath(); ctx.arc(d.x * W, d.y * H, d.r, 0, 6.28); ctx.fill();
      }
      ctx.globalAlpha = 1;
      const pos = (s) => s.links ? [s.x + Math.sin(t * .35 + s.phase) * 2.5, s.y + Math.cos(t * .28 + s.phase) * 2.5] : [s.x, s.y];
      // axis
      const ay = roles.axisY;
      ctx.strokeStyle = P.dim; ctx.globalAlpha = .45; ctx.lineWidth = 1;
      ctx.beginPath(); ctx.moveTo(roles.xOf(roles.tmin) - 10, ay); ctx.lineTo(W - 20, ay); ctx.stroke();
      ctx.font = `400 9.5px ${mono()}`; ctx.fillStyle = P.dim; ctx.textAlign = "center"; ctx.textBaseline = "bottom";
      const step = W < 500 ? 4 : 2;
      for (let y = roles.tmin; y <= roles.tmax; y++) {
        const x = roles.xOf(y);
        ctx.globalAlpha = .5; ctx.beginPath(); ctx.moveTo(x, ay); ctx.lineTo(x, ay - (y % step ? 3 : 6)); ctx.stroke();
        if (y % step === 0) { ctx.globalAlpha = .8; ctx.fillText(y, x, ay - 8); }
      }
      // edges
      ctx.lineWidth = 1;
      for (const s of stars) {
        const [sx, sy] = pos(s);
        for (const r of s.links) {
          const on = lit && lit.has(s) && lit.has(r);
          ctx.strokeStyle = on ? P.gold : P.edge; ctx.globalAlpha = on ? .85 : lit ? .05 : (P.bg ? .22 : .13);
          ctx.lineWidth = on ? 1.4 : 1;
          ctx.beginPath(); ctx.moveTo(sx, sy); ctx.lineTo(r.x, r.y); ctx.stroke();
        }
      }
      // roles
      ctx.textBaseline = "top"; ctx.font = `500 10px ${mono()}`;
      for (const r of roles) {
        const on = lit && lit.has(r);
        ctx.globalAlpha = lit && !on ? .35 : 1;
        ctx.fillStyle = on ? P.gold : P.star;
        if (on && !P.bg) { ctx.shadowColor = P.gold; ctx.shadowBlur = 14; }
        ctx.beginPath(); ctx.arc(r.x, r.y, r.r + (on ? 1.5 : 0), 0, 6.28); ctx.fill();
        ctx.shadowBlur = 0;
        if (!r.brief || (on && r.row < roles.maxRows)) {
          ctx.fillStyle = on ? P.gold : P.dim; ctx.textAlign = "left"; ctx.globalAlpha = lit && !on ? .35 : .95;
          ctx.fillText(r.label, r.lx, ay + 10 + r.row * 13);
        }
      }
      // stars + labels
      ctx.textBaseline = "middle";
      for (const s of stars) {
        const [x, y] = pos(s), on = lit && lit.has(s), dim = lit && !on;
        const tw = .75 + .25 * Math.sin(t * 1.3 + s.phase * 3);
        ctx.globalAlpha = dim ? .3 : 1;
        if (!P.bg) {
          const g = ctx.createRadialGradient(x, y, 0, x, y, s.r * (on ? 6 : 3.2));
          g.addColorStop(0, (on ? P.gold : P.star) + (on ? "cc" : "66")); g.addColorStop(1, (on ? P.gold : P.star) + "00");
          ctx.fillStyle = g; ctx.beginPath(); ctx.arc(x, y, s.r * (on ? 6 : 3.2), 0, 6.28); ctx.fill();
        }
        ctx.fillStyle = on ? P.gold : s.level >= 5 ? "#fff" : P.star;
        if (P.bg) ctx.fillStyle = on ? P.gold : P.star;
        ctx.beginPath(); ctx.arc(x, y, s.r * (on ? 1.3 : tw), 0, 6.28); ctx.fill();
        if (labelled.has(s) || on) {
          ctx.font = `${on || s.level >= 5 ? 500 : 400} ${on ? 12 : 11}px ${sans()}`;
          ctx.textAlign = s.labelLeft ? "right" : "left";
          const lx = s.labelLeft ? x - s.r - 6 : x + s.r + 6;
          ctx.lineWidth = 3; ctx.strokeStyle = P.halo; ctx.globalAlpha = dim ? .2 : .9; ctx.lineJoin = "round";
          ctx.strokeText(s.name, lx, y);
          ctx.fillStyle = on ? P.gold : s.level >= 5 ? P.ink : P.dim;
          ctx.fillText(s.name, lx, y);
        }
      }
      ctx.globalAlpha = 1;
    }

    function frame(now) {
      if (!running) return;
      draw(now);
      requestAnimationFrame(frame);
    }
    function setRunning(on) {
      on = on && !calm && !document.hidden && !page.hidden && W > 0;
      if (on === running) return;
      running = on;
      if (on) requestAnimationFrame(frame); else draw(performance.now());
    }
    let visible = false;
    new IntersectionObserver(([e]) => { visible = e.isIntersecting; setRunning(visible); }).observe(sky);
    document.addEventListener("visibilitychange", () => setRunning(visible));

    function resize() {
      const box = sky.getBoundingClientRect();
      if (!box.width || !box.height) return;
      dpr = Math.min(devicePixelRatio || 1, 2);
      W = box.width; H = box.height;
      canvas.width = Math.round(W * dpr); canvas.height = Math.round(H * dpr);
      layout();
      draw(performance.now());
      setRunning(visible);
    }
    new ResizeObserver(resize).observe(sky);

    // Hover and click
    const near = (px, py) => {
      let best = null, bd = 16;
      for (const n of [...stars, ...roles]) { const d = Math.hypot(n.x - px, n.y - py); if (d < bd) { bd = d; best = n; } }
      return best;
    };
    function setHot(n) {
      if (n === hot) return;
      for (const r of roles) r.el.classList.toggle("cx-hot", !!n && (n === r || (n.links && n.links.includes(r))));
      hot = n;
      sky.toggleAttribute("data-hot", !!n);
      if (!running) draw(performance.now());
    }
    sky.addEventListener("pointermove", (e) => { const b = canvas.getBoundingClientRect(); setHot(near(e.clientX - b.left, e.clientY - b.top)); });
    sky.addEventListener("pointerleave", () => setHot(null));
    sky.addEventListener("click", (e) => {
      const b = canvas.getBoundingClientRect(), n = near(e.clientX - b.left, e.clientY - b.top);
      if (!n) return;
      if (n.links) { if (n.chip.dataset.hits) n.chip.click(); setHot(null); }  // the filter scrolls the page away
      else n.el.scrollIntoView({ behavior: calm ? "auto" : "smooth", block: "center" });
    });

    // Print: redraw in ink on white; back to night afterwards.
    const repaint = (p) => { palette = p; resize(); };  // print changes the box, so lay out again, now
    addEventListener("beforeprint", () => repaint(PAPER));
    addEventListener("afterprint", () => repaint(NIGHT));
    matchMedia("print").addEventListener("change", (e) => repaint(e.matches ? PAPER : NIGHT));
  }

  document.querySelectorAll(".page").forEach(build);
})();
