// Aurora interactions, shared by the themes that link to this file. Plain DOM, no dependencies; every page (language) is wired the same way.
(() => {
  const pages = [...document.querySelectorAll(".page")];
  const calm = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

  // "Postgres/MySQL" matches either half; word-ish boundaries so "AI" does not hit "maintain".
  const matcher = (term, aka = "") => new RegExp(
    [...term.split("/"), ...aka.split("|").filter(Boolean)]
      .map((t) => `(?<![\\w-])${esc(t.trim())}(?![\\w-])`).join("|"), "i");

  function wire(page) {
    const ui = JSON.parse(page.dataset.ui || "{}");
    // Looked up on use: pagination into sheets (slate.js) replaces the sections and splits a role into a head
    // item and ".cont" items on later sheets. A role is its parts.
    const sections = () => [...page.querySelectorAll(".experience")];
    const roles = () => {
      const out = [];
      for (const it of page.querySelectorAll(".experience .item")) {
        if (it.classList.contains("cont") && out.length) out.at(-1).push(it);
        else out.push([it]);
      }
      return out;
    };
    const chips = [...page.querySelectorAll("aside .kv[style], .tools li")];
    // Join per element: adjacent tool tags have no whitespace between them in textContent.
    const textOf = (role) => role.flatMap((part) => [...part.querySelectorAll("h3, li, p")].map((e) => e.textContent)).join(" \n ");
    const termOf = (el) => (el.matches(".kv") ? el.firstElementChild : el).textContent.trim();

    // Only offer chips that match at least one role.
    const akaOf = (el) => el.dataset.aka || "";
    for (const chip of chips) {
      const re = matcher(termOf(chip), akaOf(chip));
      const hits = roles().filter((r) => re.test(textOf(r))).length;
      if (hits) chip.dataset.hits = hits;
      else chip.title = ui.no_match || "";
    }

    const pill = document.createElement("div");
    pill.className = "filter-pill";
    pill.setAttribute("aria-hidden", "true");
    pill.setAttribute("role", "status");
    pill.innerHTML = `<span></span><button type="button">${ui.clear || "Clear"}</button>`;
    page.append(pill);

    let current = null;
    function unmark() {
      for (const exp of sections()) {
        for (const m of exp.querySelectorAll("mark")) m.replaceWith(...m.childNodes);
        exp.normalize();
      }
    }
    function mark(re) {  // highlight the term in bullet text
      for (const li of page.querySelectorAll(".experience .item.match > ul:not(.tools) li")) {
        const walker = document.createTreeWalker(li, NodeFilter.SHOW_TEXT);
        const nodes = [];
        while (walker.nextNode()) nodes.push(walker.currentNode);
        for (const node of nodes) {
          const m = node.data.match(re);
          if (!m) continue;
          const hit = node.splitText(m.index);
          hit.splitText(m[0].length);
          const el = document.createElement("mark");
          hit.replaceWith(el);
          el.append(hit);
        }
      }
    }
    function filter(term, aka = "") {
      current = term === current ? null : term;
      unmark();
      for (const c of chips) c.classList.toggle("active", current !== null && termOf(c) === current);
      for (const exp of sections()) exp.classList.toggle("filtering", current !== null);
      if (current === null) {
        for (const role of roles()) for (const part of role) part.classList.remove("match");
        pill.setAttribute("aria-hidden", "true");
        return;
      }
      const re = matcher(current, aka);
      let n = 0;
      for (const role of roles()) {
        const hit = re.test(textOf(role));
        for (const part of role) part.classList.toggle("match", hit);
        n += hit;
      }
      mark(re);
      pill.firstElementChild.innerHTML = `${ui.filtering || "Roles using"} <b></b> · ${n}`;
      pill.querySelector("b").textContent = current;
      pill.setAttribute("aria-hidden", "false");
      const first = page.querySelector(".experience .item.match");
      if (first && first.getBoundingClientRect().top > innerHeight) first.scrollIntoView({ behavior: calm ? "auto" : "smooth", block: "center" });
    }
    for (const chip of chips.filter((c) => c.dataset.hits)) {
      chip.tabIndex = 0;
      chip.setAttribute("role", "button");
      chip.addEventListener("click", () => filter(termOf(chip), akaOf(chip)));
      chip.addEventListener("keydown", (e) => (e.key === "Enter" || e.key === " ") && (e.preventDefault(), filter(termOf(chip), akaOf(chip))));
    }
    pill.querySelector("button").addEventListener("click", () => filter(current));
    addEventListener("keydown", (e) => e.key === "Escape" && current !== null && !page.hidden && filter(current));

    // Earlier roles: one toggle to show their details.
    const firstBrief = page.querySelector(".experience .brief[data-earlier]");
    if (firstBrief) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "details-toggle";
      btn.setAttribute("aria-expanded", "false");
      btn.textContent = ui.show_details || "Show details";
      btn.addEventListener("click", () => {
        const open = !btn.closest(".experience").classList.contains("expanded");
        for (const exp of sections()) exp.classList.toggle("expanded", open);
        btn.setAttribute("aria-expanded", open);
        btn.textContent = open ? ui.hide_details || "Hide details" : ui.show_details || "Show details";
      });
      firstBrief.before(btn);
    }

    // Headline numbers count up once, the first time they are seen.
    const stats = [...page.querySelectorAll(".highlights b")];
    const io = new IntersectionObserver((entries) => {
      for (const { target, isIntersecting } of entries) {
        if (!isIntersecting) continue;
        io.unobserve(target);
        const m = target.textContent.match(/^(\D*)(\d+(?:[.,]\d+)?)(.*)$/);
        if (!m || calm) continue;
        const [, pre, num, post] = m;
        target.dataset.final = target.textContent;
        const end = parseFloat(num.replace(",", ".")), dec = (num.split(/[.,]/)[1] || "").length;
        const t0 = performance.now(), dur = 1100;
        const step = (t) => {
          const k = Math.min(1, (t - t0) / dur), v = end * (1 - Math.pow(1 - k, 3));
          target.textContent = pre + v.toFixed(dec).replace(".", num.includes(",") ? "," : ".") + post;
          if (target.dataset.final === undefined) return;  // finished early (print)
          if (k < 1) requestAnimationFrame(step);
          else delete target.dataset.final;
        };
        requestAnimationFrame(step);
      }
    });
    stats.forEach((s) => io.observe(s));
  }

  pages.forEach(wire);

  // Printing mid count-up would put partial numbers on paper: jump to the final values.
  const finish = () => {
    for (const b of document.querySelectorAll(".highlights b[data-final]")) {
      b.textContent = b.dataset.final;
      delete b.dataset.final;
    }
  };
  addEventListener("beforeprint", finish);
  matchMedia("print").addEventListener("change", (e) => e.matches && finish());
})();
