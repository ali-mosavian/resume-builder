// Slate: flow the CV into A4 sheets (210 × 297 mm), each with the slate sidebar.
// Without JS, or on narrow screens, the page stays one scrolling sheet; print then uses @page A4.
(() => {
  const wide = matchMedia("(min-width: 860px)");

  // Units are what may not be split; lists split between bullets. kind: "li" | "tools" | undefined.
  function mainUnits(main) {
    const out = [];
    for (const sec of [...main.children]) {
      for (const el of [...sec.children]) {
        if (!el.classList.contains("item")) { out.push({ sec, node: el, keep: el.matches("h2") }); continue; }
        const head = el.cloneNode(false);  // the item's class and data attributes, no children
        head.append(...[...el.children].filter((c) => !c.matches("ul")));
        const lists = el.querySelectorAll(":scope > ul");
        // Keep a role's header with its first bullet; a testimonial (no list) stands alone.
        out.push({ sec, item: el, node: head, head: true, keep: lists.length > 0 });
        for (const ul of lists) {
          const kind = ul.classList.contains("tools") ? "tools" : "li";
          if (kind === "tools") out.push({ sec, item: el, node: ul, kind });
          else for (const li of [...ul.children]) out.push({ sec, item: el, node: li, kind, list: ul });
        }
      }
    }
    return out;
  }

  // Group a "keep with next" unit (heading, role header) with the unit after it.
  function groups(units) {
    const out = [];
    for (let i = 0; i < units.length; i++) {
      const g = [units[i]];
      while (g.at(-1).keep && i + 1 < units.length) g.push(units[++i]);
      out.push(g);
    }
    return out;
  }

  function paginate(page) {
    if (page.classList.contains("paged")) return;
    const header = page.querySelector("header"), aside = page.querySelector("aside"), main = page.querySelector("main");
    // --header-top:1 (theme CSS): the header spans the first sheet above both columns instead of topping the sidebar.
    const top = getComputedStyle(page).getPropertyValue("--header-top").trim() === "1";
    const sideGroups = groups([...(top ? [] : [{ node: header }]), ...[...aside.children].map((n) => ({ node: n, keep: n.matches("h2") }))]);
    const mainGroups = groups(mainUnits(main));
    page.querySelector(".grid").remove();
    page.classList.add("paged");

    const sheets = [];
    const sheet = (i) => {
      while (sheets.length <= i) {
        const s = document.createElement("div");
        s.className = "sheet";
        const side = document.createElement("aside"), body = document.createElement("main");
        s.append(side, body);
        page.append(s);
        sheets.push({ side, body, sec: null, item: null, ul: null });
      }
      return sheets[i];
    };
    const over = (box) => box.scrollHeight > box.clientHeight + 1;

    function placeMain(unit, st) {  // returns the nodes it created, so a failed placement can be undone
      const made = [];
      if (!st.sec || st.sec.dataset.from !== unit.sec.className) {
        st.sec = unit.sec.cloneNode(false);
        st.sec.dataset.from = unit.sec.className;
        st.body.append(st.sec);
        st.item = st.ul = null;
        made.push(st.sec);
      }
      if (unit.item) {
        if (unit.head) {
          st.sec.append(unit.node);
          st.item = unit.node; st.itemOf = unit.item; st.ul = null;
          return [...made, unit.node];
        }
        if (st.itemOf !== unit.item) {  // role continues from the previous sheet
          st.item = unit.item.cloneNode(false);
          st.item.classList.add("cont");
          st.itemOf = unit.item;
          st.sec.append(st.item);
          st.ul = null;
          made.push(st.item);
        }
        if (unit.kind === "tools") { st.item.append(unit.node); return [...made, unit.node]; }
        if (!st.ul || st.ulOf !== unit.list) {
          st.ul = unit.list.cloneNode(false);
          if (st.item.classList.contains("cont") || st.item.querySelector(":scope > ul")) st.ul.classList.add("cont");
          st.ulOf = unit.list;
          st.item.append(st.ul);
          made.push(st.ul);
        }
        st.ul.append(unit.node);
        return [...made, unit.node];
      }
      st.sec.append(unit.node);
      st.item = st.ul = null; st.itemOf = null;
      return [...made, unit.node];
    }

    function fill(gs, box, place) {
      let i = 0;
      for (const g of gs) {
        for (;;) {
          const st = sheet(i), target = box(st), hadContent = target.childElementCount > 0;
          const made = g.flatMap((u) => place(u, st));
          // Fits, or nothing else is on this sheet (too tall for any sheet: let it overflow rather than loop).
          if (!over(target) || !hadContent) break;
          for (const n of made.reverse()) n.remove();
          st.sec = st.item = st.ul = st.itemOf = null;
          i++;
        }
      }
    }

    if (top) sheet(0).side.before(header);
    fill(sideGroups, (st) => st.side, (u, st) => (st.side.append(u.node), [u.node]));
    fill(mainGroups, (st) => st.body, placeMain);
  }

  const run = () => document.querySelectorAll(".page:not([hidden])").forEach(paginate);
  const go = () => (wide.matches ? run() : addEventListener("beforeprint", run, { once: true }));
  (document.fonts ? document.fonts.ready : Promise.resolve()).then(go);
})();
