// Section menu under the headline numbers, as on Brittany Chiang's page: a line that grows and a label that
// brightens for the section being read. Built from each language's own headings.
(() => {
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
        (sel === "aside" ? skills : h).scrollIntoView({ behavior: "smooth", block: "start" });
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
