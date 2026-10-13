// Language switch, inlined by build_html.py when the page holds more than one language.
// Shows one .page at a time, keeps the choice in the URL hash and fires "langchange" for theme scripts.
(() => {
  const pages = [...document.querySelectorAll(".page")];
  const show = (lang) => {
    for (const p of pages) p.hidden = p.lang !== lang;
    document.documentElement.lang = lang;
    for (const b of document.querySelectorAll(".lang-switch button")) b.setAttribute("aria-pressed", b.dataset.lang === lang);
    history.replaceState(null, "", `#${lang}`);
    dispatchEvent(new Event("langchange"));
  };
  for (const p of pages) {
    const sw = document.createElement("div");
    sw.className = "lang-switch";
    sw.setAttribute("role", "group");
    for (const q of pages) {
      const b = document.createElement("button");
      b.type = "button";
      b.dataset.lang = q.lang;
      b.textContent = JSON.parse(q.dataset.ui || "{}").lang_name || q.lang.toUpperCase();
      b.addEventListener("click", () => show(q.lang));
      sw.append(b);
    }
    p.querySelector(".contact").prepend(sw);
  }
  const langs = pages.map((p) => p.lang);
  const wanted = location.hash.slice(1);
  const browser = (navigator.language || "").slice(0, 2);
  show(langs.includes(wanted) ? wanted : langs.includes(browser) ? browser : langs[0]);
})();
