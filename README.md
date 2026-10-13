# Resume

One YAML per language, built to DOCX and to self-contained web pages that print to A4.

| Path | What |
|---|---|
| `cv/` | `cv.yaml`: every language; translated text sits under `en:`/`sv:` keys, and a missing language is an error |
| `assets/` | Photo; `originals/` holds the original resume downloads (git-ignored) |
| `templates/` | DOCX templates (regenerate with `resume template <name>`) |
| `themes/` | Web themes: `<name>.css`, plus optional `<name>.js` / `<name>.*.js` |
| `src/resume/` | The `resume` CLI: `build` (page, PDFs, .docx), `template` |
| `build/` | Output, git-ignored |

```bash
uv run resume build
```

```bash
uv run pytest
```

```bash
uv run isort . && uv run ruff format && uv run ruff check && uv run ty check
```
