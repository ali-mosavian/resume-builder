#!/usr/bin/env bash
# The README's showcase: per theme, the first A4 of the web page, the PDF and the .docx (as Word lays it out).
# Usage: tools/showcase.sh [THEME...]   ->  docs/showcase/THEME.jpg (default: every theme)
set -euo pipefail
cd "$(dirname "$0")/.."
H=700  # px: every panel is scaled to this height
mkdir -p docs/showcase
themes=("$@")
[ ${#themes[@]} -gt 0 ] || themes=($(ls themes/*.css | xargs -n1 basename | sed 's/\.css$//'))
for theme in "${themes[@]}"; do
  out=$(tools/compare.sh "$theme" en | tail -1)
  name=$(cd "$out/site" && ls *-en.pdf | sed 's/\.pdf$//')
  # The screen shot is 640 px wide (half of 1280): its top A4-shaped slice.
  magick "$out/screen.png" -crop 640x905+0+0 +repage -resize "x$H" "$out/web.png"
  pdftoppm -png -f 1 -l 1 -scale-to-y "$H" -scale-to-x -1 "$out/site/$name.pdf" "$out/pdf1"
  pdftoppm -png -f 1 -l 1 -scale-to-y "$H" -scale-to-x -1 "$out/word.pdf" "$out/word1"
  montage -font assets/fonts/Inter-Regular.ttf -pointsize 16 -background white -bordercolor "#d0d7de" -border 1 \
    -label web "$out/web.png" -label pdf "$out"/pdf1*.png -label docx "$out"/word1*.png \
    -tile 3x1 -geometry +8+8 -quality 82 "docs/showcase/$theme.jpg"
  echo "docs/showcase/$theme.jpg"
done
