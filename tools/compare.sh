#!/usr/bin/env bash
# Render one theme three ways and put them side by side: web page on screen, its PDF print, and the .docx as
# Microsoft Word lays it out (LibreOffice differs from Word; only Word counts).
# Usage: tools/compare.sh THEME [LANG]   ->  build/compare/THEME/LANG/{screen.png, page-N.png (pdf | word), pages.txt}
set -euo pipefail
cd "$(dirname "$0")/.."
theme=$1 lang=${2:-en}
out=build/compare/$theme/$lang
rm -rf "$out" && mkdir -p "$out"
uv run -q resume build --out "$out/site" --theme "$theme" --template "templates/$theme.docx" >/dev/null
name=$(cd "$out/site" && ls *-"$lang".docx | sed 's/\.docx$//')

chrome="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
"$chrome" --headless=new --disable-gpu --hide-scrollbars --window-size=1280,14000 --virtual-time-budget=6000 \
  --screenshot="$out/screen-full.png" "file://$PWD/$out/site/index.html#$lang" 2>/dev/null
magick "$out/screen-full.png" -trim +repage -resize 50% "$out/screen.png" && rm "$out/screen-full.png"

# Word is one app with one sandbox grant (for build/word_test.docx): one conversion at a time.
lock=build/.word.lock
until mkdir "$lock" 2>/dev/null; do sleep 2; done
trap 'rmdir "$lock" 2>/dev/null || true' EXIT
osascript -e 'tell application "Microsoft Word" to close document "word_test.docx" saving no' >/dev/null 2>&1 || true
cp "$out/site/$name.docx" build/word_test.docx
rm -f build/word_test.pdf
timeout 120 osascript -e 'tell application "Microsoft Word"' -e "open (POSIX file \"$PWD/build/word_test.docx\")" -e 'delay 5' \
  -e "save as document \"word_test.docx\" file name \"$PWD/build/word_test.pdf\" file format format PDF" \
  -e 'close document "word_test.docx" saving no' -e 'end tell' >/dev/null
cp build/word_test.pdf "$out/word.pdf"
trap - EXIT  # released: the trap would otherwise delete the next run's lock
rmdir "$lock"

pdftoppm -png -r 50 "$out/site/$name.pdf" "$out/pdf"
pdftoppm -png -r 50 "$out/word.pdf" "$out/word"
pdf_pages=$(pdfinfo "$out/site/$name.pdf" | awk '/^Pages/{print $2}')
word_pages=$(pdfinfo "$out/word.pdf" | awk '/^Pages/{print $2}')
echo "pdf $pdf_pages pages, word $word_pages pages; $(pdfinfo "$out/word.pdf" | grep 'Page size')" | tee "$out/pages.txt"
for i in $(seq 1 $((pdf_pages > word_pages ? pdf_pages : word_pages))); do
  pdf=$(ls "$out"/pdf-*"$i".png 2>/dev/null | grep -E -- "-0*$i\.png$" || echo xc:white)
  word=$(ls "$out"/word-*"$i".png 2>/dev/null | grep -E -- "-0*$i\.png$" || echo xc:white)
  montage -font assets/fonts/Inter-Regular.ttf -label 'pdf' "$pdf" -label 'word' "$word" -tile 2x1 -geometry +6+6 "$out/page-$i.png"
done
rm -f "$out"/pdf-*.png "$out"/word-*.png
echo "$out"
