#!/bin/bash
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
REPO="$(pwd)"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  ASG LAB — Integrazione Hugo v15"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"

# Verifica root Hugo
if [ ! -f "config/_default/hugo.toml" ] && [ ! -f "config.toml" ]; then
  echo "ERRORE: esegui dalla root del repo Hugo"
  exit 1
fi

mkdir -p layouts/gestionale assets/gestionale content/gestionale

echo "1. Copio layout autonomo (bypassa Blowfish)..."
cp "$DIR/layout_gestionale.html" layouts/gestionale/list.html
cp "$DIR/layout_gestionale.html" layouts/gestionale/single.html
echo "   ✓ layouts/gestionale/list.html e single.html"

echo "2. Copio assets..."
cp "$DIR/dashboard.css"      assets/gestionale/dashboard.css
cp "$DIR/dashboard.js"       assets/gestionale/dashboard.js
cp "$DIR/dashboard_body.html" assets/gestionale/dashboard-body.html
echo "   ✓ assets/gestionale/ (3 file)"

echo "3. Aggiorno pagina content..."
cat > content/gestionale/index.it.md << 'MD'
---
title: "Gestionale"
layout: "gestionale"
draft: false
robots: "noindex"
---
MD
echo "   ✓ content/gestionale/index.it.md"

echo "4. Build Hugo..."
hugo --minify 2>&1 | tail -5

echo "5. Verifica output..."
COUNT=$(grep -c "sidebar\|nav-item" public/gestionale/index.html 2>/dev/null || echo 0)
echo "   Match dashboard: $COUNT"

if [ "$COUNT" -lt 3 ]; then
  echo "   ⚠ Pochi match — controlla public/gestionale/index.html"
fi

echo "6. Deploy su gh-pages..."
cd public
git add -A
git commit -m "deploy: dashboard v15 con layout autonomo Hugo"
git push origin gh-pages --force
cd ..

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  ✓ FATTO — https://asglab.it/gestionale/"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
