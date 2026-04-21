#!/bin/bash
# ============================================================================
# Build script — sestavi combined.gs z 6 modulu pro Google Ads Scripts UI.
#
# Pouziti:
#   cd /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/
#   bash build-combined.sh
#
# Nebo z kteregokoli mista:
#   bash /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/build-combined.sh
#
# Modul order (musi byt zachovano kvuli dependencies):
#   1. utils.gs       — zakladni helpers, neni na nikom zavisly
#   2. config.gs      — pouziva Utils
#   3. data.gs        — pouziva Utils
#   4. classifier.gs  — pouziva Utils
#   5. output.gs      — pouziva Utils
#   6. main.gs        — pouziva vse vyse + CONFIG object + setupOutputSheet()
# ============================================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTPUT="$SCRIPT_DIR/combined.gs"
TIMESTAMP=$(date "+%Y-%m-%d %H:%M")

cat > "$OUTPUT" << HEADER_EOF
/**
 * ============================================================================
 * Shopping/PMAX Loser Detector — COMBINED FILE (pro Google Ads Scripts UI)
 * ============================================================================
 *
 * Tento soubor je automaticky sestaveny z 7 modulu:
 *   _config (TVOJE KONFIGURACE), utils, config (validation), data,
 *   classifier, output, main.
 *
 * Google Ads Scripts ma jediny editor — nelze pridat multi-file structure.
 * Proto celou kombinaci paste do jednoho skriptu v UI.
 *
 * ⬇⬇⬇ NAHORE SOUBORU JE CONFIG — UPRAVUJ TAM ⬇⬇⬇
 *
 * PRVNI RUN workflow:
 *   1. Paste combined.gs do Google Ads Scripts editoru
 *   2. Nech CONFIG.outputSheetId prazdny ('')
 *   3. Pust main() (Nahled/Preview)
 *   4. V logu najdi ID a URL vytvoreneho sheetu
 *   5. Paste ID do CONFIG.outputSheetId a nastav targetPnoPct, atd.
 *   6. Pust znovu s dryRun=true pro test
 *   7. Po verifikaci prepni dryRun=false a pust ostry run
 *
 * DOKUMENTACE:
 *   /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/README.md
 *   /Users/matousnovy/Documents/PPC/skripty/shopping-pmax-loser-detector/DEPLOYMENT-GUIDE.md
 *
 * VYGENEROVANO: $TIMESTAMP
 * BUILD SCRIPT: build-combined.sh
 * ============================================================================
 */

HEADER_EOF

# Append files in dependency order
# _config.gs je zamerne PRVNI (underscore = viditelne nahore u usera)
for MODULE in _config utils config data classifier output main; do
  SRC="$SCRIPT_DIR/${MODULE}.gs"
  if [ ! -f "$SRC" ]; then
    echo "ERROR: Missing source $SRC" >&2
    exit 1
  fi
  echo "" >> "$OUTPUT"
  cat "$SRC" >> "$OUTPUT"
done

LINES=$(wc -l < "$OUTPUT")
echo "Built: $OUTPUT ($LINES lines)"
