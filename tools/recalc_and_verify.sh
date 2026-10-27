#!/usr/bin/env bash
# Builds the workbook under each scenario, recalculates it headless in LibreOffice, and checks
# every output against the Python reference model. Usage: tools/recalc_and_verify.sh
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
PY="${PYTHON:-$ROOT/.venv/bin/python}"
SOFFICE="${SOFFICE:-$(command -v soffice || echo /Applications/LibreOffice.app/Contents/MacOS/soffice)}"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

# Throwaway LibreOffice profile that always recalculates .xlsx files on load
mkdir -p "$WORK/profile/user"
cat > "$WORK/profile/user/registrymodifications.xcu" <<'XCU'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
XCU

cd "$ROOT/tools"
status=0
for s in 1 2 3; do
  "$PY" -c "import build_workbook as B, pathlib; B.build(pathlib.Path('$WORK/in/scenario_$s.xlsx'), $s)"
  "$SOFFICE" -env:UserInstallation="file://$WORK/profile" --headless --norestore \
    --convert-to xlsx:"Calc MS Excel 2007 XML" --outdir "$WORK/out" "$WORK/in/scenario_$s.xlsx" >/dev/null 2>&1
  "$PY" verify_workbook.py "$WORK/out/scenario_$s.xlsx" "$s" || status=1
done
"$PY" build_workbook.py >/dev/null   # leave the published workbook on the Base scenario
exit $status
