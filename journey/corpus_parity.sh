#!/usr/bin/env bash
# Compare this library's output against the real ExcelDataReader across its own
# test corpus, and refresh testdata/corpus_baseline.tsv.
#
# Needs: .NET SDK, a Go toolchain, and a checkout of ExcelDataReader.
# See journey/chapters/07-the-corpus.md for what the verdicts mean.
#
#   ./journey/corpus_parity.sh /path/to/ExcelDataReader/src/TestData
set -euo pipefail

CORPUS="${1:?usage: corpus_parity.sh <ExcelDataReader/src/TestData>}"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

echo "building the two dumpers..."
# The C# side: emits one line per non-empty cell, numbers as IEEE 754 bit
# patterns so C#/Go float formatting cannot manufacture differences.
mkdir -p "$WORK/edrdump" && cd "$WORK/edrdump"
dotnet new console -o . --force >/dev/null
dotnet add package ExcelDataReader >/dev/null
dotnet add package System.Text.Encoding.CodePages >/dev/null
cp "$ROOT/journey/edrdump.cs" Program.cs
dotnet build -c Release -v q --nologo >/dev/null
EDR="$(ls -d "$WORK"/edrdump/bin/Release/net*/edrdump.dll | head -1)"

cd "$CORPUS"
: > "$WORK/results.tsv"
for f in $(find . -maxdepth 2 \( -iname '*.xls' -o -iname '*.xlsx' \) | sort); do
  name="${f#./}"
  # TZ=UTC is load-bearing. ExcelDataReader converts OOXML strict ISO dates
  # through the machine's local timezone, so the same file decodes differently
  # on different hosts. Pinning it compares values rather than locales.
  TZ=UTC timeout 60 dotnet "$EDR" "$f" > "$WORK/e.txt" 2>/dev/null && erc=0 || erc=$?
  timeout 60 go run "$ROOT/journey/godump" "$f" > "$WORK/g.txt" 2>/dev/null && grc=0 || grc=$?

  if   [ $erc -ne 0 ] && [ $grc -ne 0 ]; then v=BOTH_REJECT
  elif [ $erc -ne 0 ]; then v=ONLY_GO_READS
  elif [ $grc -ne 0 ]; then v=ONLY_EDR_READS
  elif cmp -s "$WORK/e.txt" "$WORK/g.txt"; then v=IDENTICAL
  else v=DIFFER; fi
  printf '%s\t%s\n' "$v" "$name" >> "$WORK/results.tsv"
done

echo
echo "=== verdicts ==="
cut -f1 "$WORK/results.tsv" | sort | uniq -c | sort -rn
echo
echo "Verdicts are what the committed baseline records. Refresh the hashes with:"
echo "  EXCEL_CORPUS=$CORPUS EXCEL_CORPUS_REGEN=1 go test -run TestCorpusParity ."
echo "and update the counts in TestCorpusBaselineShape if they moved."
