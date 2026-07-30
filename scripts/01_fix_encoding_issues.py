"""
Fixes raw_data.csv rows that were torn in two by stray CR/LF characters
embedded (unquoted) inside the `vendoraccountid` field.

Approach: read the file as one blob, split into logical records on the
real record delimiter (\\r\\n), then within each record strip any leftover
bare \\r or \\n characters from field values (these are the stray bytes
that caused naive parsers to split rows early). This does NOT touch
any other data.
"""
import csv
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "raw" / "raw_data.csv"
DST = ROOT / "data" / "processed" / "raw_data_fixed.csv"
ENCODING = "cp1252"

with open(SRC, encoding=ENCODING, newline="") as f:
    raw = f.read()

# Split into logical lines on the file's real record delimiter only.
lines = [ln for ln in raw.split("\r\n") if ln.strip() != ""]

reader_rows = []
header = None
bad_rows = []

for i, line in enumerate(lines, start=1):
    # Strip any stray bare \r or \n that survived within the line
    # (these are the characters that caused rows to split early).
    cleaned = line.replace("\r", "").replace("\n", "")
    fields = next(csv.reader([cleaned]))
    if header is None:
        header = fields
        continue
    if len(fields) != len(header):
        bad_rows.append((i, len(fields), fields))
        continue
    reader_rows.append(fields)

print(f"Header columns: {len(header)}")
print(f"Clean data rows: {len(reader_rows)}")
print(f"Rows still misaligned after fix: {len(bad_rows)}")
for i, n, fields in bad_rows[:20]:
    print(f"  line {i}: {n} fields -> {fields}")

with open(DST, "w", encoding=ENCODING, newline="") as f:
    writer = csv.writer(f, lineterminator="\r\n")
    writer.writerow(header)
    writer.writerows(reader_rows)

print(f"\nWrote fixed file to: {DST}")
