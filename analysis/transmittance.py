#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import openpyxl

TARGET_NM = 550.0
DEFAULT_OUTPUT = "transmittance_550nm.dat"
DEFAULT_OUTPUT_DIR = "results"
DEFAULT_SAMPLE_TABLE = "samples.xlsx"
SEARCH_LEVELS = 4
MISSING = "—"

COLUMNS = ["Sample", "Alias", "Layers", "Additive", "SolutionAge",
           "T550_percent", "Date", "SourceFile"]
COL = {name: i for i, name in enumerate(COLUMNS)}


def log(msg: str = "") -> None:
    print(msg, flush=True)


def warn(msg: str) -> None:
    print(f"  ! {msg}", flush=True)


def to_float(text, default=None):
    if text is None:
        return default
    if isinstance(text, (int, float)):
        return float(text)
    s = str(text).strip()
    if not s:
        return default
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return default


NUMBER_PREFIX = "V"


@dataclass
class SampleConfig:
    id: str = ""
    coatings: str = ""
    alias: str = ""
    additive: str = ""
    age: str = ""


TABLE_FIELDS = {
    "id":       ("sample", "νούμερο", "νουμερο", "number"),
    "alias":    ("alias",),
    "coatings": ("layers", "spin coatings"),
    "additive": ("additive", "αιθαλολαμίνη", "αιθαλολαμινη"),
    "age":      ("solution age", "ηλικία διαλύματος", "ηλικια διαλυματος"),
    "area":     ("area", "εμβαδόν", "εμβαδον"),
}


def id_key(text) -> str:
    key = re.sub(r"\s+", "", str(text or "")).upper()
    return re.sub(r"(?<![0-9])0+(?=\d)", "", key)


def make_id(value) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip() if value is not None else ""
    if re.fullmatch(r"\d+", text):
        return f"{NUMBER_PREFIX}{int(text)}"
    return text


def _norm_header(text) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip().lower()


def _text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_sample_table(path: Path) -> dict[str, SampleConfig]:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb.worksheets[0]
        header_row, columns = None, {}
        for row in range(1, min(ws.max_row, 20) + 1):
            found = {}
            for col in range(1, ws.max_column + 1):
                norm = _norm_header(ws.cell(row, col).value)
                if not norm:
                    continue
                for name, keys in TABLE_FIELDS.items():
                    if name not in found and any(k in norm for k in keys):
                        found[name] = col
                        break
            if "id" in found and "area" in found:
                header_row, columns = row, found
                break
        if header_row is None:
            raise ValueError(
                f"No header row with 'Sample' and 'Area' columns found in {path.name}."
            )

        table: dict[str, SampleConfig] = {}
        for row in range(header_row + 1, ws.max_row + 1):
            sample_id = make_id(ws.cell(row, columns["id"]).value)
            if not sample_id:
                continue

            def cell(name: str):
                col = columns.get(name)
                return ws.cell(row, col).value if col else None

            table[id_key(sample_id)] = SampleConfig(
                id=sample_id,
                coatings=_text(cell("coatings")),
                alias=_text(cell("alias")),
                additive=_text(cell("additive")),
                age=_text(cell("age")),
            )
        return table
    finally:
        wb.close()


def lookup_sample(token: str, table: dict[str, SampleConfig]) -> SampleConfig | None:
    key = id_key(token)
    if not key:
        return None
    if key in table:
        return table[key]
    aliases = {id_key(c.alias): c for c in table.values() if c.alias}
    if key in aliases:
        return aliases[key]
    if key.isdigit():
        return table.get(id_key(f"{NUMBER_PREFIX}{key}"))
    return None


def sample_in_name(stem: str, table: dict[str, SampleConfig] | None) -> str | None:
    found = [f"{m.group(1).upper()}{int(m.group(2))}" for m in SAMPLE_RE.finditer(stem)
             if m.group(1).upper() != "SC"]
    known = [s for s in found if table and lookup_sample(s, table)]
    return (known or found or [None])[0]


def search_up(name: str, starts: list[Path], want_dir: bool = False) -> Path | None:
    seen = set()
    for start in starts:
        for folder in [start, *list(start.parents)[:SEARCH_LEVELS]]:
            candidate = folder / name
            if candidate in seen:
                continue
            seen.add(candidate)
            if candidate.is_dir() if want_dir else candidate.exists():
                return candidate
    return None


SAMPLE_RE = re.compile(r"(?<![A-Za-z0-9])([A-Za-z]{1,3})\s*0*(\d+)(?!\d)")
COATINGS_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:(\d+)\s*SC|SC\s*(\d+))(?![A-Za-z0-9])", re.IGNORECASE)
DATE_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")


@dataclass
class Spectrum:
    sample: str
    coatings: str
    date: str
    value: float
    interpolated: bool
    source: Path


def parse_header_date(lines: list[str]) -> str | None:
    if len(lines) < 4:
        return None
    for fmt in ("%y/%m/%d", "%Y/%m/%d", "%d/%m/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(lines[3].strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return None


def value_at(pairs: list[tuple[float, float]], target: float) -> tuple[float, bool]:
    for x, y in pairs:
        if abs(x - target) < 1e-6:
            return y, False
    ordered = sorted(pairs)
    if not ordered or target < ordered[0][0] or target > ordered[-1][0]:
        raise ValueError(f"{target:g} nm is outside the measured range")
    for (x0, y0), (x1, y1) in zip(ordered, ordered[1:]):
        if x0 <= target <= x1:
            if x1 == x0:
                return y0, False
            weight = (target - x0) / (x1 - x0)
            return y0 + weight * (y1 - y0), True
    raise ValueError(f"no value found at {target:g} nm")


def read_spectrum(path: Path, table: dict[str, SampleConfig] | None = None) -> Spectrum:
    with path.open(encoding="latin-1") as fh:
        lines = [line.rstrip("\n") for line in fh]

    try:
        gr = lines.index("#GR")
        data = lines.index("#DATA")
    except ValueError:
        raise ValueError("not a PerkinElmer file (#GR or #DATA is missing)")

    x_unit = lines[gr + 1].strip() if gr + 1 < len(lines) else ""
    y_unit = lines[gr + 2].strip() if gr + 2 < len(lines) else ""
    if x_unit.lower() != "nm":
        raise ValueError(f"the x axis is '{x_unit}', not nm")
    if y_unit.upper() not in ("%T", "T"):
        warn(f"{path.name}: the y axis is '{y_unit}', not %T — "
             f"the value is not a transmittance.")

    pairs: list[tuple[float, float]] = []
    for line in lines[data + 1:]:
        parts = line.split("\t")
        if len(parts) != 2:
            continue
        x, y = to_float(parts[0]), to_float(parts[1])
        if x is not None and y is not None:
            pairs.append((x, y))
    if not pairs:
        raise ValueError("no data pairs found after #DATA")

    value, interpolated = value_at(pairs, TARGET_NM)

    stem = path.name
    sample = sample_in_name(stem, table)
    if not sample:
        raise ValueError("no sample name (e.g. S12) found in the file name")

    cm = COATINGS_RE.search(stem)
    coatings = next((g for g in (cm.groups() if cm else ()) if g), "")

    date = parse_header_date(lines)
    dm = DATE_RE.search(stem)
    from_name = None
    if dm:
        try:
            from_name = datetime.strptime(dm.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            from_name = None
    if date and from_name and date != from_name:
        warn(f"{path.name}: the file name says {from_name} but the file says {date}. "
             f"Using the date stored in the file.")
    if not date:
        date = from_name

    return Spectrum(sample=sample, coatings=coatings, date=date or MISSING,
                    value=value, interpolated=interpolated, source=path)


def row_key(row: list[str]) -> tuple:
    return (row[COL["Sample"]], row[COL["Layers"]], row[COL["Date"]])


def sort_key(row: list[str]) -> tuple:
    m = re.match(r"([A-Z]+)(\d+)", row[COL["Sample"]])
    prefix, number = (m.group(1), int(m.group(2))) if m else (row[COL["Sample"]], 0)
    return (prefix, number, to_float(row[COL["Layers"]], 0.0) or 0.0,
            row[COL["Date"]])


def read_dat(path: Path) -> list[list[str]]:
    if not path.exists():
        return []
    rows = []
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        parts = [p.strip() for p in re.split(r"\s{2,}|\t", line.strip())]
        if len(parts) == len(COLUMNS):
            rows.append(parts)
        else:
            warn(f"{path.name}: ignoring a line with {len(parts)} fields instead "
                 f"of {len(COLUMNS)}.")
    return rows


GAP = "  "


def write_dat(path: Path, rows: list[list[str]]) -> None:
    widths = [max([len(COLUMNS[i])] + [len(r[i]) for r in rows])
              for i in range(len(COLUMNS))]
    head = [COLUMNS[0].ljust(max(widths[0] - 2, 0))]
    head += [c.ljust(w) for c, w in zip(COLUMNS[1:], widths[1:])]
    body = [GAP.join(v.ljust(w) for v, w in zip(row, widths)).rstrip() for row in rows]
    path.write_text(
        "\n".join([
            f"# Transmittance at {TARGET_NM:g} nm (%T) from PerkinElmer spectra",
            f"# Updated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "# Columns are separated by 2+ spaces; the file is extended on every run.",
            "# " + GAP.join(head).rstrip(),
            *body,
            "",
        ]),
        encoding="utf-8",
    )


def collect_files(inputs: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in inputs:
        p = Path(raw).expanduser()
        if p.is_dir():
            files.extend(sorted(p.rglob("*.asc")))
        elif p.is_file():
            files.append(p)
        else:
            warn(f"Not found: {raw}")
    seen, unique = set(), []
    for f in files:
        key = f.resolve()
        if key not in seen:
            seen.add(key)
            unique.append(f)
    return unique


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=f"Transmittance at {TARGET_NM:g} nm, collected into a cumulative .dat file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python transmittance.py data/Transmittance
  python transmittance.py data/Transmittance/S12
  python transmittance.py data/Transmittance --out "T550.dat"
""")
    parser.add_argument("inputs", nargs="+", help="Folder with .asc files (searched recursively), or single files.")
    parser.add_argument("-o", "--out", default=DEFAULT_OUTPUT, help="Name of the .dat file.")
    parser.add_argument("--outdir", help=f"Output folder (default: '{DEFAULT_OUTPUT_DIR}').")
    parser.add_argument("--samples", default=DEFAULT_SAMPLE_TABLE,
                        help=f"Sample table (default: {DEFAULT_SAMPLE_TABLE}).")
    parser.add_argument("--fresh", action="store_true",
                        help="Rewrite the .dat from scratch instead of extending it.")
    args = parser.parse_args(argv)

    files = collect_files(args.inputs)
    if not files:
        log("No .asc files found.")
        return 1

    script_dir = Path(__file__).resolve().parent
    starts = [Path.cwd(), script_dir] + [f.parent for f in files]

    table_path = Path(args.samples).expanduser()
    if not table_path.is_absolute():
        table_path = search_up(args.samples, starts) or (script_dir / args.samples)
    table = read_sample_table(table_path) if table_path.exists() else {}
    if table:
        log(f"Sample table: {table_path.name}  ({len(table)} samples)")
    else:
        warn(f"Sample table not found ({args.samples}) — "
             f"additive and solution age will be left blank.")

    out_dir = Path(args.outdir).expanduser() if args.outdir else (
        search_up(DEFAULT_OUTPUT_DIR, starts, want_dir=True)
        or script_dir.parent / DEFAULT_OUTPUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / Path(args.out).name
    if out_path.suffix.lower() != ".dat":
        out_path = out_path.with_suffix(".dat")

    existing = [] if args.fresh else read_dat(out_path)
    rows = {row_key(r): r for r in existing}
    before = len(rows)

    log(f"Files to process: {len(files)}")
    added = updated = failed = 0
    for path in files:
        try:
            spec = read_spectrum(path, table)
        except Exception as exc:
            warn(f"Skipping {path.name}: {exc}")
            failed += 1
            continue

        cfg = lookup_sample(spec.sample, table)
        coatings = spec.coatings or (cfg.coatings if cfg else "")
        if cfg and spec.coatings and cfg.coatings and spec.coatings != cfg.coatings:
            log(f"  · {spec.sample}: {spec.coatings} layers when measured "
                f"(the sample table says {cfg.coatings})")

        sample = cfg.id if cfg and cfg.id else spec.sample
        if cfg is None:
            warn(f"{spec.sample}: not in the sample table — "
                 f"its properties are left blank.")
        row = [
            sample,
            (cfg.alias if cfg and cfg.alias else MISSING),
            coatings or MISSING,
            (cfg.additive if cfg and cfg.additive else MISSING),
            (cfg.age if cfg and cfg.age else MISSING),
            f"{spec.value:.4f}",
            spec.date,
            spec.source.name,
        ]
        if spec.interpolated:
            warn(f"{path.name}: no data point exactly at {TARGET_NM:g} nm; "
                 f"used linear interpolation.")
        key = row_key(row)
        if key in rows:
            if rows[key] != row:
                updated += 1
        else:
            added += 1
        rows[key] = row

    if not rows:
        log("Nothing was written.")
        return 1

    ordered = sorted(rows.values(), key=sort_key)
    try:
        write_dat(out_path, ordered)
    except PermissionError:
        log(f"Could not write {out_path.name} — is it open in another program?")
        return 1

    log("")
    log(f"Saved: {out_path}")
    log(f"  {len(ordered)} rows in total "
        f"(previously {before}, new {added}, updated {updated}"
        + (f", failed {failed}" if failed else "") + ")")
    for row in ordered:
        log(f"    {row[COL['Sample']]:<5} {row[COL['Alias']]:<5} "
            f"{row[COL['Layers']] + 'L':<6} {row[COL['Additive']]:<5} "
            f"{row[COL['SolutionAge']]:<13} "
            f"T550 = {row[COL['T550_percent']]:>8} %   {row[COL['Date']]}")
    return 1 if failed else 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
