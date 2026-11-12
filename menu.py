#!/usr/bin/env python3
from __future__ import annotations

import importlib.util
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

ELECTROLYTE = "AlCl3"
DEFAULT_SCAN = "4"
STYLE = "report"
PNG = True

HERE = Path(__file__).resolve().parent
ANALYSIS = HERE / "analysis"
DATA = HERE / "data"
RESULTS = HERE / "results"
TABLE = HERE / "samples.xlsx"
TEMPLATE = HERE / "samples_template.xlsx"

CV_PY = ANALYSIS / "cv_analysis.py"
GCD_PY = ANALYSIS / "gcd_analysis.py"
LUM_PY = ANALYSIS / "transmittance.py"
THK_PY = ANALYSIS / "thickness_analysis.py"

CV_DIR = DATA / "CV"
GCD_DIR = DATA / "GCD"
LUM_DIR = DATA / "Transmittance"
THK_DIR = DATA / "Thickness"

NUMBER_PREFIX = "V"
ROUND_RE = re.compile(r"(?<![A-Za-z0-9])(?:\d+(?:st|nd|rd|th)|\d+c\d*)(?![A-Za-z])",
                      re.IGNORECASE)
CYCLES_PER_ROUND_RE = re.compile(r"(\d+)c(?![A-Za-z])", re.IGNORECASE)
RATE_RE = re.compile(r"([\d.,]+)\s*mV[/\s_]*(?:per|p)?[\s_]*s", re.IGNORECASE)
WINDOW_RE = re.compile(r"\s*([^=\s]+)\s*=\s*(-?[\d.]+)\s*:\s*(-?[\d.]+)\s*")
OPEN_RESULTS = "--no-open" not in sys.argv
MIN_PYTHON = (3, 10)
LIBRARIES = {"numpy": "numpy", "openpyxl": "openpyxl", "matplotlib": "matplotlib"}

HEADER_KEYS = {
    "id": ("sample", "νούμερο", "νουμερο", "number"),
    "alias": ("alias",),
    "layers": ("layers", "spin coatings"),
    "additive": ("additive", "αιθαλολαμίνη", "αιθαλολαμινη"),
    "age": ("solution age", "ηλικία διαλύματος", "ηλικια διαλυματος"),
    "area": ("area", "εμβαδόν", "εμβαδον"),
}


def ask(prompt: str) -> str | None:
    try:
        return input(prompt).replace("﻿", "").strip()
    except (EOFError, KeyboardInterrupt):
        return None


def open_path(path: Path) -> None:
    if not OPEN_RESULTS:
        return
    try:
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception:
        pass


def python_ok() -> bool:
    if sys.version_info >= MIN_PYTHON:
        return True
    need = ".".join(map(str, MIN_PYTHON))
    print(f"This is Python {sys.version.split()[0]}; version {need} or newer is needed.")
    print("Download the latest from https://www.python.org/downloads/")
    return False


def missing_libraries() -> list[str]:
    return [pip_name for module, pip_name in LIBRARIES.items()
            if importlib.util.find_spec(module) is None]


def libraries_ok() -> bool:
    missing = missing_libraries()
    if not missing:
        return True
    command = f'"{sys.executable}" -m pip install {" ".join(missing)}'
    print("\nSome libraries the analyses need are missing: " + ", ".join(missing))
    answer = ask("Install them now? An internet connection is needed. (Enter = yes · n = no)  ")
    if answer is None or answer.lower() not in ("", "y", "yes"):
        print(f"\nTo install them later, run in a terminal:\n   {command}")
        return False
    print()
    done = subprocess.run([sys.executable, "-m", "pip", "install", *missing])
    if done.returncode != 0 or missing_libraries():
        print("\nThe installation did not finish. Check the connection and try in a terminal:")
        print(f"   {command}")
        return False
    print("\nThe libraries are installed.")
    return True


def ensure_table() -> None:
    if TABLE.exists() or not TEMPLATE.exists():
        return
    shutil.copyfile(TEMPLATE, TABLE)
    print(f"\nCreated {TABLE.name} from the template. Add one row per sample to it.")


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


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def load_table() -> dict[str, str]:
    try:
        import openpyxl
        wb = openpyxl.load_workbook(TABLE, read_only=True, data_only=True)
    except Exception:
        return {}
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    wb.close()
    header, cols = None, {}
    for i, row in enumerate(rows[:20]):
        found = {}
        for c, value in enumerate(row):
            name = re.sub(r"\s+", " ", str(value or "")).strip().lower()
            for field, keys in HEADER_KEYS.items():
                if name and field not in found and any(k in name for k in keys):
                    found[field] = c
                    break
        if "id" in found and "area" in found:
            header, cols = i, found
            break
    if header is None:
        return {}

    def get(row, field):
        c = cols.get(field)
        return cell_text(row[c]) if c is not None and c < len(row) else ""

    table = {}
    for row in rows[header + 1:]:
        sample_id = make_id(row[cols["id"]]) if cols["id"] < len(row) else ""
        if not sample_id:
            continue
        bits = []
        if get(row, "layers"):
            bits.append(f"{get(row, 'layers')} layers")
        if get(row, "additive"):
            bits.append(f"additive: {get(row, 'additive')}")
        if get(row, "age"):
            bits.append(get(row, "age"))
        bits.append(f"{get(row, 'area')} cm²" if get(row, "area") else "⚠ no area")
        text = " · ".join(bits)
        table[id_key(sample_id)] = text
        if get(row, "alias"):
            table.setdefault(id_key(get(row, "alias")), text)
    return table


def natural(path: Path) -> list:
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", path.name)]


def by_rate(path: Path) -> tuple:
    m = RATE_RE.search(path.stem)
    try:
        rate = float(m.group(1).replace(",", ".")) if m else float("inf")
    except ValueError:
        rate = float("inf")
    return (rate, natural(path))


def samples_in(folder: Path) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted((d for d in folder.iterdir() if d.is_dir() and not d.name.startswith(".")),
                  key=natural)


def describe(sample_dir: Path, table: dict[str, str]) -> str:
    return table.get(id_key(sample_dir.name),
                     f"⚠ not in {TABLE.name} — add a row for it")


def is_round(path: Path) -> bool:
    return bool(ROUND_RE.search(path.stem))


def round_files(sample_dir: Path) -> list[Path]:
    return sorted((f for f in sample_dir.glob("*.xlsx") if is_round(f)), key=natural)


def total_cycles(files: list[Path]) -> int:
    return sum(int(m.group(1)) if (m := CYCLES_PER_ROUND_RE.search(f.stem)) else 50
               for f in files)


def parse_choice(text: str, items: list[Path], by_sample: dict[str, list[int]]) -> list[int]:
    picked = []
    for tok in re.split(r"[\s,;]+", text):
        if not tok:
            continue
        if re.fullmatch(r"\d+", tok):
            picked.append(int(tok) - 1)
        elif m := re.fullmatch(r"(\d+)-(\d+)", tok):
            a, b = int(m.group(1)), int(m.group(2))
            picked += range(min(a, b) - 1, max(a, b))
        elif id_key(tok) in by_sample:
            picked += by_sample[id_key(tok)]
        else:
            raise ValueError(tok)
    bad = [p + 1 for p in picked if not 0 <= p < len(items)]
    if bad:
        raise ValueError(str(bad[0]))
    return list(dict.fromkeys(picked))


def pick(items: list[Path], by_sample: dict[str, list[int]], hint: str,
         check=None) -> list[Path] | None:
    print(f"\n   {hint}")
    while True:
        text = ask("   Which? (Enter = back)  ")
        if not text:
            return None
        try:
            chosen = [items[i] for i in parse_choice(text, items, by_sample)]
        except ValueError as exc:
            print(f"   Could not understand '{exc}'. Type numbers from the list or sample names.")
            continue
        problem = check(chosen) if check else None
        if problem:
            print(f"   {problem}")
            continue
        if chosen:
            return chosen


def empty_message(folder: Path) -> None:
    rel = folder.relative_to(HERE)
    print(f"   Nothing in '{rel}' yet. Put each sample's files in its own sub-folder,")
    print(f"   e.g. {rel / 'S12'}. See {rel / 'README.md'} for the file names.")


def list_files(title: str, folder: Path, pattern: str, table: dict[str, str]):
    items, by_sample = [], {}
    print(f"\n{title}\n")
    for sd in samples_in(folder):
        files = sorted(sd.glob(pattern), key=by_rate)
        if not files:
            continue
        print(f"   {sd.name:6} {describe(sd, table)}")
        for f in files:
            items.append(f)
            by_sample.setdefault(id_key(sd.name), []).append(len(items) - 1)
            print(f"      [{len(items):2}]  {f.name}")
    if not items:
        empty_message(folder)
    return items, by_sample


def list_samples(title: str, folder: Path, table: dict[str, str], keep, note):
    items, by_sample = [], {}
    print(f"\n{title}\n")
    for sd in samples_in(folder):
        files = keep(sd)
        if not files:
            continue
        items.append(sd)
        by_sample.setdefault(id_key(sd.name), []).append(len(items) - 1)
        print(f"   [{len(items):2}]  {sd.name:6} {describe(sd, table)}   ({note(files)})")
    if not items:
        empty_message(folder)
    return items, by_sample


def show(path) -> str:
    p = Path(path)
    try:
        rel = str(p.relative_to(HERE))
    except ValueError:
        rel = str(p)
    return f'"{rel}"' if " " in rel else rel


def run(script: Path, inputs: list[Path], options: list) -> bool:
    options = [*options, "--samples", TABLE]
    print("\n   Command:")
    print("   python " + show(script))
    for a in inputs:
        print("       " + show(a))
    print("       " + " ".join(show(o) for o in options))
    print("\n" + "-" * 72, flush=True)

    env = dict(os.environ, PYTHONUTF8="1", PYTHONIOENCODING="utf-8", PYTHONUNBUFFERED="1")
    made, locked = [], False
    proc = subprocess.Popen([sys.executable, str(script), *map(str, inputs), *map(str, options)],
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                            encoding="utf-8", errors="replace", env=env)
    for line in proc.stdout:
        print(line, end="", flush=True)
        stripped = line.strip()
        if stripped.startswith("Saved:"):
            made.append(Path(stripped.split(":", 1)[1].strip()))
        if "could not write" in stripped.lower() or "permission denied" in stripped.lower():
            locked = True
    proc.wait()
    print("-" * 72)

    if proc.returncode != 0:
        print("\n   ✗ The analysis did not finish.")
        if locked:
            print("   The result file is open in Excel. Close it and try again.")
        return False
    print(f"\n   ✓ Done · {len(made)} file(s) in {show(RESULTS)}")
    if made:
        open_path(made[0] if len(made) == 1 else made[0].parent)
    return True


def cv_options(scan: str) -> list:
    return (["--scan", scan, "--electrolyte", ELECTROLYTE, "--style", STYLE,
             "--outdir", RESULTS / "CV"] + (["--png"] if PNG else []))


def no_two_rounds(files: list[Path]) -> str | None:
    per_sample: dict[str, int] = {}
    for f in files:
        if is_round(f):
            per_sample[f.parent.name] = per_sample.get(f.parent.name, 0) + 1
    many = [s for s, n in per_sample.items() if n > 1]
    if many:
        return (f"You picked several rounds of {', '.join(many)}: together they are one "
                f"long cycling measurement. Use option [2] for that; here keep at most "
                f"one round per sample.")
    return None


def do_cv(table):
    items, by_sample = list_files("CYCLIC VOLTAMMETRY — compare samples or scan rates",
                                  CV_DIR, "*.xlsx", table)
    if not items:
        return
    files = pick(items, by_sample,
                 "Type numbers (e.g. 1 4 9), a range (2-6) or whole samples (S12 S13).",
                 check=no_two_rounds)
    if not files:
        return
    scan = ask(f"   Which cycle? (Enter = {DEFAULT_SCAN} · 'last' · 'all')  ")
    if scan is None:
        return
    run(CV_PY, files, cv_options(scan or DEFAULT_SCAN))


def do_cycles(table):
    items, by_sample = list_samples(
        "CYCLIC VOLTAMMETRY — stability over many cycles", CV_DIR, table,
        keep=round_files,
        note=lambda fs: f"{len(fs)} rounds · {total_cycles(fs)} cycles")
    if not items:
        return
    dirs = pick(items, by_sample, "Type numbers (e.g. 1 3) or sample names (S12 S13).")
    if not dirs:
        return
    files = [f for d in dirs for f in round_files(d)]
    total = min(total_cycles(round_files(d)) for d in dirs)
    default = ",".join(str(n) for n in (4, total // 4, total // 2, 3 * total // 4, total))
    scan = ask(f"   Which cycles? (Enter = {default} · or e.g. 1,10,50)  ")
    if scan is None:
        return
    scan = scan or default
    if scan.lower() == "last":
        scan = f"{total},"
    elif scan.isdigit():
        scan += ","
    run(CV_PY, files, cv_options(scan))


def do_gcd(table):
    items, by_sample = list_samples(
        "GALVANOSTATIC CHARGE–DISCHARGE — available samples", GCD_DIR, table,
        keep=lambda d: [f for f in d.iterdir() if f.is_file() and f.suffix.lower() == ".xlsx"],
        note=lambda fs: f"{len(fs)} files")
    if not items:
        return
    dirs = pick(items, by_sample, "Type numbers (e.g. 1 3) or sample names (S12 S13).")
    if dirs:
        run(GCD_PY, dirs, ["--style", STYLE, "--outdir", RESULTS / "GCD"]
            + (["--png"] if PNG else []))


def do_lum(table):
    items, by_sample = list_files("TRANSMITTANCE AT 550 nm — available spectra",
                                  LUM_DIR, "*.asc", table)
    if not items:
        return
    files = pick(items, by_sample,
                 "Type numbers (e.g. 1 4 9), a range (2-6) or whole samples (S12 S13).")
    if files:
        run(LUM_PY, files, ["--fresh", "--outdir", RESULTS / "Transmittance"])


def do_thickness(table):
    items, by_sample = list_files("FILM THICKNESS — available profilometer scans",
                                  THK_DIR, "*.dat", table)
    if not items:
        return
    files = pick(items, by_sample,
                 "Type numbers (e.g. 1 4 9), a range (2-6) or whole samples (S12 S13).")
    if not files:
        return
    print("\n   Note: the thickness is read automatically from the flat part of the profile")
    print("   after the edge. This is not always accurate — check each sample's chart.")
    windows: list[str] = []
    while True:
        options = ["--outdir", RESULTS / "Thickness"]
        for w in windows:
            options += ["--window", w]
        if not run(THK_PY, files, options):
            return
        print("\n   If the red window of a sample is in the wrong place, set it by hand:")
        print("   type SAMPLE=X0:X1 with the positions in µm (e.g. S12=1221:1317).")
        text = ask("   Correct a window? (Enter = no)  ")
        if not text:
            return
        if not WINDOW_RE.fullmatch(text):
            print("   That is not of the form SAMPLE=X0:X1.")
            continue
        windows.append(text.replace(" ", ""))


MENU = [
    ("Cyclic voltammetry — compare samples or scan rates", do_cv),
    ("Cyclic voltammetry — stability over many cycles", do_cycles),
    ("Galvanostatic charge–discharge (GCD)", do_gcd),
    ("Transmittance at 550 nm", do_lum),
    ("Film thickness (profilometer)", do_thickness),
]


def main() -> int:
    missing = [p for p in (CV_PY, GCD_PY, LUM_PY, THK_PY) if not p.exists()]
    if missing:
        print(f"The analysis scripts are missing from '{ANALYSIS}'.")
        return 1
    if not python_ok() or not libraries_ok():
        return 1
    ensure_table()
    for folder in (CV_DIR, GCD_DIR, LUM_DIR, THK_DIR):
        folder.mkdir(parents=True, exist_ok=True)

    open_key = str(len(MENU) + 1)
    while True:
        table = load_table()
        print("\n" + "=" * 72)
        print("   THIN-FILM ANALYSIS")
        print("=" * 72)
        for i, (label, _) in enumerate(MENU, start=1):
            print(f"   [{i}]  {label}")
        print(f"   [{open_key}]  Open the results folder")
        print("   [0]  Exit")
        answer = ask("\n   Choose a number:  ")
        if answer is None or answer == "0":
            return 0
        if answer == open_key:
            RESULTS.mkdir(exist_ok=True)
            open_path(RESULTS)
            continue
        if answer.isdigit() and 1 <= int(answer) <= len(MENU):
            MENU[int(answer) - 1][1](table)
        elif answer:
            print("   Type one of the numbers in the list.")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", line_buffering=True)
    except Exception:
        pass
    sys.exit(main())
