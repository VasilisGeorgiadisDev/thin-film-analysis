#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import math
import re
import sys
import tempfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import numpy as np
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RichTextProperties,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

try:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FormatStrFormatter, MultipleLocator
except ImportError:
    print("matplotlib is missing. Install it with:  pip install matplotlib", flush=True)
    raise SystemExit(1)


PALETTE = [
    "#1A5276",
    "#C0392B",
    "#117A65",
    "#784212",
    "#6C3483",
    "#1A6B3C",
    "#B7950B",
    "#2471A3",
    "#922B21",
    "#148F77",
]

NAVY = "1F3864"
BORDER_BLUE = "2F5496"
META_BLUE = "BDD7EE"
TABLE_BLUE = "EBF5FB"
WHITE = "FFFFFF"

SHEET_FONT = "Arial"
THIN = Side(style="thin", color=BORDER_BLUE)
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TITLE_FONT = Font(name=SHEET_FONT, size=12, bold=True, color=NAVY)
LABEL_FONT = Font(name=SHEET_FONT, size=10, bold=True, color=WHITE)
VALUE_FONT = Font(name=SHEET_FONT, size=10)
TABLE_KEY_FONT = Font(name=SHEET_FONT, size=10, bold=True)

NAVY_FILL = PatternFill("solid", fgColor=NAVY)
META_FILL = PatternFill("solid", fgColor=META_BLUE)
TABLE_FILL = PatternFill("solid", fgColor=TABLE_BLUE)

CENTRE = Alignment(horizontal="center", vertical="center")
CENTRE_WRAP = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")

TITLE_ROW_HEIGHT = 18.0
TABLE_HEAD_HEIGHT = 42.0
TABLE_ROW_HEIGHT = 21.95

PLOT_DPI = 180
TEXT_COLOR = "#1F1F1F"
SPINE_COLOR = "#000000"
GRID_COLOR = "#D7D7D7"
ZERO_COLOR = "#BBBBBB"
LEGEND_FACE = "#FAFAFA"
LEGEND_EDGE = "#BBBBBB"

TICK_SIZE = 9.3
AXIS_LABEL_SIZE = 11.0
PLOT_TITLE_SIZE = 11.5
LEGEND_SIZE = 9.3
CURVE_WIDTH = 1.9

LEGEND_GAP_PT = 26.0
TITLE_GAP_PT = 20.0
LEGEND_BORDER_PAD = 0.55
LEGEND_LABEL_SPACING = 0.5

GROUP_AXES = (5.53, 2.99)
TOTAL_AXES = (6.39, 3.24)
MARGIN_LEFT, MARGIN_RIGHT = 0.63, 0.10
MARGIN_BOTTOM = 0.50
MARGIN_TOP_GROUP, MARGIN_TOP_TOTAL = 1.20, 1.28

X_LABEL = "Potential,  E  (V)"
Y_LABEL = "Current Density,  I  (mA cm⁻²)"

PAPER_DPI = 300
PAPER_PALETTE = ["#515151", "#F14040", "#1A6FDF", "#37AD6B", "#B177DE",
                 "#FB6501", "#6E2FA0", "#00A0A0", "#C1272D", "#7F7F7F"]
PAPER_AXES = (8.8, 5.9)
PAPER_MARGIN = (1.20, 0.90, 1.25, 0.60)
PAPER_CURVE_WIDTH = 3.4
PAPER_AXIS_WIDTH = 2.4
PAPER_TICK_SIZE = 19
PAPER_LABEL_SIZE = 23
PAPER_LEGEND_SIZE = 20
PAPER_PANEL_SIZE = 28
PAPER_X_LABEL = "V (mV) vs Ag/AgCl"
PAPER_Y_LABEL = "J (mA/cm$^2$)"
XL_PAPER_X_LABEL = "V (mV) vs Ag/AgCl"
XL_PAPER_Y_LABEL = "J (mA/cm²)"
XL_PAPER_AXIS_EMU = 22000
XL_PAPER_CURVE_EMU = 34000
XL_PAPER_TICK_SIZE = 1400
XL_PAPER_AXIS_TITLE_SIZE = 1600
XL_PAPER_LEGEND_SIZE = 1400

IMAGE_SCALE = 1.15
EXCEL_PX_PER_INCH = 96
EXCEL_ROW_CM = 0.529

GROUP_CHART_CM = (18.5, 12.5)
TOTAL_CHART_CM = (21.5, 13.5)
XL_CURVE_EMU = 22000
XL_BOX_EMU = 15240
XL_ZERO_EMU = 15240
XL_GRID_EMU = 9525
XL_GRID_COLOR = "D9D9D9"
XL_GRID_DASH = "dash"
XL_TITLE_SIZE = 1400
XL_AXIS_TITLE_SIZE = 1200
XL_TICK_SIZE = 1100
XL_LEGEND_SIZE = 1100

GROUP_BLOCK_COLS = 10

META_LABELS = [
    "Sample:",
    "Electrolyte:",
    "Area (cm²):",
    "Layers:",
    "Additive:",
    "Solution age:",
    "Date:",
]
DATA_HEADERS = ["WE(1).Potential (V)", "WE(1).Current (A)", "V (mV)", "I (mA/cm²)"]
DATA_FORMATS = ["0.000000", "0.00000000", "0.0000", "0.000000"]
DATA_COL_WIDTHS = [22.0, 18.0, 14.0, 16.0]

SUMMARY_HEADERS = [
    ("Sample", 10.0),
    ("Properties", 24.0),
    ("Film area\n(cm²)", 14.0),
    ("Loop area\n(A·V)", 15.0),
    ("Mean charge\nQ (mC)", 13.0),
    ("Charge\ndensity\n(mC/cm²)", 13.0),
    ("Areal\ncapacitance\n(mF/cm²)", 16.0),
]

DEFAULT_SAMPLE_TABLE = "samples.xlsx"
ROUND_SHEET = 8
SIG_FIGS = 6


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


def trapezoid(y, x):
    fn = getattr(np, "trapezoid", None) or np.trapz
    return float(fn(y, x))


def sig(value: float, figures: int = SIG_FIGS):
    if value is None or not np.isfinite(value) or value == 0:
        return None if value is None else 0.0
    return round(float(value), -int(math.floor(math.log10(abs(value)))) + (figures - 1))


SUBSCRIPTS = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")


def pretty_formula(text: str) -> str:
    return re.sub(r"(?<=[A-Za-zΑ-Ωα-ω\)])(\d+)",
                  lambda m: m.group(1).translate(SUBSCRIPTS), text or "")


WINDOW_RE = re.compile(r"\[([^\]]*)\]")
RATE_RE = re.compile(r"([\d.,]+)\s*mV[/\s_]*(?:per|p)?[\s_]*s", re.IGNORECASE)
DATE_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")

SAMPLE_TOKEN_RE = re.compile(r"[A-Za-z]{1,3}\d+")
NUMBER_PREFIX = "V"
FORMULA_RE = re.compile(r"[A-Za-z]{1,4}[A-Za-z0-9]*\d$")
COATING_RE = re.compile(r"\d*SC\d*$", re.IGNORECASE)
ROUND_RE = re.compile(r"(?<![A-Za-z0-9])(?:\d+(?:st|nd|rd|th)|\d+c\d*)(?![A-Za-z])",
                      re.IGNORECASE)


def split_sample(tokens: list[str], table: dict | None = None) -> tuple[str, str | None]:
    spots = [i for i, t in enumerate(tokens)
             if SAMPLE_TOKEN_RE.fullmatch(t) and not COATING_RE.fullmatch(t)]
    if not spots:
        return tokens[-1], " ".join(tokens[:-1]) or None
    known = [i for i in spots if table and lookup_sample(tokens[i], table)]
    spot = (known or spots)[-1]
    before = tokens[spot - 1] if spot else ""
    formula = (before if FORMULA_RE.fullmatch(before)
               and not COATING_RE.fullmatch(before) else None)
    return tokens[spot], formula


@dataclass
class FileMeta:
    sample: str
    electrolyte: str | None = None
    rate_mV_s: float | None = None
    v_min: float | None = None
    v_max: float | None = None
    date: str | None = None


def split_window(raw: str) -> list[str]:
    patterns = [r"\s+,\s+", r"\s*;\s*"]
    if "." in raw:
        patterns.append(r"\s*,\s*")
    patterns.append(r"\s+")
    for pattern in patterns:
        parts = [p.strip().rstrip("Vv").strip() for p in re.split(pattern, raw.strip())]
        parts = [p for p in parts if p]
        if len(parts) == 2:
            return parts
    return []


def parse_filename(path: Path, table: dict | None = None) -> FileMeta:
    stem = path.stem.strip()

    window_m = WINDOW_RE.search(stem)
    rate_m = RATE_RE.search(stem)
    date_m = DATE_RE.search(stem)

    v_min = v_max = None
    if window_m:
        raw = window_m.group(1)
        parts = split_window(raw)
        if len(parts) == 2:
            v_min, v_max = to_float(parts[0]), to_float(parts[1])
            if v_min is not None and v_max is not None and v_min > v_max:
                v_min, v_max = v_max, v_min
        else:
            warn(f"Could not understand the potential window '{raw}' in {path.name}.")

    date = None
    if date_m:
        try:
            date = datetime.strptime(date_m.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            warn(f"Invalid date '{date_m.group(1)}' in {path.name}.")

    cut = min(m.start() for m in (window_m, rate_m, date_m) if m) if any(
        (window_m, rate_m, date_m)) else len(stem)
    parts = [p for p in re.split(r"[_\s]+", stem[:cut].strip()) if p]
    if parts and parts[0].upper() == "CV":
        parts = parts[1:]

    if not parts:
        warn(f"No sample name found in '{path.name}'; using the whole name.")
        return FileMeta(sample=stem, rate_mV_s=to_float(rate_m.group(1)) if rate_m else None,
                        v_min=v_min, v_max=v_max, date=date)

    sample, electrolyte = split_sample(parts, table)
    return FileMeta(
        sample=sample,
        electrolyte=electrolyte,
        rate_mV_s=to_float(rate_m.group(1)) if rate_m else None,
        v_min=v_min,
        v_max=v_max,
        date=date,
    )


def _norm(text) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


REQUIRED_COLUMNS = {
    "potential": ["we1potentialv", "wepotentialv", "potentialv"],
    "current":   ["we1currenta", "wecurrenta", "currenta"],
    "scan":      ["scan"],
    "applied":   ["potentialappliedv", "appliedpotentialv"],
    "time":      ["times", "time"],
}


@dataclass
class Cycle:
    number: int
    v_applied: np.ndarray
    t: np.ndarray
    i_a: np.ndarray
    v_meas: np.ndarray

    def __len__(self) -> int:
        return len(self.t)


def _rows_from_xlsx(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    ws.reset_dimensions()
    rows = ws.iter_rows(values_only=True)
    try:
        header = next(rows)
    except StopIteration:
        wb.close()
        raise ValueError(f"{path.name} is empty.")

    def gen():
        try:
            yield from rows
        finally:
            wb.close()

    return header, gen()


def _rows_from_text(path: Path):
    lines = path.read_text(encoding="utf-8-sig", errors="replace").splitlines()
    if not lines:
        raise ValueError(f"{path.name} is empty.")
    header = lines[0].split("\t")

    def gen():
        for line in lines[1:]:
            if not line.strip():
                continue
            yield [to_float(cell) for cell in line.split("\t")]

    return header, gen()


def read_raw(path: Path) -> dict[int, Cycle]:
    reader = _rows_from_text if path.suffix.lower() == ".txt" else _rows_from_xlsx
    header, rows = reader(path)

    index = {_norm(h): c for c, h in enumerate(header) if h is not None}
    col: dict[str, int] = {}
    for key, candidates in REQUIRED_COLUMNS.items():
        for cand in candidates:
            if cand in index:
                col[key] = index[cand]
                break
        else:
            raise ValueError(
                f"No column for '{key}' in {path.name}. "
                f"Found: {[h for h in header if h is not None]}"
            )

    buckets: dict[int, list[tuple[float, float, float, float]]] = {}
    for row in rows:
        if col["scan"] >= len(row):
            continue
        scan = row[col["scan"]]
        if scan is None:
            continue
        values = (row[col["applied"]], row[col["time"]],
                  row[col["current"]], row[col["potential"]])
        if any(v is None for v in values):
            continue
        buckets.setdefault(int(scan), []).append(tuple(float(v) for v in values))

    if not buckets:
        raise ValueError(f"{path.name} contains no numeric data.")

    cycles = {}
    for number in sorted(buckets):
        arr = np.asarray(buckets[number], dtype=float)
        cycles[number] = Cycle(number, arr[:, 0], arr[:, 1], arr[:, 2], arr[:, 3])
    return cycles


@dataclass
class SampleConfig:
    id: str = ""
    area_cm2: float | None = None
    coatings: str = ""
    alias: str = ""
    rpm: str = ""
    additive: str = ""
    age: str = ""
    dimensions: str = ""
    thickness: str = ""
    broken: bool = False

    @property
    def label(self) -> str:
        return f"{self.coatings} layers" if self.coatings else ""


TABLE_FIELDS = {
    "id":        ("sample", "νούμερο", "νουμερο", "number"),
    "alias":     ("alias",),
    "coatings":  ("layers", "spin coatings"),
    "broken":    ("excluded", "χαλασμένο", "χαλασμενο"),
    "rpm":       ("rpm",),
    "additive":  ("additive", "αιθαλολαμίνη", "αιθαλολαμινη"),
    "age":       ("solution age", "ηλικία διαλύματος", "ηλικια διαλυματος"),
    "dims":      ("dimensions", "διαστάσεις", "διαστασεις"),
    "area":      ("area", "εμβαδόν", "εμβαδον"),
    "thickness": ("thickness", "πάχος", "παχος"),
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


LOOKALIKE = str.maketrans("ABEZHIKMNOPTXY", "ΑΒΕΖΗΙΚΜΝΟΡΤΧΥ")


def normalise_additive(value: str) -> str:
    guess = value.upper().translate(LOOKALIKE)
    return guess if guess in ("ΝΑΙ", "ΟΧΙ") else value


def lookup_sample(token: str, table: dict[str, SampleConfig]) -> SampleConfig | None:
    key = id_key(token)
    if not key:
        return None
    if key in table:
        return table[key]
    aliases = {id_key(cfg.alias): cfg for cfg in table.values() if cfg.alias}
    if key in aliases:
        return aliases[key]
    if key.isdigit():
        return table.get(id_key(f"{NUMBER_PREFIX}{key}"))
    return None


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
                area_cm2=to_float(cell("area")),
                coatings=_text(cell("coatings")),
                alias=_text(cell("alias")),
                rpm=_text(cell("rpm")),
                additive=normalise_additive(_text(cell("additive"))),
                age=_text(cell("age")),
                dimensions=_text(cell("dims")),
                thickness=_text(cell("thickness")),
                broken=bool(_text(cell("broken"))),
            )
        return table
    finally:
        wb.close()


def read_sample_csv(path: Path) -> dict[str, SampleConfig]:
    table: dict[str, SampleConfig] = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            sample_id = make_id(row.get("sample"))
            if not sample_id:
                continue
            table[id_key(sample_id)] = SampleConfig(
                id=sample_id,
                area_cm2=to_float(row.get("area_cm2")),
                coatings=_text(row.get("coatings") or row.get("label")),
                alias=_text(row.get("alias")),
            )
    return table


def load_sample_table(path: Path) -> dict[str, SampleConfig]:
    if not path.exists():
        return {}
    if path.suffix.lower() == ".csv":
        return read_sample_csv(path)
    return read_sample_table(path)


@dataclass
class CycleResult:
    number: int
    v_volt: np.ndarray
    i_amp: np.ndarray
    v_mV: np.ndarray
    i_mA_cm2: np.ndarray
    rate_from_data_V_s: float
    delta_v_V: float
    loop: float
    charge_density_mC_cm2: float = float("nan")
    charge_mC: float = float("nan")
    c_areal_mF_cm2: float = float("nan")
    color: str = PALETTE[0]
    sheet: str = ""
    first_row: int = 0
    last_row: int = 0


@dataclass
class SampleResult:
    key: str
    area_cm2: float
    electrolyte: str
    date: str
    rate_mV_s: float
    source: Path
    config: SampleConfig
    cycles: list[CycleResult] = field(default_factory=list)

    @property
    def label(self) -> str:
        return self.config.label

    @property
    def additive(self) -> str:
        return (self.config.additive or "").strip()

    @property
    def age(self) -> str:
        return (self.config.age or "").strip()


def wanted_cycles(spec: str) -> list[int] | None:
    if "," not in spec:
        return None
    return [int(p) for p in spec.split(",") if p.strip()]


def last_complete(cycles: dict[int, Cycle]) -> int:
    spans = {n: float(np.ptp(c.v_meas)) if len(c.v_meas) else 0.0
             for n, c in cycles.items()}
    full = max(spans.values())
    return [n for n in sorted(cycles) if spans[n] >= 0.9 * full][-1]


def select_cycles(cycles: dict[int, Cycle], spec: str, name: str) -> list[Cycle]:
    available = sorted(cycles)
    if spec == "all":
        return [cycles[n] for n in available]
    if spec == "last":
        return [cycles[last_complete(cycles)]]
    picks = wanted_cycles(spec)
    if picks is not None:
        return [cycles[n] for n in available]
    number = int(spec)
    if number not in cycles:
        fallback = last_complete(cycles)
        full = "complete" if fallback != available[-1] else ""
        warn(f"{name} has no Scan {number} (it has {available[0]}–{available[-1]}). "
             f"Using the last{' ' + full if full else ''} one, Scan {fallback}.")
        return [cycles[fallback]]
    return [cycles[number]]


def scan_rate_from_data(cycle: Cycle) -> float | None:
    dv = np.abs(np.diff(cycle.v_applied))
    dt = np.diff(cycle.t)
    valid = (dt > 0) & (dv > 0)
    if not valid.any():
        return None
    return float(np.median(dv[valid] / dt[valid]))


def analyse_cycle(cycle: Cycle, area_cm2: float) -> CycleResult:
    v_mV = cycle.v_meas * 1000.0
    i_dens = cycle.i_a * 1000.0 / area_cm2

    v_closed = np.append(cycle.v_meas, cycle.v_meas[0])
    i_closed = np.append(i_dens, i_dens[0])

    return CycleResult(
        number=cycle.number,
        v_volt=cycle.v_meas,
        i_amp=cycle.i_a,
        v_mV=v_mV,
        i_mA_cm2=i_dens,
        rate_from_data_V_s=scan_rate_from_data(cycle) or 0.0,
        delta_v_V=float(cycle.v_meas.max() - cycle.v_meas.min()),
        loop=abs(trapezoid(i_closed, v_closed)),
    )


def apply_rate(result: CycleResult, rate_mV_s: float, area_cm2: float) -> None:
    nu = rate_mV_s / 1000.0
    if nu <= 0 or result.delta_v_V <= 0 or area_cm2 <= 0:
        return
    result.charge_density_mC_cm2 = result.loop / nu
    result.charge_mC = result.charge_density_mC_cm2 * area_cm2
    result.c_areal_mF_cm2 = result.charge_density_mC_cm2 / result.delta_v_V


def resolve_area(meta: FileMeta, table: dict[str, SampleConfig],
                 area_override: float | None) -> tuple[float, SampleConfig]:
    cfg = lookup_sample(meta.sample, table)

    if cfg is not None and cfg.broken:
        warn(f"Sample {cfg.id} ({cfg.alias or '—'}) is marked as "
             f"excluded in the sample table.")

    if area_override:
        return area_override, cfg or SampleConfig()

    if cfg is None:
        raise ValueError(
            f"'{meta.sample}' is not in the sample table — neither as a Sample "
            f"nor as an Alias. Add it there or give --area."
        )
    if not cfg.area_cm2 or cfg.area_cm2 <= 0:
        raise ValueError(
            f"Sample {cfg.id} ({cfg.alias or '—'}) has no 'Area (cm²)' "
            f"in the sample table. Fill it in or give --area."
        )
    return cfg.area_cm2, cfg


def build_sample(path: Path, table: dict[str, SampleConfig], scan_spec: str,
                 area_override: float | None, electrolyte: str | None = None) -> SampleResult:
    meta = parse_filename(path, table)
    area, cfg = resolve_area(meta, table, area_override)

    cycles = read_raw(path)
    chosen = select_cycles(cycles, scan_spec, path.name)
    results = [analyse_cycle(c, area) for c in chosen]

    measured_mV = results[0].rate_from_data_V_s * 1000.0
    rate_mV = meta.rate_mV_s
    if rate_mV and measured_mV and abs(rate_mV - measured_mV) / rate_mV > 0.05:
        warn(f"{meta.sample}: the file name says {rate_mV:g} mV/s but the data show "
             f"{measured_mV:.1f} mV/s. Using the measured value.")
        rate_mV = measured_mV
    elif not rate_mV:
        rate_mV = measured_mV

    for result in results:
        apply_rate(result, rate_mV or 0.0, area)

    if meta.v_min is not None and meta.v_max is not None:
        window = abs(meta.v_max - meta.v_min)
        drift = abs(window - results[0].delta_v_V)
        if drift > 0.05:
            warn(f"{meta.sample}: the window in the file name ({window:.3f} V) differs by "
                 f"{drift:.3f} V from the measured one ({results[0].delta_v_V:.3f} V).")

    key = cfg.id or meta.sample

    solution = meta.electrolyte or electrolyte or ""
    if not solution:
        warn(f"{key}: the file name does not say the electrolyte. "
             f"Give e.g. --electrolyte AlCl3 to show it in the titles.")

    return SampleResult(
        key=key,
        area_cm2=area,
        electrolyte=solution,
        date=meta.date or "",
        rate_mV_s=rate_mV or 0.0,
        source=path,
        config=cfg,
        cycles=results,
    )


AGE_UNITS = (
    (("week", "εβδομ"), 7),
    (("month", "μήν", "μην"), 30),
    (("day", "ημέρ", "ημερ", "μέρ", "μερ"), 1),
)


def age_rank(text: str) -> float:
    t = (text or "").strip().lower()
    if not t:
        return float("inf")
    if "same" in t or "ίδια" in t or "ιδια" in t:
        return 0.0
    number = to_float(re.sub(r"[^\d.,]", "", t) or None, 1.0) or 1.0
    for keys, days in AGE_UNITS:
        if any(k in t for k in keys):
            return number * days
    return float("inf")


def legend_label(sample: SampleResult, cycle: CycleResult, multi: bool) -> str:
    bits = [sample.key]
    if multi:
        bits[0] = f"{sample.key} · Scan {cycle.number}"
    if sample.label:
        bits.append(sample.label)
    if sample.additive:
        bits.append(f"additive: {sample.additive}")
    if sample.age:
        bits.append(sample.age)
    return " | ".join(bits)


def paper_label(sample: SampleResult, cycle: CycleResult, multi: bool,
                one_sample: bool) -> str:
    if multi and one_sample:
        return f"Cycle {cycle.number}"
    if multi:
        return f"{sample.key} · cycle {cycle.number}"
    bits = [b for b in (sample.label, sample.age) if b]
    return f"{sample.key} ({', '.join(bits)})" if bits else sample.key


def chart_entries(samples: list[SampleResult], multi: bool,
                  style: str = "report") -> list[tuple]:
    if style == "paper":
        one = len(samples) == 1
        return [(cycle, paper_label(sample, cycle, multi, one), cycle.color)
                for sample in samples for cycle in sample.cycles]
    return [(cycle, legend_label(sample, cycle, multi), cycle.color)
            for sample in samples for cycle in sample.cycles]


def traits_label(sample: SampleResult, show_additive: bool) -> str:
    bits = [b for b in (sample.label,
                        f"additive: {sample.additive}" if show_additive and sample.additive else "",
                        sample.age) if b]
    return " | ".join(bits)


@dataclass
class Group:
    coatings: str
    additive: str
    samples: list[SampleResult]

    @property
    def keys(self) -> list[str]:
        return [s.key for s in self.samples]

    @property
    def common(self) -> str:
        additive = f"additive: {self.additive}" if self.additive else ""
        return " | ".join(b for b in (self.coatings, additive) if b)


def group_samples(samples: list[SampleResult]) -> list[Group]:
    order: list[tuple[str, str]] = []
    buckets: dict[tuple[str, str], list[SampleResult]] = {}
    for s in samples:
        key = (s.label, s.additive)
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(s)

    groups = []
    for coatings, additive in order:
        members = sorted(buckets[(coatings, additive)], key=lambda s: age_rank(s.age))
        groups.append(Group(coatings, additive, members))
    return groups


def scan_note(samples: list[SampleResult], multi: bool) -> str:
    if multi:
        return "all scans"
    numbers = {c.number for s in samples for c in s.cycles}
    return f"Scan {numbers.pop()}" if len(numbers) == 1 else "Scan " + \
        "/".join(str(n) for n in sorted(numbers))


def group_titles(group: Group, index: int) -> tuple[str, str]:
    keys = " vs ".join(group.keys)
    ages = " / ".join(dict.fromkeys(s.age for s in group.samples if s.age))
    head = f"Comparison {index}: {keys}"
    common = group.common
    cell = f"{head}  |  {common}" if common else head
    if ages:
        cell = f"{cell} | {ages}" if common else f"{cell}  |  {ages}"
    return cell, (f"{head}  |  {common}" if common else head)


def total_titles(samples: list[SampleResult], multi: bool) -> tuple[str, str]:
    keys = ", ".join(s.key for s in samples)
    head = f"All samples: {keys}"

    electrolytes = [s.electrolyte for s in samples if s.electrolyte]
    parts = []
    if electrolytes:
        parts.append(pretty_formula(electrolytes[0]))
    parts.append(scan_note(samples, multi))
    cell = f"{head}  |  {' | '.join(parts)}"

    coatings = list(dict.fromkeys(s.label for s in samples if s.label))
    additives = list(dict.fromkeys(s.additive for s in samples if s.additive))
    if coatings and additives and not (len(coatings) > 1 and len(additives) > 1):
        cell += f"  |  {' & '.join(coatings)} | {' & '.join(additives)}"
    return cell, head


NICE_STEPS = [0.02, 0.05, 0.1, 0.2, 0.25, 0.5, 1.0, 2.0, 2.5, 5.0, 10.0, 20.0, 25.0, 50.0]
Y_TICKS_PER_SIDE = 5


@dataclass
class Axes:
    x_min: float
    x_max: float
    y_lim: float
    x_step: float
    y_step: float


def plan_axes(samples: list[SampleResult]) -> Axes:
    x_min = min(float(c.v_volt.min()) for s in samples for c in s.cycles)
    x_max = max(float(c.v_volt.max()) for s in samples for c in s.cycles)
    pad = 0.08 * max(x_max - x_min, 1e-6)
    x_min, x_max = x_min - pad, x_max + pad

    peak = max(float(np.abs(c.i_mA_cm2).max()) for s in samples for c in s.cycles)
    needed = peak * 1.25 / Y_TICKS_PER_SIDE
    y_step = next((v for v in NICE_STEPS if v >= needed), NICE_STEPS[-1])

    span = x_max - x_min
    x_step = next((v for v in NICE_STEPS if span / v <= 9), NICE_STEPS[-1])
    return Axes(x_min, x_max, y_step * Y_TICKS_PER_SIDE, x_step, y_step)


def render_chart(entries: list[tuple], title: str,
                 axes: Axes, path: Path, wide: bool) -> tuple[float, float]:
    axes_w, axes_h = TOTAL_AXES if wide else GROUP_AXES
    margin_top = MARGIN_TOP_TOTAL if wide else MARGIN_TOP_GROUP
    fig_w = MARGIN_LEFT + axes_w + MARGIN_RIGHT
    fig_h = MARGIN_BOTTOM + axes_h + margin_top

    fig = plt.figure(figsize=(fig_w, fig_h), dpi=PLOT_DPI, facecolor="white")
    ax = fig.add_axes([MARGIN_LEFT / fig_w, MARGIN_BOTTOM / fig_h,
                       axes_w / fig_w, axes_h / fig_h])

    ax.grid(True, color=GRID_COLOR, linestyle=":", linewidth=0.8, zorder=0)
    ax.axhline(0.0, color=ZERO_COLOR, linestyle="--", linewidth=1.2, zorder=1)
    ax.axvline(0.0, color=ZERO_COLOR, linestyle="--", linewidth=1.2, zorder=1)

    for cycle, label, color in entries:
        ax.plot(cycle.v_volt, cycle.i_mA_cm2, color=color,
                linewidth=CURVE_WIDTH, solid_joinstyle="round", label=label, zorder=3)

    ax.set_xlim(axes.x_min, axes.x_max)
    ax.set_ylim(-axes.y_lim, axes.y_lim)
    ax.xaxis.set_major_locator(MultipleLocator(axes.x_step))
    ax.yaxis.set_major_locator(MultipleLocator(axes.y_step))
    ax.xaxis.set_major_formatter(FormatStrFormatter("%.1f"))
    ax.yaxis.set_major_formatter(FormatStrFormatter("%.1f"))

    ax.set_xlabel(X_LABEL, fontsize=AXIS_LABEL_SIZE, color=TEXT_COLOR, labelpad=9)
    ax.set_ylabel(Y_LABEL, fontsize=AXIS_LABEL_SIZE, color=TEXT_COLOR, labelpad=9)
    ax.tick_params(axis="both", which="major", direction="out", length=4.5, width=1.1,
                   labelsize=TICK_SIZE, colors=TEXT_COLOR)
    for spine in ax.spines.values():
        spine.set_color(SPINE_COLOR)
        spine.set_linewidth(1.2)

    count = len(entries)
    ncol = count if count <= 4 else 3
    rows = math.ceil(count / ncol)
    ax.legend(loc="lower center",
              bbox_to_anchor=(0.5, 1 + LEGEND_GAP_PT / (axes_h * 72)), ncol=ncol,
              fontsize=LEGEND_SIZE, frameon=True, facecolor=LEGEND_FACE,
              edgecolor=LEGEND_EDGE, framealpha=1.0, borderpad=LEGEND_BORDER_PAD,
              columnspacing=1.6, handlelength=2.0, handletextpad=0.6,
              labelspacing=LEGEND_LABEL_SPACING, labelcolor=TEXT_COLOR)
    legend_height = LEGEND_SIZE * (rows * 1.35
                                   + (rows - 1) * LEGEND_LABEL_SPACING
                                   + 2 * LEGEND_BORDER_PAD)
    ax.set_title(title, fontsize=PLOT_TITLE_SIZE, fontweight="bold", color=TEXT_COLOR,
                 pad=LEGEND_GAP_PT + legend_height + TITLE_GAP_PT)

    fig.savefig(path, dpi=PLOT_DPI, bbox_inches="tight", pad_inches=0.05,
                facecolor="white")
    plt.close(fig)
    return png_size_inches(path)


def render_chart_paper(entries: list[tuple], title: str, axes: Axes, path: Path,
                       panel: str = "") -> tuple[float, float]:
    left, right, top, bottom = PAPER_MARGIN
    axes_w, axes_h = PAPER_AXES
    fig_w, fig_h = left + axes_w + right, bottom + axes_h + top
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=PAPER_DPI, facecolor="white")
    ax = fig.add_axes([left / fig_w, bottom / fig_h, axes_w / fig_w, axes_h / fig_h])

    for i, (cycle, label, _) in enumerate(entries):
        ax.plot(cycle.v_mV, cycle.i_mA_cm2, color=PAPER_PALETTE[i % len(PAPER_PALETTE)],
                linewidth=PAPER_CURVE_WIDTH, solid_joinstyle="round",
                solid_capstyle="round", label=label, clip_on=False, zorder=3)

    ax.set_xlim(axes.x_min * 1000.0, axes.x_max * 1000.0)
    ax.set_ylim(-axes.y_lim, axes.y_lim)
    ax.xaxis.set_major_locator(MultipleLocator(axes.x_step * 1000.0))
    ax.yaxis.set_major_locator(MultipleLocator(axes.y_step))

    ax.spines["bottom"].set_position(("data", 0.0))
    ax.spines["right"].set_position(("data", 0.0))
    for side in ("top", "left"):
        ax.spines[side].set_visible(False)
    for side in ("bottom", "right"):
        ax.spines[side].set_linewidth(PAPER_AXIS_WIDTH)
        ax.spines[side].set_color("black")
        ax.spines[side].set_zorder(4)
    ax.yaxis.tick_right()
    ax.tick_params(axis="both", direction="out", length=10, width=PAPER_AXIS_WIDTH,
                   labelsize=PAPER_TICK_SIZE, colors="black", zorder=4)
    ax.tick_params(axis="x", pad=14)
    ax.tick_params(axis="y", pad=8)
    for label in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
        label.set_fontweight("bold")

    ax.annotate(PAPER_X_LABEL, xy=(0.0, 0.0), xycoords="data",
                xytext=(22, -96), textcoords="offset points", ha="left", va="top",
                fontsize=PAPER_LABEL_SIZE, fontweight="bold", annotation_clip=False)
    ax.annotate(PAPER_Y_LABEL, xy=(0.0, 1.0), xycoords=("data", "axes fraction"),
                xytext=(58, 4), textcoords="offset points", ha="left", va="top",
                rotation=90, fontsize=PAPER_LABEL_SIZE, fontweight="bold",
                annotation_clip=False)
    if panel:
        fig.text(0.005, 0.995, f"({panel})", ha="left", va="top",
                 fontsize=PAPER_PANEL_SIZE, fontweight="bold")

    count = len(entries)
    ncol = count if count <= 3 else 3
    legend = ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.01), ncol=ncol,
                       fontsize=PAPER_LEGEND_SIZE, frameon=False,
                       columnspacing=1.1, handlelength=1.6, handletextpad=0.5,
                       labelspacing=0.35, borderpad=0.0)
    for text in legend.get_texts():
        text.set_fontweight("bold")
    for line in legend.get_lines():
        line.set_linewidth(PAPER_CURVE_WIDTH)
    _ = title

    fig.savefig(path, dpi=PAPER_DPI, bbox_inches="tight", pad_inches=0.08,
                facecolor="white")
    plt.close(fig)
    return png_size_inches(path, PAPER_DPI)


def png_size_inches(path: Path, dpi: int = PLOT_DPI) -> tuple[float, float]:
    head = path.read_bytes()[16:24]
    width = int.from_bytes(head[:4], "big")
    height = int.from_bytes(head[4:], "big")
    return width / dpi, height / dpi


def hexcolor(value: str) -> str:
    return value.lstrip("#").upper()


def text_props(size: int, bold: bool = False, color: str = TEXT_COLOR) -> RichText:
    props = CharacterProperties(sz=size, b=bold,
                                solidFill=hexcolor(color))
    return RichText(
        bodyPr=RichTextProperties(),
        p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)],
    )


def restyle_title(title_obj, size: int, bold: bool = True,
                  color: str = TEXT_COLOR) -> None:
    if title_obj is None:
        return
    title_obj.overlay = False
    rich = getattr(getattr(title_obj, "tx", None), "rich", None)
    if rich is None:
        return
    props = CharacterProperties(sz=size, b=bold, solidFill=hexcolor(color))
    for paragraph in rich.p:
        paragraph.pPr = ParagraphProperties(defRPr=props)
        for run in paragraph.r or []:
            run.rPr = props


def style_excel_axis(axis, title: str, lo: float, hi: float, step: float,
                     style: str = "report", number_format: str = "0.0") -> None:
    axis.delete = False
    axis.title = title
    axis.number_format = number_format
    axis.scaling.min, axis.scaling.max = lo, hi
    axis.majorUnit = step
    axis.crosses = "autoZero"

    if style == "paper":
        axis.tickLblPos = "nextTo"
        axis.majorTickMark, axis.minorTickMark = "out", "none"
        axis.spPr = GraphicalProperties(
            ln=LineProperties(solidFill="000000", w=XL_PAPER_AXIS_EMU))
        axis.txPr = text_props(XL_PAPER_TICK_SIZE, bold=True, color="#000000")
        axis.majorGridlines = None
        restyle_title(axis.title, XL_PAPER_AXIS_TITLE_SIZE, bold=True,
                      color="#000000")
        return

    axis.tickLblPos = "low"
    axis.majorTickMark = axis.minorTickMark = "none"
    axis.spPr = GraphicalProperties(
        ln=LineProperties(solidFill=hexcolor(ZERO_COLOR), w=XL_ZERO_EMU, prstDash="dash")
    )
    axis.txPr = text_props(XL_TICK_SIZE)
    axis.majorGridlines = ChartLines(spPr=GraphicalProperties(
        ln=LineProperties(solidFill=XL_GRID_COLOR, w=XL_GRID_EMU, prstDash=XL_GRID_DASH)
    ))
    restyle_title(axis.title, XL_AXIS_TITLE_SIZE, bold=False)


def build_excel_chart(wb: openpyxl.Workbook, entries: list[tuple], title: str,
                      axes: Axes, size_cm: tuple[float, float],
                      style: str = "report") -> ScatterChart:
    paper = style == "paper"
    chart = ScatterChart()
    if not paper:
        chart.title = title
    chart.style = None
    chart.scatterStyle = "lineMarker"
    chart.width, chart.height = size_cm
    chart.visible_cells_only = False

    for i, (cycle, label, color) in enumerate(entries):
        ws = wb[cycle.sheet]
        x_col = 3 if paper else 1
        x = Reference(ws, min_col=x_col, min_row=cycle.first_row, max_row=cycle.last_row)
        y = Reference(ws, min_col=4, min_row=cycle.first_row, max_row=cycle.last_row)
        series = Series(y, x, title=label)
        series.marker = Marker(symbol="none")
        series.smooth = False
        series.graphicalProperties.line.solidFill = hexcolor(
            PAPER_PALETTE[i % len(PAPER_PALETTE)] if paper else color)
        series.graphicalProperties.line.width = (XL_PAPER_CURVE_EMU if paper
                                                 else XL_CURVE_EMU)
        chart.series.append(series)

    if paper:
        step = axes.x_step * 1000.0
        lo = math.floor(axes.x_min * 1000.0 / step) * step
        hi = math.ceil(axes.x_max * 1000.0 / step) * step
        style_excel_axis(chart.x_axis, XL_PAPER_X_LABEL, lo, hi, step, style, "0")
        style_excel_axis(chart.y_axis, XL_PAPER_Y_LABEL, -axes.y_lim, axes.y_lim,
                         axes.y_step, style, "0")
    else:
        style_excel_axis(chart.x_axis, X_LABEL, axes.x_min, axes.x_max, axes.x_step)
        style_excel_axis(chart.y_axis, Y_LABEL, -axes.y_lim, axes.y_lim, axes.y_step)
    chart.x_axis.axPos, chart.y_axis.axPos = "b", "l"
    restyle_title(chart.title, XL_TITLE_SIZE)

    chart.plot_area.spPr = GraphicalProperties(
        noFill=True,
        ln=LineProperties(noFill=True) if paper
        else LineProperties(solidFill=hexcolor(SPINE_COLOR), w=XL_BOX_EMU)
    )
    chart.graphical_properties = GraphicalProperties(
        solidFill="FFFFFF", ln=LineProperties(noFill=True)
    )

    if chart.legend is not None:
        chart.legend.position = "t"
        chart.legend.overlay = False
        if paper:
            chart.legend.txPr = text_props(XL_PAPER_LEGEND_SIZE, bold=True,
                                           color="#000000")
            chart.legend.spPr = GraphicalProperties(
                noFill=True, ln=LineProperties(noFill=True))
        else:
            chart.legend.txPr = text_props(XL_LEGEND_SIZE)
            chart.legend.spPr = GraphicalProperties(
                solidFill=hexcolor(LEGEND_FACE),
                ln=LineProperties(solidFill=hexcolor(LEGEND_EDGE), w=9525),
            )
    if len(chart.series) < 2:
        chart.legend = None
    return chart


def set_width(ws: Worksheet, col: int, width: float) -> None:
    letter = get_column_letter(col)
    current = (ws.column_dimensions[letter].width or 0) if letter in ws.column_dimensions else 0
    ws.column_dimensions[letter].width = max(current, width)


def label_cell(ws: Worksheet, row: int, col: int, value, align=LEFT):
    cell = ws.cell(row=row, column=col, value=value)
    cell.fill, cell.font, cell.alignment, cell.border = NAVY_FILL, LABEL_FONT, align, BOX
    return cell


def value_cell(ws: Worksheet, row: int, col: int, value, fill=META_FILL,
               font=VALUE_FONT, align=CENTRE):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font, cell.alignment, cell.border = font, align, BOX
    if fill is not None:
        cell.fill = fill
    return cell


def write_meta_block(ws: Worksheet, sample: SampleResult, row: int, col: int) -> int:
    values = [
        sample.key,
        sample.electrolyte,
        sample.area_cm2,
        sample.label,
        sample.config.additive,
        sample.config.age,
        sample.date,
    ]
    for offset, (label, value) in enumerate(zip(META_LABELS, values)):
        label_cell(ws, row + offset, col, label)
        value_cell(ws, row + offset, col + 1, value)
    return row + len(META_LABELS) + 1


def write_data_block(ws: Worksheet, sample: SampleResult, header_row: int, col: int,
                     multi: bool, remember: bool = False) -> None:
    headers = DATA_HEADERS + (["Scan"] if multi else [])
    for offset, header in enumerate(headers):
        label_cell(ws, header_row, col + offset, header, align=CENTRE_WRAP)
    for offset, width in enumerate(DATA_COL_WIDTHS):
        set_width(ws, col + offset, width)
    if multi:
        set_width(ws, col + len(DATA_HEADERS), 8.0)

    row = header_row + 1
    for cycle in sample.cycles:
        if remember:
            cycle.sheet, cycle.first_row = ws.title, row
            cycle.last_row = row + len(cycle.v_volt) - 1
        for v_volt, i_amp, v_mV, i_dens in zip(
                cycle.v_volt, cycle.i_amp, cycle.v_mV, cycle.i_mA_cm2):
            for offset, (value, fmt) in enumerate(
                    zip((v_volt, i_amp, v_mV, i_dens), DATA_FORMATS)):
                cell = value_cell(ws, row, col + offset,
                                  round(float(value), ROUND_SHEET), fill=None)
                cell.number_format = fmt
            if multi:
                value_cell(ws, row, col + len(DATA_HEADERS), cycle.number, fill=None)
            row += 1


def write_summary_table(ws: Worksheet, samples: list[SampleResult], row: int, col: int,
                        show_additive: bool, multi: bool) -> int:
    for offset, (header, width) in enumerate(SUMMARY_HEADERS):
        label_cell(ws, row, col + offset, header, align=CENTRE_WRAP)
        set_width(ws, col + offset, width)
    ws.row_dimensions[row].height = TABLE_HEAD_HEIGHT

    widest = 0
    line = row + 1
    for sample in samples:
        traits = traits_label(sample, show_additive)
        widest = max(widest, len(traits))
        for cycle in sample.cycles:
            name = f"{sample.key} – Scan {cycle.number}" if multi else sample.key
            values = [
                name,
                traits,
                sig(sample.area_cm2),
                sig(cycle.loop),
                sig(cycle.charge_mC),
                sig(cycle.charge_density_mC_cm2),
                sig(cycle.c_areal_mF_cm2),
            ]
            for offset, value in enumerate(values):
                value_cell(ws, line, col + offset, value, fill=TABLE_FILL,
                           font=TABLE_KEY_FONT if offset == 0 else VALUE_FONT,
                           align=CENTRE_WRAP)
            ws.row_dimensions[line].height = TABLE_ROW_HEIGHT
            line += 1

    set_width(ws, col + 1, widest + 4)
    return line


def rows_for(height_cm: float) -> int:
    return math.ceil(height_cm / EXCEL_ROW_CM) + 2


def place_image(ws: Worksheet, path: Path, row: int, col: int,
                size_in: tuple[float, float]) -> float:
    image = XLImage(str(path))
    image.width = size_in[0] * EXCEL_PX_PER_INCH * IMAGE_SCALE
    image.height = size_in[1] * EXCEL_PX_PER_INCH * IMAGE_SCALE
    ws.add_image(image, f"{get_column_letter(col)}{row}")
    return size_in[1] * IMAGE_SCALE * 2.54


def place_chart(wb: openpyxl.Workbook, ws: Worksheet, entries: list[tuple], title: str,
                axes: Axes, row: int, col: int, wide: bool, tmp: Path, mode: str,
                png_name: str, png_dir: Path | None, style: str = "report",
                panel: str = "") -> float:
    height_cm = 0.0
    if mode in ("image", "both") or png_dir is not None:
        png = tmp / f"{png_name}.png"
        size = (render_chart_paper(entries, title, axes, png, panel=panel)
                if style == "paper"
                else render_chart(entries, title, axes, png, wide=wide))
        if png_dir is not None:
            png_dir.mkdir(parents=True, exist_ok=True)
            (png_dir / png.name).write_bytes(png.read_bytes())
        if mode in ("image", "both"):
            height_cm = place_image(ws, png, row, col, size)
    if mode in ("excel", "both"):
        size_cm = TOTAL_CHART_CM if wide else GROUP_CHART_CM
        chart = build_excel_chart(wb, entries, title, axes, size_cm, style)
        offset = GROUP_BLOCK_COLS if mode == "both" else 0
        ws.add_chart(chart, f"{get_column_letter(col + offset)}{row}")
        height_cm = max(height_cm, size_cm[1])
    return height_cm


def build_chart_sheet(wb: openpyxl.Workbook, samples: list[SampleResult],
                      groups: list[Group], multi: bool, axes: Axes,
                      tmp: Path, override_title: str | None, mode: str,
                      png_dir: Path | None, style: str = "report",
                      panel: str = "") -> None:
    ws = wb.create_sheet(title="CV Chart")
    show_additive = len({s.additive for s in samples}) > 1
    ws.row_dimensions[1].height = TITLE_ROW_HEIGHT
    stride = GROUP_BLOCK_COLS * (2 if mode == "both" else 1)

    table_row = 3
    if len(groups) > 1:
        for index, group in enumerate(groups, start=1):
            col = 1 + (index - 1) * stride
            cell_title, chart_title = group_titles(group, index)
            ws.cell(row=1, column=col, value=cell_title).font = TITLE_FONT
            height = place_chart(wb, ws, chart_entries(group.samples, multi, style),
                                 chart_title, axes, 2, col, False, tmp, mode,
                                 f"Comparison {index}", png_dir, style)
            table_row = max(table_row, 2 + rows_for(height))

        end = 0
        for index, group in enumerate(groups, start=1):
            col = 1 + (index - 1) * stride
            end = max(end, write_summary_table(ws, group.samples, table_row, col,
                                               show_additive, multi))
        table_row = end + 2

    cell_title, chart_title = total_titles(samples, multi)
    ws.cell(row=table_row, column=1, value=override_title or cell_title).font = TITLE_FONT
    ws.row_dimensions[table_row].height = TITLE_ROW_HEIGHT

    height = place_chart(wb, ws, chart_entries(samples, multi, style),
                         override_title or chart_title, axes,
                         table_row + 1, 1, True, tmp, mode,
                         "All samples", png_dir, style, panel)
    write_summary_table(ws, samples, table_row + 1 + rows_for(height), 1,
                        show_additive, multi)


def build_combined_sheet(wb: openpyxl.Workbook, samples: list[SampleResult],
                         multi: bool) -> None:
    ws = wb.create_sheet(title="All Samples Combined")
    data_cols = len(DATA_HEADERS) + (1 if multi else 0)
    block = data_cols + 1

    for i, sample in enumerate(samples):
        col = 1 + i * block
        traits = traits_label(sample, show_additive=True) or sample.label
        header = label_cell(ws, 1, col, f"Sample: {sample.key}  |  {traits}",
                            align=CENTRE)
        for offset in range(1, data_cols):
            ws.cell(row=1, column=col + offset).border = BOX
        ws.merge_cells(start_row=1, start_column=col,
                       end_row=1, end_column=col + data_cols - 1)
        header.fill = NAVY_FILL

        header_row = write_meta_block(ws, sample, row=2, col=col)
        write_data_block(ws, sample, header_row, col, multi)

    ws.freeze_panes = ws.cell(row=len(META_LABELS) + 4, column=1)


def build_sample_sheets(wb: openpyxl.Workbook, samples: list[SampleResult],
                        multi: bool) -> None:
    for sample in samples:
        ws = wb.create_sheet(title=unique_sheet_title(wb, sample.key))
        header_row = write_meta_block(ws, sample, row=1, col=1)
        write_data_block(ws, sample, header_row, col=1, multi=multi, remember=True)
        ws.freeze_panes = ws.cell(row=header_row + 1, column=1)


def unique_sheet_title(wb: openpyxl.Workbook, wanted: str) -> str:
    title = wanted[:31] or "Sample"
    if title not in wb.sheetnames:
        return title
    for n in range(2, 100):
        suffix = f" ({n})"
        candidate = title[:31 - len(suffix)] + suffix
        if candidate not in wb.sheetnames:
            return candidate
    raise ValueError(f"Could not find a unique sheet name for '{wanted}'.")


def build_workbook(samples: list[SampleResult], groups: list[Group], multi: bool,
                   axes: Axes, tmp: Path, title: str | None, mode: str,
                   png_dir: Path | None, style: str = "report",
                   panel: str = "") -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_sample_sheets(wb, samples, multi)
    build_combined_sheet(wb, samples, multi)
    build_chart_sheet(wb, samples, groups, multi, axes, tmp, title, mode, png_dir,
                      style, panel)
    for name in ("All Samples Combined", "CV Chart"):
        ws = wb[name]
        wb.move_sheet(ws, offset=-wb.index(ws))
    return wb


def collect_inputs(paths: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in paths:
        p = Path(raw).expanduser()
        if p.is_dir():
            files.extend(sorted(q for q in p.iterdir()
                                if q.suffix.lower() in (".xlsx", ".txt")
                                and not q.name.startswith("~$")))
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


DEFAULT_OUTPUT_DIR = "results"
SEARCH_LEVELS = 3


def resolve_output_dir(given: str | None, script_dir: Path, inputs: list[Path]) -> Path:
    if given:
        return Path(given).expanduser()
    starts = [p.parent for p in inputs] + [Path.cwd(), script_dir]
    for start in starts:
        for folder in [start, *list(start.parents)[:SEARCH_LEVELS]]:
            candidate = folder / DEFAULT_OUTPUT_DIR
            if candidate.is_dir():
                return candidate
    return script_dir.parent / DEFAULT_OUTPUT_DIR


def find_sample_table(name: str, script_dir: Path, inputs: list[Path]) -> Path:
    given = Path(name).expanduser()
    if given.is_absolute():
        return given

    starts = [Path.cwd(), script_dir]
    starts += [p.parent for p in inputs]
    seen, candidates = set(), []
    for start in starts:
        for folder in [start, *start.parents]:
            candidate = folder / given
            if candidate not in seen:
                seen.add(candidate)
                candidates.append(candidate)

    for candidate in candidates:
        if candidate.exists():
            return candidate
    return script_dir / given


TAG_SEP = " · "


def default_output_name(samples: list[SampleResult]) -> str:
    keys = "_".join(dict.fromkeys(s.key.split(TAG_SEP)[0] for s in samples))
    if any(TAG_SEP in s.key for s in samples):
        rates = sorted({s.rate_mV_s for s in samples if s.rate_mV_s})
        suffix = (f"_{rates[0]:g}-{rates[-1]:g}mVpers" if len(rates) > 1
                  else f"_{rates[0]:g}mVpers" if rates else "")
    else:
        rate = samples[0].rate_mV_s if samples else 0
        suffix = f"_{rate:g}mVpers" if rate else ""
    return f"CV_Analysis_{keys}{suffix}.xlsx"


def scan_spec(value: str) -> str:
    v = value.strip().lower()
    if v in ("last", "all"):
        return v
    try:
        if "," in v:
            if all(int(p) >= 1 for p in v.split(",") if p.strip()):
                return v
        elif int(v) >= 1:
            return v
    except ValueError:
        pass
    raise argparse.ArgumentTypeError(
        "--scan accepts: last, all, a positive integer (e.g. 4), or a list (e.g. "
        "4,25,50,75,100).")


def join_rounds(samples: list[SampleResult]) -> list[SampleResult]:
    by_key: dict[tuple, list[SampleResult]] = {}
    for s in samples:
        is_round = bool(ROUND_RE.search(s.source.stem))
        group_id = (s.key, s.rate_mV_s) if is_round else (s.key, s.rate_mV_s, id(s))
        by_key.setdefault(group_id, []).append(s)

    joined, rounds = [], {}
    for (key, *_), group in by_key.items():
        first = group[0]
        rounds[id(first)] = len(group) if ROUND_RE.search(first.source.stem) else 0
        if len(group) > 1:
            offset = 0
            for part in group:
                for cycle in part.cycles:
                    cycle.number += offset
                offset += len(part.cycles)
            first.cycles = [c for part in group for c in part.cycles]
            log(f"  · {key}: {len(group)} rounds joined into "
                f"{len(first.cycles)} consecutive cycles "
                f"({', '.join(p.source.name for p in group)})")
        joined.append(first)

    by_sample: dict[str, list[SampleResult]] = {}
    for s in joined:
        by_sample.setdefault(s.key, []).append(s)
    for key, same in by_sample.items():
        if len(same) < 2:
            continue
        many_rates = len({s.rate_mV_s for s in same}) > 1
        many_dates = len({s.date for s in same}) > 1
        mixed = len({rounds[id(s)] > 0 for s in same}) > 1
        tags = []
        for s in same:
            bits = []
            if many_rates:
                bits.append(f"{s.rate_mV_s:g} mV s⁻¹")
            if mixed:
                bits.append(f"{rounds[id(s)]} rounds" if rounds[id(s)] else "single")
            if many_dates and s.date:
                bits.append(s.date)
            tags.append(TAG_SEP.join(bits))
        for i, (s, tag) in enumerate(zip(same, tags), start=1):
            if not tag or tags.count(tag) > 1:
                tag = TAG_SEP.join(b for b in (tag, f"#{i}") if b)
            s.key = f"{key}{TAG_SEP}{tag}"
    return joined


def keep_cycles(samples: list[SampleResult], picks: list[int]) -> None:
    for sample in samples:
        available = {c.number: c for c in sample.cycles}
        missing = [n for n in picks if n not in available]
        if missing:
            warn(f"{sample.key}: cycles {missing} do not exist "
                 f"(it has 1–{max(available)}).")
        sample.cycles = [available[n] for n in picks if n in available]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Cyclic voltammetry (CV) analysis, with ready-made Excel charts.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python cv_analysis.py "data/CV/S12/CV_AlCl3_S12 [-1,2 , 0,2] 50mVpers.xlsx"
  python cv_analysis.py data/CV/S12/*.xlsx --scan 4
  python cv_analysis.py file1.xlsx file2.xlsx -o Comparison.xlsx --scan 4
""")
    parser.add_argument("inputs", nargs="+", help="Raw .xlsx files (or a folder).")
    parser.add_argument("-o", "--out", help="Name of the output file.")
    parser.add_argument("--outdir",
                        help=f"Output folder (default: '{DEFAULT_OUTPUT_DIR}').")
    parser.add_argument("--scan", type=scan_spec, default="last",
                        help="Which cycle: last (default), all, a number (e.g. 4) or a list (e.g. 4,25,50).")
    parser.add_argument("--samples", default=DEFAULT_SAMPLE_TABLE,
                        help=f"Sample table with 'Sample' and 'Area (cm²)' columns "
                             f"(default: {DEFAULT_SAMPLE_TABLE}).")
    parser.add_argument("--area", type=float,
                        help="Same area (cm²) for all files; overrides the sample table.")
    parser.add_argument("--electrolyte",
                        help="Electrolyte, when the file name does not say (e.g. AlCl3).")
    parser.add_argument("--title", help="Title of the combined chart.")
    parser.add_argument("--no-groups", action="store_true",
                        help="Only the combined chart, without the per-group charts.")
    parser.add_argument("--charts", choices=("image", "excel", "both"), default="excel",
                        help="excel: native Excel charts (default). "
                             "image: matplotlib images. "
                             "both: both, side by side.")
    parser.add_argument("--style", choices=("report", "paper"), default="report",
                        help="report: worksheet style (default). "
                             "paper: publication style — axes crossing "
                             "at zero, no frame, mV and J.")
    parser.add_argument("--panel", default="",
                        help="Panel letter for the paper style, e.g. b.")
    parser.add_argument("--png", action="store_true",
                        help="Also save the charts as separate 180 dpi PNG images, "
                             "in a folder next to the output file.")
    args = parser.parse_args(argv)

    files = collect_inputs(args.inputs)
    if not files:
        log("No valid input file was given.")
        return 1

    script_dir = Path(__file__).resolve().parent
    samples_path = find_sample_table(args.samples, script_dir, files)
    try:
        table = load_sample_table(samples_path)
    except Exception as exc:
        log(f"Could not read the sample table {samples_path.name}: {exc}")
        return 1

    if table:
        log(f"Sample table: {samples_path.name}  ({len(table)} samples)")
    elif not args.area:
        log(f"Sample table not found ({samples_path}). "
            f"Give --samples <file> or --area <cm²>.")
        return 1

    log(f"Files to process: {len(files)}   |   Cycle: {args.scan}")
    results: list[SampleResult] = []
    for path in files:
        log(f"  · {path.name}")
        try:
            results.append(build_sample(path, table, args.scan, args.area,
                                        args.electrolyte))
        except Exception as exc:
            warn(f"Skipping {path.name}: {exc}")

    if not results:
        log("No file was processed successfully.")
        return 1

    results = join_rounds(results)
    picks = wanted_cycles(args.scan)
    if picks:
        keep_cycles(results, picks)
        results = [s for s in results if s.cycles]
        if not results:
            log("No sample has the cycles you asked for.")
            return 1

    multi = args.scan == "all" or bool(picks)
    groups = [Group("", "", results)] if args.no_groups else group_samples(results)
    ordered = [s for g in groups for s in g.samples]
    palette = PAPER_PALETTE if args.style == "paper" else PALETTE
    slot = 0
    for sample in ordered:
        for cycle in sample.cycles:
            cycle.color = palette[slot % len(palette)]
            if multi:
                slot += 1
        if not multi:
            slot += 1
    axes = plan_axes(ordered)

    if args.out:
        out_path = Path(args.out)
        if not out_path.is_absolute():
            out_path = Path.cwd() / out_path
    else:
        out_dir = resolve_output_dir(args.outdir, script_dir, files)
        out_dir.mkdir(parents=True, exist_ok=True)
        out_path = out_dir / default_output_name(ordered)
    if out_path.suffix.lower() != ".xlsx":
        out_path = out_path.with_suffix(".xlsx")

    png_dir = out_path.with_suffix("") if args.png else None
    with tempfile.TemporaryDirectory(prefix="cv_charts_") as tmpdir:
        wb = build_workbook(ordered, groups, multi, axes, Path(tmpdir), args.title,
                            args.charts, png_dir, args.style, args.panel)
        try:
            wb.save(out_path)
        except PermissionError:
            log(f"\nCould not write {out_path.name} — is it open in Excel? "
                f"Close it and run again.")
            return 1

    log("")
    log(f"Saved: {out_path}")
    if png_dir is not None:
        log(f"Images: {png_dir}")
    for index, group in enumerate(groups, start=1):
        if len(groups) > 1:
            log(f"   Comparison {index}: {', '.join(group.keys)}"
                f"{'  ·  ' + group.common if group.common else ''}")
        for sample in group.samples:
            for cycle in sample.cycles:
                cap = ("—" if not np.isfinite(cycle.c_areal_mF_cm2)
                       else f"{cycle.c_areal_mF_cm2:.4f} mF/cm²")
                log(f"      {sample.key}  Scan {cycle.number}  ·  "
                    f"{len(cycle.v_mV)} points  ·  C = {cap}")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
