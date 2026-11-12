#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import numpy as np
import openpyxl
from openpyxl.chart import BarChart, Reference, ScatterChart, Series
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import (
    CharacterProperties,
    Paragraph,
    ParagraphProperties,
    RichTextProperties,
)
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet


WINDOW_POINTS = 21
PEAK_FRACTION = 0.5
SLOPE_TOL_NM_PER_UM = 0.5
ROUGH_TOL_UM = 0.006
RELAX_STEPS = 4
BASELINE_SPAN = 3
ENVELOPE_SPAN = 2

TILT_LEVEL_THRESHOLD = 0.3


CENTRE = Alignment(horizontal="center", vertical="center")
TITLE_FONT = Font(bold=True, size=12)
META_FILL = PatternFill("solid", fgColor="1F4E79")
META_FONT = Font(bold=True, color="FFFFFF", size=11)
VALUE_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill("solid", fgColor="D6E4F0")
HEADER_FONT = Font(bold=True, size=10)
DATA_FONT = Font(size=10)

PROFILE_COLOR = "9DC3E6"
WINDOW_COLOR = "C00000"
LEVEL_COLOR = "404040"
PALETTE = ["1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD",
           "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF"]

PROFILE_WIDTH = 12700
WINDOW_WIDTH = 38100
LEVEL_WIDTH = 12700
CHART_W_CM, CHART_H_CM = 26, 15
CHART_ROW_SPAN = int(CHART_H_CM / 0.53) + 2

AXIS_COLOR = "000000"
AXIS_WIDTH_EMU = 28575
GRID_COLOR = "E6E6E6"
GRID_WIDTH_EMU = 6350
TICK_LABEL_SIZE = 1200
AXIS_TITLE_SIZE = 1400
CHART_TITLE_SIZE = 1600

DEFAULT_SAMPLE_TABLE = "samples.xlsx"
DEFAULT_OUTPUT_DIR = "results"
DEFAULT_OUTPUT = "Film_Thickness_Analysis.xlsx"
SEARCH_LEVELS = 4
MISSING = "—"


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
    rpm: str = ""


TABLE_FIELDS = {
    "id":       ("sample", "νούμερο", "νουμερο", "number"),
    "alias":    ("alias",),
    "coatings": ("layers", "spin coatings"),
    "additive": ("additive", "αιθαλολαμίνη", "αιθαλολαμινη"),
    "age":      ("solution age", "ηλικία διαλύματος", "ηλικια διαλυματος"),
    "rpm":      ("rpm",),
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
                id=sample_id, coatings=_text(cell("coatings")),
                alias=_text(cell("alias")), additive=_text(cell("additive")),
                age=_text(cell("age")), rpm=_text(cell("rpm")))
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
RPM_RE = re.compile(r"(?<![A-Za-z0-9])(\d+)\s*(K?)\s*RPM(?![A-Za-z0-9])", re.IGNORECASE)
DATE_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")

Z_TO_UM = {"ANGSTROM": 1e-4, "NM": 1e-3, "NANOMETER": 1e-3,
           "MICRON": 1.0, "MICROMETER": 1.0, "UM": 1.0}
X_TO_UM = {"MICRON": 1.0, "MICROMETER": 1.0, "UM": 1.0,
           "MM": 1e3, "MILLIMETER": 1e3}


@dataclass
class Profile:
    sample: str
    coatings: str
    rpm: str
    date: str
    x_um: np.ndarray
    z_um: np.ndarray
    source: Path


def read_profile(path: Path, table: dict[str, SampleConfig] | None = None) -> Profile:
    lines = path.read_text(encoding="latin-1").splitlines()

    header = {}
    data: list[tuple[float, float]] = []
    for line in lines:
        parts = [p.strip().strip('"') for p in line.split(",")]
        if len(parts) == 2 and parts[0].endswith(":"):
            header[parts[0].rstrip(":").upper()] = parts[1]
            continue
        if len(parts) == 2:
            a, b = to_float(parts[0]), to_float(parts[1])
            if a is not None and b is not None:
                data.append((a, b))
    if len(data) < WINDOW_POINTS * 2:
        raise ValueError(f"too few points ({len(data)})")

    x_unit = header.get("X UNITS", "MICRON").upper()
    z_unit = header.get("Z UNITS", "ANGSTROM").upper()
    if x_unit not in X_TO_UM:
        warn(f"{path.name}: unknown X unit '{x_unit}', assuming micrometres.")
    if z_unit not in Z_TO_UM:
        warn(f"{path.name}: unknown Z unit '{z_unit}', assuming ångström.")

    arr = np.asarray(data, dtype=float)
    x = arr[:, 0] * X_TO_UM.get(x_unit, 1.0)
    z = arr[:, 1] * Z_TO_UM.get(z_unit, 1e-4)

    stem = path.name
    sample = sample_in_name(stem, table)
    if not sample:
        raise ValueError("no sample name (e.g. S12) found in the file name")

    cm = COATINGS_RE.search(stem)
    coatings = next((g for g in (cm.groups() if cm else ()) if g), "")
    rm = RPM_RE.search(stem)
    rpm = f"{rm.group(1)}{'000' if rm.group(2) else ''}" if rm else ""

    date = ""
    dm = DATE_RE.search(stem)
    if dm:
        try:
            date = datetime.strptime(dm.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            date = ""
    if not date and "DATE" in header:
        for fmt in ("%m-%d-%Y", "%d-%m-%Y", "%Y-%m-%d"):
            try:
                date = datetime.strptime(header["DATE"], fmt).strftime("%Y-%m-%d")
                break
            except ValueError:
                continue

    return Profile(sample=sample, coatings=coatings, rpm=rpm, date=date,
                   x_um=x, z_um=z, source=path)


@dataclass
class Analysis:
    profile: Profile
    config: SampleConfig
    z_um: np.ndarray
    levelled: bool
    peak_index: int
    peak_x: float
    start: int
    stop: int
    film_mean: float
    baseline_mean: float
    roughness: float
    slope_nm_per_um: float
    tilt_nm_per_um: float
    scan_length_um: float
    relaxed: int
    manual: bool
    baseline_x: float = 0.0

    @property
    def name(self) -> str:
        return self.config.id or self.profile.sample

    @property
    def tilt_total_nm(self) -> float:
        return self.tilt_nm_per_um * self.scan_length_um

    @property
    def tilt_error_nm(self) -> float:
        x = self.profile.x_um
        reach = float(x[(self.start + self.stop) // 2]) - self.baseline_x
        return self.tilt_nm_per_um * reach

    @property
    def thickness_um(self) -> float:
        return self.film_mean - self.baseline_mean

    @property
    def raw_um(self) -> float:
        return self.film_mean


def window_metrics(x: np.ndarray, z: np.ndarray, n: int):
    count = len(x) - n + 1
    rough = np.empty(count)
    slope = np.empty(count)
    mean = np.empty(count)
    for i in range(count):
        wx, wz = x[i:i + n], z[i:i + n]
        a, b = np.polyfit(wx, wz, 1)
        rough[i] = np.std(wz - (a * wx + b))
        slope[i] = a
        mean[i] = wz.mean()
    return rough, slope, mean


def find_baseline_window(z: np.ndarray, edge_start: int, n: int) -> int:
    return max(0, edge_start - BASELINE_SPAN * n)


def find_baseline(x: np.ndarray, z: np.ndarray, edge_start: int, n: int) -> float:
    lo = find_baseline_window(z, edge_start, n)
    hi = max(lo + 1, edge_start)
    return float(np.median(z[lo:hi]))


def edge_bounds(x: np.ndarray, z: np.ndarray) -> tuple[int, int]:
    if len(x) < 3:
        return 0, 0
    d = np.abs(np.diff(z) / np.diff(x))
    threshold = max(float(np.median(d)) * 20.0, float(d.max()) * 0.15)
    big = np.flatnonzero(d > threshold)
    if not len(big):
        return 0, 0
    return int(big[0]), int(big[-1]) + 1


def find_edge(x: np.ndarray, z: np.ndarray) -> int:
    return edge_bounds(x, z)[1]


def substrate_tilt(x: np.ndarray, z: np.ndarray, edge: int, n: int) -> float:
    cut = max(edge - n, n)
    if cut < n or cut > len(x):
        return 0.0
    a, _ = np.polyfit(x[:cut], z[:cut], 1)
    return float(a * 1000.0)


def find_edge_start(x: np.ndarray, z: np.ndarray) -> int:
    return edge_bounds(x, z)[0]


SUBSTRATE_SPAN_UM = 400.0
SUBSTRATE_GAP_UM = 30.0


def clean_substrate_tilt(x: np.ndarray, z: np.ndarray, n: int) -> float:
    start = find_edge_start(x, z)
    if start <= 0:
        return 0.0
    right = x[start] - SUBSTRATE_GAP_UM
    zone = (x >= right - SUBSTRATE_SPAN_UM) & (x <= right)
    if zone.sum() < 2 * n:
        zone = x <= right
    if zone.sum() < n:
        return 0.0
    a, _ = np.polyfit(x[zone], z[zone], 1)
    return float(a * 1000.0)


def find_plateau(x: np.ndarray, z: np.ndarray, n: int) -> tuple[int, int, int, int]:
    peak = find_edge(x, z)
    rough, slope, mean = window_metrics(x, z, n)
    first = min(peak + 1, len(rough) - 1)

    for relaxed in range(RELAX_STEPS + 1):
        factor = 1.5 ** relaxed
        slope_tol = SLOPE_TOL_NM_PER_UM * factor / 1000.0
        rough_tol = ROUGH_TOL_UM * factor
        ok = [i for i in range(first, len(rough))
              if abs(slope[i]) <= slope_tol and rough[i] <= rough_tol]
        if not ok:
            continue
        start = ok[0]
        end = start
        for i in ok[1:]:
            if i != end + 1:
                break
            end = i
        limit = min(end, start + ENVELOPE_SPAN * n)
        best = int(start + np.argmin(mean[start:limit + 1]))
        return peak, best, best + n - 1, relaxed

    fallback = int(first + np.argmin(rough[first:]))
    return peak, fallback, fallback + n - 1, -1


def level_profile(x: np.ndarray, z: np.ndarray, edge: int, n: int, enabled: bool
                  ) -> tuple[np.ndarray, float, bool]:
    tilt = substrate_tilt(x, z, edge, n)
    if not enabled or abs(tilt) <= TILT_LEVEL_THRESHOLD:
        return z, tilt, False
    cut = max(edge - n, n)
    a, b = np.polyfit(x[:cut], z[:cut], 1)
    return z - (a * x + b), tilt, True


def analyse(profile: Profile, config: SampleConfig,
            manual: tuple[float, float] | None, level: bool = False) -> Analysis:
    x, raw, n = profile.x_um, profile.z_um, WINDOW_POINTS
    z, tilt, levelled = level_profile(x, raw, find_edge(x, raw), n, level)
    peak, start, stop, relaxed = find_plateau(x, z, n)

    if manual is not None:
        lo, hi = manual
        inside = np.flatnonzero((x >= lo) & (x <= hi))
        if len(inside) < 3:
            raise ValueError(f"the window {lo:g}–{hi:g} µm contains {len(inside)} points")
        start, stop = int(inside[0]), int(inside[-1])
        relaxed = 0

    window_x, window_z = x[start:stop + 1], z[start:stop + 1]
    a, b = np.polyfit(window_x, window_z, 1)
    edge_start = find_edge_start(x, z)
    base = find_baseline_window(z, edge_start, n)
    if not levelled:
        tilt = clean_substrate_tilt(x, z, n)
    return Analysis(
        profile=profile, config=config, z_um=z, levelled=levelled,
        peak_index=peak, peak_x=float(x[peak]),
        start=start, stop=stop,
        film_mean=float(window_z.mean()),
        baseline_mean=find_baseline(x, z, edge_start, n),
        roughness=float(np.std(window_z - (a * window_x + b))),
        slope_nm_per_um=float(a * 1000.0),
        tilt_nm_per_um=tilt,
        scan_length_um=float(x[-1] - x[0]),
        relaxed=relaxed, manual=manual is not None,
        baseline_x=float(x[min(base + (edge_start - base) // 2, len(x) - 1)]),
    )


def text_properties(size: int, bold: bool) -> RichText:
    props = CharacterProperties(sz=size, b=bold)
    return RichText(bodyPr=RichTextProperties(),
                    p=[Paragraph(pPr=ParagraphProperties(defRPr=props), endParaRPr=props)])


def emphasise_title(title_obj, size: int, bold: bool = True) -> None:
    if title_obj is None:
        return
    title_obj.overlay = False
    rich = getattr(getattr(title_obj, "tx", None), "rich", None)
    if rich is None:
        return
    props = CharacterProperties(sz=size, b=bold)
    for paragraph in rich.p:
        paragraph.pPr = ParagraphProperties(defRPr=props)
        for run in paragraph.r or []:
            run.rPr = props


def style_axis(axis, title: str, number_format: str) -> None:
    axis.delete = False
    axis.title = title
    axis.number_format = number_format
    axis.crosses = "autoZero"
    axis.tickLblPos = "low"
    axis.majorTickMark = "none"
    axis.minorTickMark = "none"
    axis.spPr = GraphicalProperties(
        ln=LineProperties(solidFill=AXIS_COLOR, w=AXIS_WIDTH_EMU))
    axis.txPr = text_properties(TICK_LABEL_SIZE, bold=True)
    emphasise_title(axis.title, AXIS_TITLE_SIZE)
    if axis.majorGridlines is not None:
        axis.majorGridlines.spPr = GraphicalProperties(
            ln=LineProperties(solidFill=GRID_COLOR, w=GRID_WIDTH_EMU))


def make_scatter(title: str, x_title: str, y_title: str) -> ScatterChart:
    chart = ScatterChart()
    chart.title = title
    chart.style = 13
    style_axis(chart.x_axis, x_title, "0")
    style_axis(chart.y_axis, y_title, "0.0")
    chart.visible_cells_only = False
    chart.width, chart.height = CHART_W_CM, CHART_H_CM
    emphasise_title(chart.title, CHART_TITLE_SIZE)
    if chart.legend is not None:
        chart.legend.position = "r"
        chart.legend.overlay = False
        chart.legend.txPr = text_properties(TICK_LABEL_SIZE, bold=False)
    return chart


def add_series(chart: ScatterChart, ws: Worksheet, x_col: int, y_col: int,
               first_row: int, last_row: int, name: str, color: str,
               width: int, dashed: bool = False) -> None:
    x_ref = Reference(ws, min_col=x_col, min_row=first_row, max_row=last_row)
    y_ref = Reference(ws, min_col=y_col, min_row=first_row, max_row=last_row)
    s = Series(y_ref, x_ref, title=name)
    s.marker = Marker(symbol="none")
    s.smooth = False
    line = LineProperties(solidFill=color, w=width)
    if dashed:
        line.prstDash = "dash"
    s.graphicalProperties = GraphicalProperties(ln=line)
    chart.series.append(s)


def put(ws: Worksheet, row: int, col: int, value, font=DATA_FONT, fill=None):
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = font
    cell.alignment = CENTRE
    if fill is not None:
        cell.fill = fill
    return cell


def autofit(ws: Worksheet, columns, rows, minimum: float = 10.0) -> None:
    for col in columns:
        longest = 0
        for row in rows:
            value = ws.cell(row=row, column=col).value
            if value is not None:
                longest = max(longest, len(str(value)))
        letter = get_column_letter(col)
        current = ws.column_dimensions[letter].width or 0
        ws.column_dimensions[letter].width = max(current, minimum, longest + 3)


def label(a: Analysis) -> str:
    bits = [a.name]
    sc = a.profile.coatings or a.config.coatings
    if sc:
        bits.append(f"{sc} layers")
    if a.config.additive:
        value = a.config.additive.strip()
        if value.upper() in ("ΝΑΙ", "NAI", "YES", "Y"):
            bits.append("with additive")
        elif value.upper() in ("ΟΧΙ", "OXI", "NO", "N"):
            bits.append("no additive")
        else:
            bits.append(value)
    if a.config.age:
        bits.append(a.config.age)
    return " · ".join(bits)


SUMMARY_COLUMNS = [
    ("Sample",             10, lambda a: a.name),
    ("Alias",              10, lambda a: a.config.alias or MISSING),
    ("Layers",             14, lambda a: a.profile.coatings or a.config.coatings or MISSING),
    ("Additive",           15, lambda a: a.config.additive or MISSING),
    ("Solution age",       18, lambda a: a.config.age or MISSING),
    ("RPM",                 9, lambda a: a.profile.rpm or a.config.rpm or MISSING),
    ("Thickness (nm)",     13, lambda a: round(a.thickness_um * 1000, 2)),
    ("Thickness (µm)",     13, lambda a: round(a.thickness_um, 6)),
    ("Film mean (µm)",     15, lambda a: round(a.film_mean, 6)),
    ("Baseline mean (µm)", 18, lambda a: round(a.baseline_mean, 6)),
    ("Without baseline (nm)", 19, lambda a: round(a.raw_um * 1000, 2)),
    ("Window X0 (µm)",     17, lambda a: round(float(a.profile.x_um[a.start]), 1)),
    ("Window X1 (µm)",     17, lambda a: round(float(a.profile.x_um[a.stop]), 1)),
    ("Points",              9, lambda a: a.stop - a.start + 1),
    ("Roughness (nm)",     16, lambda a: round(a.roughness * 1000, 2)),
    ("Slope (nm/µm)",      15, lambda a: round(a.slope_nm_per_um, 4)),
    ("Edge peak (µm)",     18, lambda a: round(a.peak_x, 1)),
    ("Sample tilt (nm/µm)", 24, lambda a: round(a.tilt_nm_per_um, 4)),
    ("Tilt over the scan (nm)", 27, lambda a: round(a.tilt_total_nm, 1)),
    ("Tilt uncertainty (nm)", 24, lambda a: round(abs(a.tilt_error_nm), 1)),
    ("Levelling",          14, lambda a: "done here" if a.levelled else "as measured"),
    ("Window choice",      12, lambda a: "manual" if a.manual
                                          else ("relaxed" if a.relaxed > 0
                                                else ("fallback" if a.relaxed < 0 else "automatic"))),
    ("File",               40, lambda a: a.profile.source.name),
]


def build_all_sheet(wb: openpyxl.Workbook, results: list[Analysis]) -> None:
    ws = wb.create_sheet("ALL")
    put(ws, 1, 1, f"Film thickness — {len(results)} samples", font=TITLE_FONT)
    for col, (header, width, _) in enumerate(SUMMARY_COLUMNS, start=1):
        put(ws, 2, col, header, font=HEADER_FONT, fill=HEADER_FILL)
        ws.column_dimensions[get_column_letter(col)].width = width
    for i, a in enumerate(results):
        for col, (_, _, getter) in enumerate(SUMMARY_COLUMNS, start=1):
            put(ws, 3 + i, col, getter(a))
    autofit(ws, range(1, len(SUMMARY_COLUMNS) + 1), range(2, 3 + len(results)))
    ws.freeze_panes = "A3"

    note = ws.cell(row=3 + len(results) + 1, column=1,
                   value="Thickness = mean height in the window − substrate level. "
                         "Automatic detection is not always accurate: check every "
                         "sample's chart, and if the red window is misplaced rerun "
                         "with --window SAMPLE=X0:X1.")
    note.font = Font(italic=True, size=9)

    if len(results) < 2:
        return
    base = len(SUMMARY_COLUMNS) + 2
    put(ws, 2, base, "Sample", font=HEADER_FONT, fill=HEADER_FILL)
    put(ws, 2, base + 1, "Thickness (nm)", font=HEADER_FONT, fill=HEADER_FILL)
    for col in (base, base + 1):
        ws.column_dimensions[get_column_letter(col)].width = 16
    for i, a in enumerate(results):
        put(ws, 3 + i, base, a.name)
        put(ws, 3 + i, base + 1, round(a.thickness_um * 1000, 2))

    bar = BarChart()
    bar.type = "col"
    bar.title = "Thickness per sample"
    style_axis(bar.x_axis, None, "General")
    style_axis(bar.y_axis, "Thickness (nm)", "0")
    bar.y_axis.scaling.min = 0
    bar.varyColors = False
    bar.visible_cells_only = False
    bar.gapWidth = 60
    bar.width, bar.height = CHART_W_CM, CHART_H_CM
    emphasise_title(bar.title, CHART_TITLE_SIZE)
    last = 2 + len(results)
    bar.add_data(Reference(ws, min_col=base + 1, min_row=2, max_row=last),
                 titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=base, min_row=3, max_row=last))
    bar.legend = None
    ws.add_chart(bar, f"{get_column_letter(base + 3)}2")


def build_sample_sheet(wb: openpyxl.Workbook, a: Analysis, color: str) -> None:
    ws = wb.create_sheet(a.name[:31])
    x, z = a.profile.x_um, a.z_um

    rows = [
        ("Sample", label(a)),
        ("Thickness (nm)", round(a.thickness_um * 1000, 2)),
        ("Thickness (µm)", round(a.thickness_um, 6)),
        ("Film mean (µm)", round(a.film_mean, 6)),
        ("Baseline mean (µm)", round(a.baseline_mean, 6)),
        ("Window (µm)", f"{x[a.start]:.1f} – {x[a.stop]:.1f}"),
        ("Edge peak (µm)", round(a.peak_x, 1)),
        ("Levelling", "done here" if a.levelled else "as measured"),
        ("Sample tilt", f"{a.tilt_nm_per_um:+.3f} nm/µm"),
        ("File", a.profile.source.name),
    ]
    for i, (name, value) in enumerate(rows):
        put(ws, 1 + i, 1, name, font=META_FONT, fill=META_FILL)
        put(ws, 1 + i, 2, value, font=VALUE_FONT)
    autofit(ws, (1, 2), range(1, len(rows) + 1))

    header_row = len(rows) + 2
    first_row = header_row + 1
    headers = ["X (µm)", "Z (µm)", "Window (µm)"]
    for col, name in enumerate(headers, start=1):
        put(ws, header_row, col, name, font=HEADER_FONT, fill=HEADER_FILL)
    for col in range(1, 4):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = max(
            ws.column_dimensions[letter].width or 0, 16)

    for i in range(len(x)):
        row = first_row + i
        put(ws, row, 1, round(float(x[i]), 4))
        put(ws, row, 2, round(float(z[i]), 6))
        if a.start <= i <= a.stop:
            put(ws, row, 3, round(float(z[i]), 6))
    last_row = first_row + len(x) - 1

    ref_col = 5
    put(ws, header_row, ref_col, "X ends", font=HEADER_FONT, fill=HEADER_FILL)
    put(ws, header_row, ref_col + 1, "Baseline", font=HEADER_FONT, fill=HEADER_FILL)
    put(ws, header_row, ref_col + 2, "Film", font=HEADER_FONT, fill=HEADER_FILL)
    for col in range(ref_col, ref_col + 3):
        ws.column_dimensions[get_column_letter(col)].width = 14
    for offset, xv in enumerate((float(x[0]), float(x[-1]))):
        put(ws, first_row + offset, ref_col, round(xv, 4))
        put(ws, first_row + offset, ref_col + 1, round(a.baseline_mean, 6))
        put(ws, first_row + offset, ref_col + 2, round(a.film_mean, 6))

    chart = make_scatter(
        f"{label(a)}  —  thickness {a.thickness_um * 1000:.1f} nm",
        "Position (µm)", "Height (µm)")
    add_series(chart, ws, 1, 2, first_row, last_row, "Profile",
               PROFILE_COLOR, PROFILE_WIDTH)
    add_series(chart, ws, ref_col, ref_col + 1, first_row, first_row + 1,
               "Substrate level", LEVEL_COLOR, LEVEL_WIDTH, dashed=True)
    add_series(chart, ws, ref_col, ref_col + 2, first_row, first_row + 1,
               "Film level", color, LEVEL_WIDTH, dashed=True)
    add_series(chart, ws, 1, 3, first_row, last_row,
               f"Measured window ({x[a.start]:.0f}–{x[a.stop]:.0f} µm)",
               WINDOW_COLOR, WINDOW_WIDTH)
    ws.add_chart(chart, f"{get_column_letter(ref_col + 4)}{header_row}")

    zoom = make_scatter(f"{a.name} — zoom on the measured window",
                        "Position (µm)", "Height (µm)")
    pad = (x[a.stop] - x[a.start]) * 3
    zoom.x_axis.scaling.min = round(max(float(x[0]), float(x[a.start]) - pad), 1)
    zoom.x_axis.scaling.max = round(min(float(x[-1]), float(x[a.stop]) + pad), 1)
    zoom.y_axis.scaling.min = round(min(a.baseline_mean, a.film_mean) - 0.2, 2)
    zoom.y_axis.scaling.max = round(max(a.baseline_mean, a.film_mean) + 0.4, 2)
    add_series(zoom, ws, 1, 2, first_row, last_row, "Profile",
               PROFILE_COLOR, PROFILE_WIDTH)
    add_series(zoom, ws, ref_col, ref_col + 1, first_row, first_row + 1,
               "Substrate level", LEVEL_COLOR, LEVEL_WIDTH, dashed=True)
    add_series(zoom, ws, ref_col, ref_col + 2, first_row, first_row + 1,
               "Film level", color, LEVEL_WIDTH, dashed=True)
    add_series(zoom, ws, 1, 3, first_row, last_row, "Measured window",
               WINDOW_COLOR, WINDOW_WIDTH)
    ws.add_chart(zoom, f"{get_column_letter(ref_col + 4)}{header_row + CHART_ROW_SPAN}")


def build_workbook(results: list[Analysis]) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_all_sheet(wb, results)
    for i, a in enumerate(results):
        build_sample_sheet(wb, a, PALETTE[i % len(PALETTE)])
    return wb


def collect_files(inputs: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in inputs:
        p = Path(raw).expanduser()
        if p.is_dir():
            files.extend(sorted(p.rglob("*.dat")))
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


def parse_windows(items: list[str] | None) -> dict[str, tuple[float, float]]:
    out: dict[str, tuple[float, float]] = {}
    for item in items or []:
        m = re.fullmatch(r"\s*([A-Za-z]*\d+)\s*=\s*([\d.,]+)\s*:\s*([\d.,]+)\s*", item)
        if not m:
            raise argparse.ArgumentTypeError(
                f"--window '{item}' is not of the form SAMPLE=X0:X1 (e.g. S12=1221:1317)")
        lo, hi = to_float(m.group(2)), to_float(m.group(3))
        if lo is None or hi is None or hi <= lo:
            raise argparse.ArgumentTypeError(f"Invalid range in '{item}'")
        out[m.group(1).upper()] = (lo, hi)
    return out


def sort_key(a: Analysis) -> tuple:
    m = re.match(r"([A-Z]+)(\d+)", a.name)
    return (m.group(1), int(m.group(2))) if m else (a.name, 0)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Film thickness from stylus profilometry, with automatic plateau detection.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python thickness_analysis.py data/Thickness
  python thickness_analysis.py data/Thickness --window S12=1221:1317 --window S13=1721:1817
""")
    parser.add_argument("inputs", nargs="+", help="Folder with .dat files (searched recursively), or single files.")
    parser.add_argument("-o", "--out", default=DEFAULT_OUTPUT, help="Name of the output file.")
    parser.add_argument("--outdir", help=f"Output folder (default: '{DEFAULT_OUTPUT_DIR}').")
    parser.add_argument("--samples", default=DEFAULT_SAMPLE_TABLE,
                        help=f"Sample table (default: {DEFAULT_SAMPLE_TABLE}).")
    parser.add_argument("--window", action="append", metavar="SAMPLE=X0:X1",
                        help="Manual measuring window for one sample, in micrometres.")
    parser.add_argument("--level", action="store_true",
                        help="Level profiles with a large tilt "
                             "(experimental — see the README).")
    args = parser.parse_args(argv)

    windows = parse_windows(args.window)
    files = collect_files(args.inputs)
    if not files:
        log("No .dat files found.")
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
             f"sample properties will be left blank.")

    log(f"Files to process: {len(files)}")
    results: list[Analysis] = []
    for path in files:
        try:
            profile = read_profile(path, table)
            cfg = lookup_sample(profile.sample, table) or SampleConfig()
            manual = windows.get(profile.sample.upper())
            if manual is None and cfg.id:
                manual = windows.get(cfg.id.upper())
            results.append(analyse(profile, cfg, manual, args.level))
        except Exception as exc:
            warn(f"Skipping {path.name}: {exc}")
    if not results:
        log("No file was processed successfully.")
        return 1

    results.sort(key=sort_key)
    for a in results:
        if abs(a.tilt_error_nm) > abs(a.thickness_um * 1000) * 0.3:
            warn(f"{a.name}: sample tilt {a.tilt_nm_per_um:+.3f} nm/µm "
                 f"→ ±{abs(a.tilt_error_nm):.0f} nm uncertainty on a thickness of "
                 f"{a.thickness_um*1000:.0f} nm. Check the chart.")
        if a.relaxed > 0:
            warn(f"{a.name}: the thresholds had to be relaxed "
                 f"×{1.5 ** a.relaxed:.1f} — check the chart.")
        elif a.relaxed < 0:
            warn(f"{a.name}: no quiet zone found; used the "
                 f"smoothest window after the edge — check the chart.")

    out_dir = Path(args.outdir).expanduser() if args.outdir else (
        search_up(DEFAULT_OUTPUT_DIR, starts, want_dir=True)
        or script_dir.parent / DEFAULT_OUTPUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / Path(args.out).name
    if out_path.suffix.lower() != ".xlsx":
        out_path = out_path.with_suffix(".xlsx")

    wb = build_workbook(results)
    try:
        wb.save(out_path)
    except PermissionError:
        log(f"Could not write {out_path.name} — is it open in Excel?")
        return 1

    log("")
    log(f"Saved: {out_path}")
    for a in results:
        x = a.profile.x_um
        flag = "" if a.relaxed == 0 else ("  [relaxed]" if a.relaxed > 0 else "  [fallback]")
        if a.manual:
            flag = "  [manual]"
        log(f"    {a.name:<5} {a.thickness_um * 1000:>8.1f} nm   "
            f"window {x[a.start]:>7.1f}–{x[a.stop]:<7.1f} µm   "
            f"(edge at {a.peak_x:.0f}){flag}")
    log("")
    log("  Note: the automatic plateau detection is not always accurate. Check each")
    log("  sample's chart; if the red window is in the wrong place, rerun with")
    log("  --window SAMPLE=X0:X1 (in µm) to set it by hand.")
    return 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
