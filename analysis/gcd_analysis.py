#!/usr/bin/env python3
from __future__ import annotations

import argparse
import math
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import numpy as np
import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
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


CENTRE = Alignment(horizontal="center", vertical="center")

TITLE_FONT = Font(bold=True, size=12)
META_FILL = PatternFill("solid", fgColor="1F4E79")
META_FONT = Font(bold=True, color="FFFFFF", size=11)
VALUE_FONT = Font(bold=True, size=11)
HEADER_FILL = PatternFill("solid", fgColor="D6E4F0")
HEADER_FONT = Font(bold=True, size=10)
GROUP_FILL = PatternFill("solid", fgColor="EAF1F8")
GROUP_FONT = Font(bold=True, size=11)
DATA_FONT = Font(size=10)

PALETTE = ["1F77B4", "FF7F0E", "2CA02C", "D62728", "9467BD",
           "8C564B", "E377C2", "7F7F7F", "BCBD22", "17BECF"]

LINE_WIDTH_EMU = 22225
CHART_W_CM, CHART_H_CM = 26, 15
CHART_ROW_SPAN = int(CHART_H_CM / 0.53) + 2

ZOOM_MARGIN = 1.15

AXIS_COLOR = "000000"
AXIS_WIDTH_EMU = 28575
GRID_COLOR = "E6E6E6"
GRID_WIDTH_EMU = 6350
TICK_LABEL_SIZE = 1200
AXIS_TITLE_SIZE = 1400
CHART_TITLE_SIZE = 1600

ROUND_TIME = 6
ROUND_POTENTIAL = 6
ROUND_CURRENT = 7

PAPER_PALETTE = ["#515151", "#F14040", "#1A6FDF", "#37AD6B", "#B177DE",
                 "#FB6501", "#6E2FA0", "#00A0A0", "#C1272D", "#7F7F7F"]
PAPER_DPI = 300
PAPER_AXES = (8.4, 6.2)
PAPER_MARGIN = (1.45, 0.55, 1.20, 1.15)
PAPER_CURVE_WIDTH = 3.4
PAPER_AXIS_WIDTH = 2.4
PAPER_TICK_SIZE = 21
PAPER_LABEL_SIZE = 25
PAPER_LEGEND_SIZE = 20
PAPER_PANEL_SIZE = 30
PAPER_X_LABEL = "time (s)"
PAPER_Y_LABEL = "V (V)"

XL_PAPER_AXIS_EMU = 22000
XL_PAPER_CURVE_EMU = 34000
XL_PAPER_TICK_SIZE = 1400
XL_PAPER_AXIS_TITLE_SIZE = 1600
XL_PAPER_LEGEND_SIZE = 1400

DEFAULT_SAMPLE_TABLE = "samples.xlsx"
DEFAULT_OUTPUT_DIR = "results"
SEARCH_LEVELS = 4


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


def as_number(text):
    s = str(text or "").strip()
    if not s:
        return ""
    value = to_float(s)
    if value is None:
        return s
    return int(value) if value.is_integer() else value


NUMBER_PREFIX = "V"


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
    def key(self) -> str:
        return self.id


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
                for field_name, keys in TABLE_FIELDS.items():
                    if field_name not in found and any(k in norm for k in keys):
                        found[field_name] = col
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
                additive=_text(cell("additive")),
                age=_text(cell("age")),
                dimensions=_text(cell("dims")),
                thickness=_text(cell("thickness")),
                broken=bool(_text(cell("broken"))),
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
    aliases = {id_key(cfg.alias): cfg for cfg in table.values() if cfg.alias}
    if key in aliases:
        return aliases[key]
    if key.isdigit():
        return table.get(id_key(f"{NUMBER_PREFIX}{key}"))
    return None


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


@dataclass
class FileMeta:
    sample: str
    density: float | None
    discharge: bool
    electrolyte: str | None = None
    date: str | None = None


DATE_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")
DENSITY_RE = re.compile(r"([\d.,]+)[\s_]*m\s*A[\s_/-]*(?:per[\s_]*)?cm", re.IGNORECASE)
FORMULA_RE = re.compile(r"(?:[A-Z][a-z]?\d*|\([A-Za-z0-9]+\)\d*)+")


def parse_filename(path: Path) -> FileMeta:
    stem = path.stem.strip()
    discharge = "discharge" in stem.lower()

    cleaned = re.sub(r"discharge", " ", stem, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s+", " ", cleaned)

    m = DENSITY_RE.search(cleaned)
    density = to_float(m.group(1)) if m else None
    if density is None:
        warn(f"No current density found in '{path.name}'.")

    date = None
    date_m = DATE_RE.search(stem)
    if date_m:
        try:
            date = datetime.strptime(date_m.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            pass

    parts = [p for p in re.split(r"[_]+", cleaned) if p.strip()]
    if parts and parts[0].strip().upper() in ("GCD", "CD"):
        parts = parts[1:]
    sample = parts[0].strip() if parts else stem
    electrolyte = None
    for part in parts[1:4]:
        token = part.strip()
        if token and not re.fullmatch(r"\d+\s*SC", token, re.IGNORECASE) \
                and not DENSITY_RE.search(token) and FORMULA_RE.fullmatch(token):
            electrolyte = token
            break

    return FileMeta(sample=sample, density=density, discharge=discharge,
                    electrolyte=electrolyte, date=date)


def _norm(text) -> str:
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


RAW_COLUMNS = {
    "time":      ["correctedtimes", "correctedtime"],
    "potential": ["we1potentialv", "wepotentialv", "potentialv"],
    "current":   ["we1currenta", "wecurrenta", "currenta"],
}
FALLBACK_TIME = ["times", "time"]


@dataclass
class Run:
    density: float
    discharge: bool
    time: np.ndarray
    potential: np.ndarray
    current: np.ndarray
    source: Path

    def __len__(self) -> int:
        return len(self.time)

    @property
    def duration(self) -> float:
        return float(self.time[-1]) if len(self.time) else 0.0


def read_raw(path: Path, meta: FileMeta) -> Run:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        ws = wb.worksheets[0]
        ws.reset_dimensions()
        rows = ws.iter_rows(values_only=True)
        try:
            header = next(rows)
        except StopIteration:
            raise ValueError(f"{path.name} is empty.")

        index = {_norm(h): c for c, h in enumerate(header) if h is not None}
        col: dict[str, int] = {}
        for key, candidates in RAW_COLUMNS.items():
            for cand in candidates:
                if cand in index:
                    col[key] = index[cand]
                    break
        if "time" not in col:
            for cand in FALLBACK_TIME:
                if cand in index:
                    col["time"] = index[cand]
                    break
        missing = [k for k in ("time", "potential", "current") if k not in col]
        if missing:
            raise ValueError(
                f"{path.name} is missing columns {missing}. "
                f"Found: {[h for h in header if h is not None]}"
            )

        data = []
        for row in rows:
            values = (row[col["time"]], row[col["potential"]], row[col["current"]])
            if any(v is None for v in values):
                continue
            data.append(tuple(float(v) for v in values))
    finally:
        wb.close()

    if not data:
        raise ValueError(f"{path.name} contains no numeric data.")

    arr = np.asarray(data, dtype=float)
    time = arr[:, 0]
    if time[0] != 0:
        time = time - time[0]
    return Run(density=meta.density, discharge=meta.discharge, time=time,
               potential=arr[:, 1], current=arr[:, 2], source=path)


@dataclass
class DensityGroup:
    density: float
    charge: Run | None = None
    discharge: Run | None = None

    @property
    def label(self) -> str:
        return f"{self.density:.1f} mA/cm²"

    @property
    def sheet_name(self) -> str:
        return f"{self.density:.1f}".replace(".", "p") + "mAcm2"

    @property
    def rows(self) -> int:
        return max(len(self.charge or ()), len(self.discharge or ()))


SLOPE_JUMP_FACTOR = 8.0
MIN_LINEAR_POINTS = 4


def find_linear_end(run: Run) -> int:
    t, v = run.time, run.potential
    last = len(t) - 1
    if last < 3:
        return last

    dt = np.diff(t)
    valid = dt > 0
    if not valid.any():
        return last
    slope = np.zeros_like(dt)
    slope[valid] = np.abs(np.diff(v)[valid] / dt[valid])

    peak = int(np.argmax(slope))

    before = slope[:peak] if peak > 0 else slope
    before = before[before > 0]
    if len(before) == 0:
        return last
    baseline = float(np.quantile(before, 0.25))
    if baseline <= 0:
        baseline = float(np.median(before))
    if baseline <= 0:
        return last

    threshold = SLOPE_JUMP_FACTOR * baseline
    if slope[peak] <= threshold:
        return last

    end = peak
    while end > 0 and slope[end - 1] > threshold:
        end -= 1
    return end


@dataclass
class Capacitance:
    density: float
    current_mA: float
    dt_linear: float
    dt_total: float
    points_linear: int
    points_total: int
    v_start: float
    v_end: float
    delta_v: float
    ir_drop: float | None
    c_areal_mF_cm2: float | None

    @property
    def trimmed(self) -> bool:
        return self.points_linear < self.points_total


def compute_capacitance(group: DensityGroup, area_cm2: float) -> Capacitance | None:
    run = group.discharge
    if run is None or len(run) < 2 or not area_cm2:
        return None

    end = max(find_linear_end(run), 1)

    current = float(np.median(np.abs(run.current[:end + 1])))
    dt = float(run.time[end] - run.time[0])
    v_start, v_end = float(run.potential[0]), float(run.potential[end])
    delta_v = abs(v_end - v_start)

    c_areal = None
    if end + 1 < MIN_LINEAR_POINTS:
        warn(f"{group.label}: the linear part of the discharge has only "
             f"{end + 1} points — capacitance is not calculated.")
    elif delta_v > 0 and dt > 0 and current > 0:
        c_areal = current * dt / (delta_v * area_cm2) * 1000.0

    ir_drop = None
    charge = group.charge
    if charge is not None and len(charge) > 1:
        ir_drop = abs(v_start - float(charge.potential[-1]))

    return Capacitance(density=group.density, current_mA=current * 1000.0,
                       dt_linear=dt, dt_total=run.duration,
                       points_linear=end + 1, points_total=len(run),
                       v_start=v_start, v_end=v_end, delta_v=delta_v,
                       ir_drop=ir_drop, c_areal_mF_cm2=c_areal)


@dataclass
class SampleRun:
    key: str
    config: SampleConfig
    electrolyte: str
    date: str
    groups: list[DensityGroup] = field(default_factory=list)
    capacitance: list[Capacitance] = field(default_factory=list)


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


def style_axis(axis, title: str, number_format: str, style: str = "report") -> None:
    axis.delete = False
    axis.title = title
    axis.number_format = number_format

    if style == "paper":
        axis.crosses = "min"
        axis.tickLblPos = "nextTo"
        axis.majorTickMark, axis.minorTickMark = "out", "none"
        axis.spPr = GraphicalProperties(
            ln=LineProperties(solidFill="000000", w=XL_PAPER_AXIS_EMU))
        axis.txPr = text_properties(XL_PAPER_TICK_SIZE, bold=True)
        axis.majorGridlines = None
        emphasise_title(axis.title, XL_PAPER_AXIS_TITLE_SIZE)
        return

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


def make_scatter(title: str, x_title: str, y_title: str,
                 style: str = "report") -> ScatterChart:
    chart = ScatterChart()
    paper = style == "paper"
    if not paper:
        chart.title = title
    chart.style = None if paper else 13
    style_axis(chart.x_axis, x_title if not paper else PAPER_X_LABEL, "0", style)
    style_axis(chart.y_axis, y_title if not paper else PAPER_Y_LABEL, "0.0", style)
    chart.visible_cells_only = False
    chart.width, chart.height = CHART_W_CM, CHART_H_CM
    emphasise_title(chart.title, CHART_TITLE_SIZE)

    if paper:
        chart.y_axis.scaling.orientation = "maxMin"
        chart.plot_area.spPr = GraphicalProperties(
            noFill=True,
            ln=LineProperties(solidFill="000000", w=XL_PAPER_AXIS_EMU))
        chart.graphical_properties = GraphicalProperties(
            solidFill="FFFFFF", ln=LineProperties(noFill=True))
        if chart.legend is not None:
            chart.legend.position = "t"
            chart.legend.overlay = False
            chart.legend.txPr = text_properties(XL_PAPER_LEGEND_SIZE, bold=True)
            chart.legend.spPr = GraphicalProperties(
                noFill=True, ln=LineProperties(noFill=True))
        return chart

    if chart.legend is not None:
        chart.legend.position = "r"
        chart.legend.overlay = False
        chart.legend.txPr = text_properties(TICK_LABEL_SIZE, bold=False)
    return chart


def add_xy_series(chart: ScatterChart, ws: Worksheet, x_col: int, y_col: int,
                  first_row: int, last_row: int, name: str, color: str,
                  style: str = "report") -> None:
    x_ref = Reference(ws, min_col=x_col, min_row=first_row, max_row=last_row)
    y_ref = Reference(ws, min_col=y_col, min_row=first_row, max_row=last_row)
    s = Series(y_ref, x_ref, title=name)
    s.marker = Marker(symbol="none")
    s.smooth = True
    s.graphicalProperties.line.solidFill = color
    s.graphicalProperties.line.width = (XL_PAPER_CURVE_EMU if style == "paper"
                                        else LINE_WIDTH_EMU)
    chart.series.append(s)


def paper_ylim(curves: list[tuple[str, list[float], list[float]]]
               ) -> tuple[float, float] | None:
    values = [v for _, _, vs in curves for v in vs if v is not None]
    if not values:
        return None
    low, high = min(values), max(values)
    if low >= 0 or high <= 0:
        return None
    return math.floor(low * 10 - 0.5) / 10, 0.0


def render_paper(curves: list[tuple[str, list[float], list[float]]], path: Path,
                 panel: str = "", x_max: float | None = None,
                 y_lim: tuple[float, float] | None = None) -> None:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    left, right, top, bottom = PAPER_MARGIN
    axes_w, axes_h = PAPER_AXES
    fig_w, fig_h = left + axes_w + right, bottom + axes_h + top
    fig = plt.figure(figsize=(fig_w, fig_h), dpi=PAPER_DPI, facecolor="white")
    ax = fig.add_axes([left / fig_w, bottom / fig_h, axes_w / fig_w, axes_h / fig_h])

    for i, (label, t, v) in enumerate(curves):
        ax.plot(t, v, color=PAPER_PALETTE[i % len(PAPER_PALETTE)],
                linewidth=PAPER_CURVE_WIDTH, solid_joinstyle="round",
                solid_capstyle="round", label=label, zorder=3)

    ax.set_xlabel(PAPER_X_LABEL, fontsize=PAPER_LABEL_SIZE, fontweight="bold",
                  labelpad=12)
    ax.set_ylabel(PAPER_Y_LABEL, fontsize=PAPER_LABEL_SIZE, fontweight="bold",
                  labelpad=12)
    ax.tick_params(axis="both", direction="out", length=10, width=PAPER_AXIS_WIDTH,
                   labelsize=PAPER_TICK_SIZE, colors="black")
    for label in list(ax.get_xticklabels()) + list(ax.get_yticklabels()):
        label.set_fontweight("bold")
    for spine in ax.spines.values():
        spine.set_linewidth(PAPER_AXIS_WIDTH)
        spine.set_color("black")

    ax.set_xlim(0, x_max)
    if y_lim:
        ax.set_ylim(*y_lim)
    ax.invert_yaxis()

    count = len(curves)
    legend = ax.legend(loc="lower center", bbox_to_anchor=(0.5, 1.01),
                       ncol=count if count <= 3 else 3,
                       fontsize=PAPER_LEGEND_SIZE, frameon=False,
                       columnspacing=1.1, handlelength=1.6, handletextpad=0.5,
                       labelspacing=0.35, borderpad=0.0)
    for text in legend.get_texts():
        text.set_fontweight("bold")
    for line in legend.get_lines():
        line.set_linewidth(PAPER_CURVE_WIDTH)
    if panel:
        fig.text(0.005, 0.995, f"({panel})", ha="left", va="top",
                 fontsize=PAPER_PANEL_SIZE, fontweight="bold")

    fig.savefig(path, dpi=PAPER_DPI, bbox_inches="tight", pad_inches=0.08,
                facecolor="white")
    plt.close(fig)


def finalize_scatter(chart: ScatterChart) -> None:
    if len(chart.series) < 2:
        chart.legend = None


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


def context_line(sample: SampleRun, extra: str = "") -> str:
    cfg = sample.config
    bits = [sample.key]
    if extra:
        bits.append(extra)
    if cfg.coatings:
        bits.append(f"Layers: {cfg.coatings}")
    if cfg.additive:
        bits.append(f"Additive: {cfg.additive}")
    if cfg.age:
        bits.append(f"Solution age: {cfg.age}")
    if cfg.area_cm2:
        bits.append(f"Area: {cfg.area_cm2:g} cm²")
    return " | ".join(bits)


def build_sample_info(wb: openpyxl.Workbook, sample: SampleRun) -> None:
    ws = wb.create_sheet("Sample_Info")
    cfg = sample.config
    put(ws, 1, 1, f"Sample {sample.key}", font=TITLE_FONT)

    rows = [
        ("Sample", cfg.id),
        ("Alias", cfg.alias),
        ("Layers", as_number(cfg.coatings)),
        ("Additive", cfg.additive),
        ("Solution age", cfg.age),
        ("Dimensions (cm)", cfg.dimensions or "—"),
        ("Area (cm²)", cfg.area_cm2),
        ("RPM", as_number(cfg.rpm)),
        ("Substrate thickness (mm)", as_number(cfg.thickness)),
        ("Electrolyte", sample.electrolyte or "—"),
        ("Date", sample.date or "—"),
    ]
    for offset, (label, value) in enumerate(rows):
        put(ws, 2 + offset, 1, label, font=META_FONT, fill=META_FILL)
        put(ws, 2 + offset, 2, value, font=VALUE_FONT)
    autofit(ws, (1, 2), range(1, 2 + len(rows)))


def zoom_limit(durations: list[float]) -> float | None:
    usable = sorted(d for d in durations if d > 0)
    if len(usable) < 3:
        return None
    limit = usable[1] * ZOOM_MARGIN
    if limit >= usable[-1] * 0.5:
        return None
    return round(limit, 1)


def build_chart_sheet(wb: openpyxl.Workbook, sample: SampleRun,
                      style: str = "report", png_dir: Path | None = None,
                      panel: str = "") -> None:
    ws = wb.create_sheet("Chart")
    densities = [g.density for g in sample.groups]
    span = (f"{min(densities):.1f}–{max(densities):.1f} mA/cm²"
            if densities else "—")
    heading = f"GCD {sample.key} – Charge & Discharge ({span})"
    parts = [heading]
    if sample.electrolyte:
        parts.append(sample.electrolyte)
    if sample.config.area_cm2:
        parts.append(f"Area: {sample.config.area_cm2:g} cm²")
    put(ws, 1, 1, "  |  ".join(parts), font=TITLE_FONT)

    header_row, first_row = 3, 4
    spans: list[tuple[DensityGroup, int, int, int, int, float]] = []

    for i, group in enumerate(sample.groups):
        t_col, v_col = 1 + i * 2, 2 + i * 2
        put(ws, header_row, t_col, f"{group.label} – Time (s)",
            font=HEADER_FONT, fill=HEADER_FILL)
        put(ws, header_row, v_col, f"{group.label} – Potential (V)",
            font=HEADER_FONT, fill=HEADER_FILL)
        ws.column_dimensions[get_column_letter(t_col)].width = 20
        ws.column_dimensions[get_column_letter(v_col)].width = 20

        row = first_row
        offset = 0.0
        for run in (group.charge, group.discharge):
            if run is None:
                continue
            for t, v in zip(run.time, run.potential):
                put(ws, row, t_col, round(float(t) + offset, ROUND_TIME))
                put(ws, row, v_col, round(float(v), ROUND_POTENTIAL))
                row += 1
            offset += run.duration
        if row > first_row:
            spans.append((group, i, t_col, v_col, row - 1, offset))

    palette = PAPER_PALETTE if style == "paper" else PALETTE

    def plot(title: str, x_max: float | None) -> ScatterChart:
        chart = make_scatter(title, "Time (s)", "Potential (V)", style)
        if x_max:
            chart.x_axis.scaling.min = 0
            chart.x_axis.scaling.max = x_max
        for group, i, t_col, v_col, last_row, _ in spans:
            add_xy_series(chart, ws, x_col=t_col, y_col=v_col,
                          first_row=first_row, last_row=last_row,
                          name=group.label, color=palette[i % len(palette)],
                          style=style)
        finalize_scatter(chart)
        return chart

    anchor = get_column_letter(len(sample.groups) * 2 + 2)
    ws.add_chart(plot(heading, None), f"{anchor}{header_row}")

    zoom = zoom_limit([total for *_, total in spans])
    if zoom:
        ws.add_chart(plot(f"{heading} — zoom 0–{zoom:g} s", zoom),
                     f"{anchor}{header_row + CHART_ROW_SPAN}")

    if png_dir is not None:
        curves = []
        for group, _, t_col, v_col, last_row, _ in spans:
            t = [ws.cell(r, t_col).value for r in range(first_row, last_row + 1)]
            v = [ws.cell(r, v_col).value for r in range(first_row, last_row + 1)]
            curves.append((group.label, t, v))
        png_dir.mkdir(parents=True, exist_ok=True)
        y_lim = paper_ylim(curves)
        if y_lim:
            log(f"    potential axis: {y_lim[0]:g} to {y_lim[1]:g} V "
                f"— the discharge tail above zero is not shown")
        render_paper(curves, png_dir / f"GCD {sample.key}.png", panel, y_lim=y_lim)
        if zoom:
            render_paper(curves, png_dir / f"GCD {sample.key} — zoom.png",
                         panel, x_max=zoom, y_lim=y_lim)


SUMMARY_COLUMNS = [
    ("Density (mA/cm²)", 20, lambda c: round(c.density, 3)),
    ("I (mA)",             12, lambda c: round(c.current_mA, 4)),
    ("Δt linear (s)",      17, lambda c: round(c.dt_linear, 3)),
    ("Δt total (s)",       17, lambda c: round(c.dt_total, 3)),
    ("Points linear",      17, lambda c: c.points_linear),
    ("Points total",       16, lambda c: c.points_total),
    ("V start (V)",        13, lambda c: round(c.v_start, 6)),
    ("V end (V)",          13, lambda c: round(c.v_end, 6)),
    ("ΔV (V)",             11, lambda c: round(c.delta_v, 6)),
    ("IR drop (V)",        14, lambda c: None if c.ir_drop is None
                                          else round(c.ir_drop, 6)),
    ("C (mF/cm²)",         14, lambda c: None if c.c_areal_mF_cm2 is None
                                          else round(c.c_areal_mF_cm2, 4)),
]


def build_summary(wb: openpyxl.Workbook, sample: SampleRun) -> None:
    ws = wb.create_sheet("Summary")
    put(ws, 1, 1, context_line(sample), font=TITLE_FONT)

    for col, (header, width, _) in enumerate(SUMMARY_COLUMNS, start=1):
        put(ws, 2, col, header, font=HEADER_FONT, fill=HEADER_FILL)
        ws.column_dimensions[get_column_letter(col)].width = width

    for i, cap in enumerate(sample.capacitance):
        for col, (_, _, getter) in enumerate(SUMMARY_COLUMNS, start=1):
            put(ws, 3 + i, col, getter(cap))
    autofit(ws, range(1, len(SUMMARY_COLUMNS) + 1),
            range(2, 3 + len(sample.capacitance)))

    note_row = 3 + len(sample.capacitance) + 1
    note = ws.cell(
        row=note_row, column=1,
        value="C = I · Δt / (ΔV · A)   —   from the linear part of the discharge "
              "only; the jump at the cut-off and the plateau after it are "
              "ignored. ΔV is measured from the 1st point, so the IR drop "
              "is already excluded.")
    note.font = Font(italic=True, size=9)


def build_all_currents(wb: openpyxl.Workbook, sample: SampleRun) -> None:
    ws = wb.create_sheet("All_Currents")
    put(ws, 1, 1, context_line(sample), font=TITLE_FONT)

    headers = ["Charge Time (s)", "Charge V (V)",
               "Discharge Time (s)", "Discharge V (V)"]
    for i, group in enumerate(sample.groups):
        base = 1 + i * (len(headers) + 1)
        put(ws, 2, base, group.label, font=GROUP_FONT, fill=GROUP_FILL)
        for offset, header in enumerate(headers):
            put(ws, 3, base + offset, header, font=HEADER_FONT, fill=HEADER_FILL)
            ws.column_dimensions[get_column_letter(base + offset)].width = 18

        for j, run in enumerate((group.charge, group.discharge)):
            if run is None:
                continue
            for k, (t, v) in enumerate(zip(run.time, run.potential)):
                put(ws, 4 + k, base + j * 2, round(float(t), ROUND_TIME))
                put(ws, 4 + k, base + j * 2 + 1, round(float(v), ROUND_POTENTIAL))
    ws.freeze_panes = "A4"


def build_density_sheets(wb: openpyxl.Workbook, sample: SampleRun) -> None:
    headers = [("Charge Time (s)", "Charge Potential (V)", "Charge Current (A)"),
               ("Discharge Time (s)", "Discharge Potential (V)", "Discharge Current (A)")]
    for group in sample.groups:
        ws = wb.create_sheet(group.sheet_name)
        put(ws, 1, 1, context_line(sample, group.label), font=TITLE_FONT)

        for j, (run, names) in enumerate(zip((group.charge, group.discharge), headers)):
            base = 1 + j * 4
            for offset, header in enumerate(names):
                put(ws, 2, base + offset, header, font=HEADER_FONT, fill=HEADER_FILL)
                ws.column_dimensions[get_column_letter(base + offset)].width = 20
            if run is None:
                continue
            for k in range(len(run)):
                put(ws, 3 + k, base, round(float(run.time[k]), ROUND_TIME))
                put(ws, 3 + k, base + 1, round(float(run.potential[k]), ROUND_POTENTIAL))
                put(ws, 3 + k, base + 2, round(float(run.current[k]), ROUND_CURRENT))
        ws.freeze_panes = "A3"


def build_workbook(sample: SampleRun, style: str = "report",
                   png_dir: Path | None = None, panel: str = "") -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_sample_info(wb, sample)
    build_chart_sheet(wb, sample, style, png_dir, panel)
    build_summary(wb, sample)
    build_all_currents(wb, sample)
    build_density_sheets(wb, sample)
    return wb


def collect_files(inputs: list[str]) -> list[Path]:
    files: list[Path] = []
    for raw in inputs:
        p = Path(raw).expanduser()
        if p.is_dir():
            files.extend(sorted(q for q in p.glob("*.xlsx")
                                if not q.name.startswith("~$")))
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


CURRENT_TOLERANCE = 0.2


def check_current(run: Run, area_cm2: float) -> None:
    if not area_cm2 or not run.density or not len(run):
        return
    nominal = run.density * area_cm2 / 1000.0
    measured = float(np.median(np.abs(run.current)))
    if nominal <= 0 or measured <= 0:
        return
    if abs(measured - nominal) / nominal > CURRENT_TOLERANCE:
        kind = "discharge" if run.discharge else "charge"
        warn(f"{run.density:g} mA/cm² ({kind}): measured current "
             f"{measured * 1000:.4f} mA but expected {nominal * 1000:.4f} mA "
             f"({measured / nominal:.2f}×) — {run.source.name}")


def build_sample(files: list[Path], table: dict[str, SampleConfig],
                 area_override: float | None,
                 electrolyte_override: str | None) -> SampleRun:
    metas = [(f, parse_filename(f)) for f in files]

    tokens = {m.sample for _, m in metas}
    if len(tokens) > 1:
        raise ValueError(
            f"The files belong to different samples: {sorted(tokens)}. "
            f"Give one folder per sample."
        )
    token = tokens.pop()

    cfg = lookup_sample(token, table) or SampleConfig()
    if not cfg.id and not area_override:
        raise ValueError(
            f"'{token}' is not in the sample table — neither as a Sample "
            f"nor as an Alias. Add it there or give --area."
        )
    if area_override:
        cfg.area_cm2 = area_override
    if not cfg.area_cm2:
        raise ValueError(
            f"Sample {cfg.id} ({cfg.alias or '—'}) has no 'Area (cm²)' "
            f"in the sample table. Fill it in or give --area."
        )
    if cfg.broken:
        warn(f"Sample {cfg.id} is marked as excluded in the sample table.")

    groups: dict[float, DensityGroup] = {}
    electrolyte, date = electrolyte_override, None
    for path, meta in metas:
        if meta.density is None:
            warn(f"Skipping {path.name}: unknown current density.")
            continue
        run = read_raw(path, meta)
        check_current(run, cfg.area_cm2)
        group = groups.setdefault(meta.density, DensityGroup(meta.density))
        slot = "discharge" if meta.discharge else "charge"
        if getattr(group, slot) is not None:
            warn(f"Two '{slot}' files for {meta.density:g} mA/cm²; "
                 f"keeping {getattr(group, slot).source.name}.")
            continue
        setattr(group, slot, run)
        electrolyte = electrolyte or meta.electrolyte
        date = date or meta.date

    if not groups:
        raise ValueError("No measurement could be read.")

    ordered = [groups[d] for d in sorted(groups)]
    for group in ordered:
        if group.charge is None:
            warn(f"{group.label}: the charge file is missing.")
        if group.discharge is None:
            warn(f"{group.label}: the discharge file is missing.")

    capacitance = [c for c in (compute_capacitance(g, cfg.area_cm2) for g in ordered)
                   if c is not None]

    return SampleRun(key=cfg.key or token, config=cfg,
                     electrolyte=electrolyte or "", date=date or "",
                     groups=ordered, capacitance=capacitance)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Galvanostatic charge–discharge (GCD) analysis, with ready-made Excel charts.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Examples:
  python gcd_analysis.py data/GCD/S12
  python gcd_analysis.py data/GCD/S12 data/GCD/S13
  python gcd_analysis.py data/GCD/S12 --area 4.72
""")
    parser.add_argument("inputs", nargs="+",
                        help="One folder per sample (or single .xlsx files).")
    parser.add_argument("-o", "--out", help="Name of the output file (for one sample).")
    parser.add_argument("--outdir",
                        help=f"Output folder (default: '{DEFAULT_OUTPUT_DIR}').")
    parser.add_argument("--samples", default=DEFAULT_SAMPLE_TABLE,
                        help=f"Sample table (default: {DEFAULT_SAMPLE_TABLE}).")
    parser.add_argument("--area", type=float, help="Area in cm²; overrides the sample table.")
    parser.add_argument("--electrolyte", help="Electrolyte, if the file names do not say.")
    parser.add_argument("--style", choices=("report", "paper"), default="report",
                        help="report: worksheet style (default). "
                             "paper: publication style — framed, no "
                             "gridlines, inverted vertical axis.")
    parser.add_argument("--panel", default="", help="Panel letter for the paper style, e.g. c.")
    parser.add_argument("--png", action="store_true",
                        help="Also save the charts as 300 dpi PNG images.")
    args = parser.parse_args(argv)

    script_dir = Path(__file__).resolve().parent
    groups = [Path(p).expanduser() for p in args.inputs]
    starts = [Path.cwd(), script_dir] + [g if g.is_dir() else g.parent for g in groups]

    table_path = Path(args.samples).expanduser()
    if not table_path.is_absolute():
        table_path = search_up(args.samples, starts) or (script_dir / args.samples)
    if not table_path.exists() and not args.area:
        log(f"Sample table not found ({args.samples}). "
            f"Give --samples <file> or --area <cm²>.")
        return 1
    table = read_sample_table(table_path) if table_path.exists() else {}
    if table:
        log(f"Sample table: {table_path.name}  ({len(table)} samples)")

    out_dir = Path(args.outdir).expanduser() if args.outdir else (
        search_up(DEFAULT_OUTPUT_DIR, starts, want_dir=True) or script_dir.parent / DEFAULT_OUTPUT_DIR)

    failures = 0
    for target in groups:
        files = collect_files([str(target)])
        if not files:
            warn(f"No .xlsx files in {target}")
            failures += 1
            continue
        log(f"\n{target.name}  ({len(files)} files)")
        try:
            sample = build_sample(files, table, args.area, args.electrolyte)
        except Exception as exc:
            warn(f"Skipping: {exc}")
            failures += 1
            continue

        name = args.out if (args.out and len(groups) == 1) \
            else f"GCD_{sample.key}_charge_discharge.xlsx"
        out_path = Path(name)
        if not out_path.is_absolute():
            out_dir.mkdir(parents=True, exist_ok=True)
            out_path = out_dir / out_path.name
        png_dir = out_path.with_suffix("") if args.png else None
        wb = build_workbook(sample, args.style, png_dir, args.panel)
        if out_path.suffix.lower() != ".xlsx":
            out_path = out_path.with_suffix(".xlsx")

        try:
            wb.save(out_path)
        except PermissionError:
            warn(f"Could not write {out_path.name} — is it open in Excel?")
            failures += 1
            continue

        log(f"  Saved: {out_path}")
        caps = {c.density: c for c in sample.capacitance}
        for group in sample.groups:
            bits = []
            if group.charge:
                bits.append(f"charge {len(group.charge)} points / {group.charge.duration:g} s")
            if group.discharge:
                bits.append(f"discharge {len(group.discharge)} points / "
                            f"{group.discharge.duration:g} s")
            cap = caps.get(group.density)
            if cap is not None and cap.c_areal_mF_cm2 is not None:
                bits.append(f"C = {cap.c_areal_mF_cm2:.3f} mF/cm²")
            log(f"    {group.label}: " + ("  ·  ".join(bits) if bits else "—"))
    return 1 if failures else 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    sys.exit(main())
