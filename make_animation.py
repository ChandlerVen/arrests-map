"""
Press Run in VS Code, choose the arrests spreadsheet in the window that opens, and
this writes an MP4 and a PNG of its last frame next to that spreadsheet. Arrests stay
on the map once they appear, so the map fills in over the period the spreadsheet covers.

The spreadsheet needs columns headed "Date of arrest", "ORI", "Latitude" and "Longitude".
Everything else on screen, including the dates, counts and key, is read from those columns.
"""

import json
import logging
import subprocess
import sys
import textwrap
from pathlib import Path

# Packages installed on the first run
REQUIREMENTS = {
    "numpy": "numpy>=1.24",
    "pandas": "pandas>=2.0",
    "geopandas": "geopandas>=1.0",
    "pyogrio": "pyogrio>=0.7",
    "matplotlib": "matplotlib>=3.8",
    "openpyxl": "openpyxl>=3.1",
    "imageio_ffmpeg": "imageio-ffmpeg>=0.4",
}
MINIMUM_PYTHON = (3, 10)


def install_requirements():
    from importlib import invalidate_caches
    from importlib.util import find_spec

    if sys.version_info < MINIMUM_PYTHON:
        sys.exit(
            f"This needs Python {'.'.join(map(str, MINIMUM_PYTHON))} or newer, "
            f"but it is running on {sys.version.split()[0]} ({sys.executable}).\n"
            "In VS Code, press Ctrl+Shift+P (Mac: Cmd+Shift+P), choose "
            '"Python: Select Interpreter", and pick a newer version.'
        )

    missing = [pkg for module, pkg in REQUIREMENTS.items() if find_spec(module) is None]
    if not missing:
        return
    print("Installing " + ", ".join(pkg.split(">")[0] for pkg in missing))
    print("This takes a few minutes.\n", flush=True)

    # Plain install first, then two flags
    for flags in ([], ["--user"], ["--break-system-packages"]):
        subprocess.call([sys.executable, "-m", "pip", "install", "--quiet", *flags, *missing])
        invalidate_caches()
        if all(find_spec(module) is not None for module in REQUIREMENTS):
            print()
            return
    sys.exit(
        "Could not install: " + ", ".join(missing) + "\nRun this in the terminal, then try again:\n"
        f"  {sys.executable} -m pip install " + " ".join(missing)
    )


install_requirements()

import geopandas as gpd
import matplotlib
import numpy as np
import pandas as pd

matplotlib.use("Agg")
import matplotlib.patheffects as patheffects
import matplotlib.pyplot as plt
from matplotlib.animation import FFMpegWriter
from matplotlib.colors import to_rgba_array
from matplotlib.lines import Line2D

TITLE = "Immigration Arrests Using Washington DOL Data\nin the Puget Sound Region"
SOURCE = "UW Center for Human Rights"
SHEET = "WA"  # which sheet to read when an Excel file has more than one
COLUMNS = {"date": "Date of arrest", "ori": "ORI", "lat": "Latitude", "lon": "Longitude"}
STYLE_FILE = Path(__file__).with_name("style.json")  # ORI colors and state names, kept between runs

EXTENT = (-123.0, -121.6, 46.95, 48.15)  # map view: west, east, south, north
WASHINGTON = (-124.9, -116.9, 45.5, 49.1)  # coordinates outside this are not in Washington
ORI_PATTERN = r"[A-Z]{2}[A-Z0-9]{7}"  # two-letter state, then seven letters or digits
NAMED_MINIMUM = 6  # a Washington ORI needs this many arrests to get its own color

PREVIEW = False  # True makes a quick low-resolution version for checking
FPS = 30
MOTION_SECONDS = 28  # length of the moving part; the speed adjusts to the date range
HOLD_SECONDS = 6  # freeze on the last frame at the end
FADE_FRAMES = 15  # how long a new arrest or key row takes to settle in

CITIES = {
    "Auburn": (-122.228, 47.307),
    "Bellevue": (-122.201, 47.610),
    "Bremerton": (-122.633, 47.567),
    "Burien": (-122.347, 47.470),
    "Everett": (-122.202, 47.979),
    "Kent": (-122.235, 47.381),
    "Lynnwood": (-122.315, 47.821),
    "Olympia": (-122.900, 47.038),
    "Renton": (-122.217, 47.482),
    "Seattle": (-122.335, 47.608),
    "Shoreline": (-122.341, 47.756),
    "Tacoma": (-122.444, 47.253),
}

# Colors given to ORIs that style.json does not have yet
SPARE_COLORS = ["#b8e1ff", "#7fd4a8", "#ff9f9f", "#d4b8ff", "#ffe08a", "#9fd8ff"]

# Agencies are read from letters 3-5 of the ORI, and appear in this order.
AGENCIES = {
    "CBP": ("o", "Customs & Border Protection (CBP) ORIs"),
    "CIS": ("D", "Citizenship & Immigration Services (CIS) ORIs"),
    "ICE": ("s", "Immigration & Customs Enforcement (ICE) ORIs"),
}
OTHER_SHAPE = "^"  # any agency code not listed above

BACKGROUND = "#141413"
LAND = "#2a2a28"
BORDER = "#45453f"
TEXT = "#f0efec"
TEXT_DIM = "#c3c2b7"
TEXT_FAINT = "#8a8984"
GRAY_DOT = "#a9a8a2"  # ORIs without a color of their own
FONTS = ["Inter", "Helvetica Neue", "Helvetica", "Arial", "DejaVu Sans"]
SIZE_DATE, SIZE_TITLE, SIZE_BODY, SIZE_KEY, SIZE_LABEL, SIZE_SMALL = 44, 26, 15, 13.5, 12.5, 11

FIGURE_INCHES = (19.2, 10.8)
DPI = 100
MARGIN = 0.02
MAP_BOTTOM, MAP_HEIGHT = 0.03, 0.94
PANEL_GAP = 0.035
TITLE_Y, DATE_Y, CAPTION_Y = 0.92, 0.795, 0.712
KEY_TOP, KEY_BOTTOM, KEY_GROUP_GAP = 0.632, 0.335, 0.014
KEY_ROW_HEIGHT = 0.0275
KEY_MARKER_X, KEY_LABEL_X, KEY_COUNT_X, KEY_COLUMN_X = 0.006, 0.02, 0.215, 0.262
PANEL_WIDTH = KEY_COLUMN_X + KEY_COUNT_X
CHART_BOTTOM, CHART_HEIGHT = 0.11, 0.165
ELSEWHERE_Y, SOURCE_Y = 0.055, 0.025

DOT_AREA = {True: 60, False: 34}
DOT_ALPHA = {True: 1.0, False: 0.6}
KEY_MARKER_SIZE = {True: 10, False: 8}
POP_GROWTH = 2.2  # how much a new arrest swells before settling
RING_GROWTH = 9  # how far its ring spreads before fading
HAIRLINE = 0.8  # county borders and marker outlines
CITY_LABEL_OFFSET = (0.012, 0.006)
CAPTION_WIDTH = 80
BAR_DAYS = 5.5
CHART_HEADROOM = 1.1


def load_style():
    """Colors and state names kept between runs."""
    try:
        style = json.loads(STYLE_FILE.read_text())
    except FileNotFoundError:
        print(f"{STYLE_FILE.name} is missing; colors may differ from other copies of the map.\n")
        style = {}
    except json.JSONDecodeError as problem:
        sys.exit(f"{STYLE_FILE.name} is not valid JSON ({problem}).\n"
                 "Fix the file, or delete it and the script will write a new one.")
    style.setdefault("ori_colors", {})
    style.setdefault("state_names", {})
    return style


def save_style(style):
    try:
        STYLE_FILE.write_text(json.dumps(style, indent=2, sort_keys=True) + "\n")
    except OSError as problem:
        print(f"Could not save {STYLE_FILE.name} ({problem}); new colors may change next run.\n")


def format_date(when, year=True):
    return f"{when:%B} {when.day}, {when.year}" if year else f"{when:%B} {when.day}"


def read_spreadsheet(path):
    """Read the sheet with the arrests on it, whether the file is Excel or CSV."""
    path = Path(path)
    if path.suffix.lower() not in (".xlsx", ".xls"):
        return pd.read_csv(path)
    sheets = pd.ExcelFile(path).sheet_names
    if SHEET in sheets:
        return pd.read_excel(path, sheet_name=SHEET)
    if len(sheets) == 1:
        return pd.read_excel(path, sheet_name=sheets[0])
    sys.exit(f"No sheet named '{SHEET}'. This file has: {sheets}")


def spreadsheet_value(value):
    """A cell as it appears in the spreadsheet, with hidden characters shown as codes."""
    if pd.isna(value):
        return "(blank)"
    text = f"{value:%Y-%m-%d}" if isinstance(value, pd.Timestamp) else str(value)
    return text.encode("ascii", "backslashreplace").decode()


def load_arrests(path):
    """Return the arrests that are formatted correctly, and a note about any that aren't."""
    raw = read_spreadsheet(path)
    missing = [column for column in COLUMNS.values() if column not in raw.columns]
    if missing:
        sys.exit(f"Missing column(s): {missing}. The spreadsheet has: {list(raw.columns)}")

    arrests = pd.DataFrame({
        "row": raw.index + 2,  # the row number in the spreadsheet, where row 1 is the headers
        "date": pd.to_datetime(raw[COLUMNS["date"]], errors="coerce").dt.normalize(),
        "ori": raw[COLUMNS["ori"]].astype(str).str.strip(),
        "lat": pd.to_numeric(raw[COLUMNS["lat"]], errors="coerce"),
        "lon": pd.to_numeric(raw[COLUMNS["lon"]], errors="coerce"),
    })
    west, east, south, north = WASHINGTON
    problems = pd.concat({
        "date": arrests.date.isna(),
        "ORI": ~arrests.ori.str.fullmatch(ORI_PATTERN),
        "latitude": ~arrests.lat.between(south, north),
        "longitude": ~arrests.lon.between(west, east),
    }, axis=1)
    skipped = problems.any(axis=1)

    notice = []
    if skipped.any():
        notice = [f"{skipped.sum()} rows were skipped because they are not formatted correctly:",
                  "  row   ORI | date | latitude, longitude"]
        for i in arrests.index[skipped]:
            values = " | ".join([
                spreadsheet_value(raw.at[i, COLUMNS["ori"]]),
                spreadsheet_value(raw.at[i, COLUMNS["date"]]),
                f"{spreadsheet_value(raw.at[i, COLUMNS['lat']])}, "
                f"{spreadsheet_value(raw.at[i, COLUMNS['lon']])}",
            ])
            notice.append(f"  {arrests.at[i, 'row']:<5} {values}")

    # A stable sort keeps same-day arrests in spreadsheet order so every run matches.
    kept = arrests[~skipped].sort_values("date", kind="stable").reset_index(drop=True)
    kept["state"] = kept.ori.str[:2]
    kept["agency"] = kept.ori.str[2:5]
    kept["marker"] = [AGENCIES.get(code, (OTHER_SHAPE,))[0] for code in kept.agency]
    return kept, notice


def assign_colors(arrests, style):
    """Give a color to every out-of-state ORI and to the Washington ORIs used most."""
    counts = arrests.ori.value_counts()
    qualifies = sorted(o for o, n in counts.items() if n >= NAMED_MINIMUM or not o.startswith("WA"))
    colors = style["ori_colors"]
    unused = [color for color in SPARE_COLORS if color not in colors.values()]

    added = []
    for ori in qualifies:
        if ori not in colors and unused:
            colors[ori] = unused.pop(0)
            added.append(ori)
    if added:
        print(f"Gave {', '.join(added)} a color, saved in {STYLE_FILE.name}.\n")

    named = [ori for ori in qualifies if ori in colors]
    if len(named) < len(qualifies):
        print("Out of spare colors; some ORIs are gray. Add more to SPARE_COLORS.\n")
    arrests["named"] = arrests.ori.isin(named)
    arrests["color"] = [colors[o] if n else GRAY_DOT for o, n in zip(arrests.ori, arrests.named)]
    return named, bool(added)


def build_key(arrests, style, named):
    """Lay the key out in groups of rows, each row appearing when its ORI is first used."""
    colors, states = style["ori_colors"], style["state_names"]
    first_used = arrests.groupby("ori").date.min()

    def key_row(oris, label, color, marker, is_named):
        return {"label": label, "color": color, "marker": marker, "oris": set(oris),
                "named": is_named, "first": min(first_used[o] for o in oris)}

    groups = []
    washington = arrests[arrests.state == "WA"]
    agency_order = list(AGENCIES)
    for agency in sorted(washington.agency.unique(),
                         key=lambda code: (code not in agency_order, agency_order.index(code)
                                           if code in agency_order else code)):
        shape, header = AGENCIES.get(agency, (OTHER_SHAPE, f"{agency} ORIs"))
        oris = sorted(washington.ori[washington.agency == agency].unique())
        rows = [key_row([o], o, colors[o], shape, True) for o in oris if o in named]
        unnamed = [o for o in oris if o not in named]
        if unnamed:
            rows.append(key_row(unnamed, "Other ORIs", GRAY_DOT, shape, False))
        groups.append({"header": header, "marker": shape, "rows": rows})

    out_of_state = arrests[arrests.state != "WA"]
    if len(out_of_state):  # these are few, so each one gets its own row whatever its count
        rows = [key_row([o], f"{o} ({states.get(o[:2], o[:2])})",
                        colors[o] if o in named else GRAY_DOT,
                        AGENCIES.get(o[2:5], (OTHER_SHAPE,))[0], o in named)
                for o in sorted(out_of_state.ori.unique())]
        groups.append({"header": "Out-of-state ORIs", "marker": None, "rows": rows})

    for group in groups:
        group["rows"].sort(key=lambda row: (not row["named"], row["first"], row["label"]))
        group["first"] = min(row["first"] for row in group["rows"])
    return groups[:1] + sorted(groups[1:], key=lambda group: group["first"])


def spread_stacked_points(arrests):
    """Stagger arrests sharing one set of coordinates."""
    meters = 150
    degrees_lat = meters / 111_320
    degrees_lon = degrees_lat / np.cos(np.radians(np.mean(EXTENT[2:])))
    stacked = arrests.duplicated(["lat", "lon"], keep="first")
    rng = np.random.default_rng(7)  # fixed seed
    arrests.loc[stacked, "lat"] += rng.normal(0, degrees_lat, stacked.sum())
    arrests.loc[stacked, "lon"] += rng.normal(0, degrees_lon, stacked.sum())


def draw_map(fig, counties, width):
    west, east, south, north = EXTENT
    aspect = 1 / np.cos(np.radians(np.mean((south, north))))
    ax = fig.add_axes([MARGIN, MAP_BOTTOM, width, MAP_HEIGHT])
    ax.set_facecolor(BACKGROUND)
    counties.plot(ax=ax, color=LAND, edgecolor=BORDER, linewidth=HAIRLINE)
    ax.set_xlim(west, east)
    ax.set_ylim(south, north)
    ax.set_aspect(aspect)
    ax.axis("off")

    halo = [patheffects.withStroke(linewidth=3, foreground=BACKGROUND)]
    dx, dy = CITY_LABEL_OFFSET
    for name, (lon, lat) in CITIES.items():
        ax.text(lon + dx, lat + dy, name, color=TEXT_DIM, fontsize=SIZE_LABEL,
                zorder=7, path_effects=halo)
    return ax


def draw_key(fig, groups, left):
    rows_left = 1 + len(groups[0]["rows"])
    rows_right = sum(1 + len(group["rows"]) for group in groups[1:])
    crowding = max(rows_left, rows_right + 0.5 * max(len(groups) - 2, 0))
    row_height = min(KEY_ROW_HEIGHT, (KEY_TOP - KEY_BOTTOM) / crowding)
    scale = row_height / KEY_ROW_HEIGHT  # shrink the text if an update brings more ORIs

    columns = [left, left + KEY_COLUMN_X]
    tops = [KEY_TOP, KEY_TOP]
    headers = []
    for position, group in enumerate(groups):
        column = 0 if position == 0 else 1
        x, y = columns[column], tops[column]
        parts = []
        if group["marker"]:
            parts.append(fig.add_artist(Line2D([x + KEY_MARKER_X], [y], marker=group["marker"],
                                               color=TEXT_FAINT, markersize=9,
                                               transform=fig.transFigure)))
        parts.append(fig.text(x + (KEY_LABEL_X if group["marker"] else 0), y, group["header"],
                              color=TEXT_FAINT, fontsize=SIZE_LABEL * scale, va="center"))
        headers.append({"parts": parts, "first": group["first"]})

        y -= row_height
        group["slots"] = [y - n * row_height for n in range(len(group["rows"]))]
        for row in group["rows"]:
            size = KEY_MARKER_SIZE[row["named"]]
            row["dot"] = fig.add_artist(Line2D([x + KEY_MARKER_X], [y], marker=row["marker"],
                                               color=row["color"], markersize=size,
                                               transform=fig.transFigure, markeredgecolor=BACKGROUND,
                                               markeredgewidth=HAIRLINE))
            ink = TEXT if row["named"] else TEXT_DIM
            row["label_text"] = fig.text(x + KEY_LABEL_X, y, row["label"], color=ink,
                                         fontsize=SIZE_KEY * scale, va="center")
            row["count_text"] = fig.text(x + KEY_COUNT_X, y, "", color=ink, fontsize=SIZE_KEY * scale,
                                         va="center", ha="right", weight="bold")
            row["size"] = size
            y -= row_height
        tops[column] = y - KEY_GROUP_GAP
    return headers


def draw_chart(fig, arrests, left, start, end):
    """Weekly bar chart under the key."""
    ax = fig.add_axes([left, CHART_BOTTOM, PANEL_WIDTH, CHART_HEIGHT])
    ax.set_facecolor(BACKGROUND)
    weeks = pd.date_range(start - pd.Timedelta(days=start.weekday()), end, freq="W-MON")
    tallest = arrests.date.dt.to_period("W-SUN").dt.start_time.value_counts().max()

    bars = ax.bar(weeks, np.zeros(len(weeks)), width=BAR_DAYS, color=TEXT_DIM, align="edge")
    ax.set_xlim(start - pd.Timedelta(days=1), weeks[-1] + pd.Timedelta(days=7))
    ax.set_ylim(0, tallest * CHART_HEADROOM)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(BORDER)
    ax.set_yticks([])
    ax.tick_params(length=0)

    months = pd.date_range(start.to_period("M").to_timestamp(), end, freq="MS")
    months = months[months >= start - pd.Timedelta(days=7)]
    across_years = start.year != end.year
    ax.set_xticks(months)
    ax.set_xticklabels([f"{m:%b}" + (f" {m.year}" if across_years and m.month == 1 else "")
                        for m in months], color=TEXT_FAINT, fontsize=SIZE_SMALL)
    ax.text(0, 1.03, "DOL-assisted arrests per week", transform=ax.transAxes,
            color=TEXT_DIM, fontsize=SIZE_LABEL, va="bottom")
    cursor = ax.axvline(start, color=TEXT_FAINT, lw=1.0, zorder=0)
    return bars, weeks, cursor


def render(arrests, counties, groups, out_path):
    """Draw every frame and write the video; plus a PNG of the last frame."""
    plt.rcParams["font.family"] = FONTS
    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
    dpi = DPI // 2 if PREVIEW else DPI

    arrests = arrests.copy()
    spread_stacked_points(arrests)
    west, east, south, north = EXTENT
    arrests["inside"] = arrests.lon.between(west, east) & arrests.lat.between(south, north)

    start, end = arrests.date.min(), arrests.date.max()
    days = (end - start).days + 1
    frames_per_day = max(1, round(MOTION_SECONDS * FPS / days))
    moving_frames = days * frames_per_day
    total_frames = moving_frames + HOLD_SECONDS * FPS
    arrests["appears"] = ((arrests.date - start).dt.days * frames_per_day).astype(int)

    width, height = FIGURE_INCHES
    fig = plt.figure(figsize=FIGURE_INCHES, dpi=dpi, facecolor=BACKGROUND)
    aspect = 1 / np.cos(np.radians(np.mean((south, north))))
    map_width = MAP_HEIGHT * height / width * (east - west) / ((north - south) * aspect)
    ax = draw_map(fig, counties, map_width)

    dots = {(marker, named): ax.scatter([], [], s=[], marker=marker, linewidths=0.6,
                                        zorder=5 if named else 3)
            for marker in sorted(set(arrests.marker)) for named in (False, True)}
    rings = ax.scatter([], [], s=[], facecolors="none", linewidths=1.6, zorder=6)

    left = MARGIN + map_width + PANEL_GAP
    fig.text(left, TITLE_Y, TITLE, color=TEXT, fontsize=SIZE_TITLE, weight="bold",
             va="top", linespacing=1.15)
    date_text = fig.text(left, DATE_Y, "", color=TEXT, fontsize=SIZE_DATE, weight="bold", va="top")
    caption_text = fig.text(left, CAPTION_Y, "", color=TEXT_DIM, fontsize=SIZE_BODY,
                            va="top", linespacing=1.3)
    headers = draw_key(fig, groups, left)
    bars, weeks, cursor = draw_chart(fig, arrests, left, start, end)
    elsewhere_text = fig.text(left, ELSEWHERE_Y, "", color=TEXT_FAINT, fontsize=SIZE_LABEL)
    covers = f"from {format_date(start, year=start.year != end.year)} to {format_date(end)}"
    fig.text(left, SOURCE_Y, f"Data: {SOURCE} (confirmed cases only, {covers}).",
             color=TEXT_FAINT, fontsize=SIZE_SMALL)

    for group in groups:
        for row in group["rows"]:
            row["appears"] = int((row["first"] - start).days * frames_per_day)
    for header in headers:
        header["appears"] = int((header["first"] - start).days * frames_per_day)

    import imageio_ffmpeg

    plt.rcParams["animation.ffmpeg_path"] = imageio_ffmpeg.get_ffmpeg_exe()
    writer = FFMpegWriter(fps=FPS, codec="libx264",
                          extra_args=["-pix_fmt", "yuv420p", "-preset", "veryfast", "-crf", "20"])
    print(f"Rendering {total_frames} frames...", flush=True)
    with writer.saving(fig, str(out_path), dpi=dpi):
        for frame in range(total_frames):
            moment = min(frame, moving_frames - 1)
            today = start + pd.Timedelta(days=moment // frames_per_day)
            so_far = arrests[arrests.appears <= moment]
            fresh = (np.clip(1 - (moment - so_far.appears) / FADE_FRAMES, 0, 1).values
                     if frame < moving_frames else np.zeros(len(so_far)))
            on_map, on_map_fresh = so_far[so_far.inside], fresh[so_far.inside.values]

            for (marker, named), layer in dots.items():
                shown = ((on_map.marker == marker) & (on_map.named == named)).values
                points = on_map[shown]
                layer.set_offsets(np.c_[points.lon, points.lat] if len(points) else np.empty((0, 2)))
                layer.set_sizes(DOT_AREA[named] * (1 + POP_GROWTH * on_map_fresh[shown] ** 2))
                face = to_rgba_array(list(points.color)) if len(points) else np.empty((0, 4))
                if len(points):
                    face[:, 3] = DOT_ALPHA[named]
                layer.set_facecolors(face)
                layer.set_edgecolors([BACKGROUND] * len(points))

            popping = on_map_fresh > 0
            points = on_map[popping]
            rings.set_offsets(np.c_[points.lon, points.lat] if len(points) else np.empty((0, 2)))
            rings.set_sizes(DOT_AREA[True] * (2 + RING_GROWTH * (1 - on_map_fresh[popping])))
            edge = to_rgba_array(list(points.color)) if len(points) else np.empty((0, 4))
            if len(points):
                edge[:, 3] = on_map_fresh[popping] * 0.9
            rings.set_edgecolors(edge)

            cases = f"{len(so_far)} confirmed case" + ("s" if len(so_far) != 1 else "")
            sentence = (f"In {cases} so far, immigration arrests have followed queries of WA~DOL "
                        "data by the following ORIs, or agency usernames:")
            date_text.set_text(format_date(today))
            caption_text.set_text(textwrap.fill(sentence, CAPTION_WIDTH).replace("~", " "))

            counts = so_far.ori.value_counts()
            for group in groups:
                visible = 0
                for row in group["rows"]:
                    if moment >= row["appears"]:
                        y = group["slots"][visible]
                        visible += 1
                        row["dot"].set_ydata([y])
                        row["label_text"].set_y(y)
                        row["count_text"].set_y(y)
                for row in group["rows"]:
                    settled = np.clip((moment - row["appears"] + 1) / FADE_FRAMES, 0, 1)
                    row["dot"].set_alpha(settled * DOT_ALPHA[row["named"]])
                    row["dot"].set_markersize(row["size"] * (1 + 0.8 * (1 - settled)))
                    row["label_text"].set_alpha(settled)
                    row["count_text"].set_alpha(settled)
                    total = counts[counts.index.isin(row["oris"])].sum()
                    row["count_text"].set_text(str(int(total)) if settled > 0 else "")
            for header in headers:
                settled = np.clip((moment - header["appears"] + 1) / FADE_FRAMES, 0, 1)
                for part in header["parts"]:
                    part.set_alpha(settled)

            elsewhere = int((~so_far.inside).sum())
            elsewhere_text.set_text(
                f"+ {elsewhere} arrest" + ("s" if elsewhere != 1 else "") + " elsewhere in Washington"
                if elsewhere else "")
            weekly = so_far.date.dt.to_period("W-SUN").dt.start_time.value_counts()
            for bar, week in zip(bars, weeks):
                bar.set_height(weekly.get(week, 0))
            cursor.set_xdata([today, today])

            writer.grab_frame(facecolor=BACKGROUND)
            if frame % (FPS * 5) == 0:
                print(f"  frame {frame}/{total_frames}", flush=True)

    still = out_path.with_name(out_path.stem + "_final_frame.png")
    fig.savefig(still, dpi=dpi, facecolor=BACKGROUND)
    return still


def choose_spreadsheet():
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        root.update()
        path = filedialog.askopenfilename(
            title="Choose the arrests spreadsheet",
            filetypes=[("Spreadsheets", ("*.csv", "*.xlsx", "*.xls")), ("All files", "*.*")])
        root.destroy()
    except Exception:
        path = input("Paste the spreadsheet's path here and press Enter: ").strip().strip("\"'")
    if not path:
        sys.exit("No spreadsheet chosen.")
    return Path(path)


def main():
    print("Choose the spreadsheet in the window that opens.")
    spreadsheet = choose_spreadsheet()

    style = load_style()
    arrests, notice = load_arrests(spreadsheet)
    if notice:
        print("\n" + "\n".join(notice) + "\n")
    named, colors_added = assign_colors(arrests, style)
    if colors_added:
        save_style(style)
    groups = build_key(arrests, style, named)

    counties = gpd.read_file(Path(__file__).with_name("wa_counties.gpkg"))
    out_path = spreadsheet.with_name("dol_arrests_map" + ("_preview" if PREVIEW else "") + ".mp4")
    still = render(arrests, counties, groups, out_path)
    print(f"\nSaved {out_path.name} and {still.name} in {out_path.parent}")


if __name__ == "__main__":
    main()
