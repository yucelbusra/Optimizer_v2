# -*- coding: utf-8 -*-

from __future__ import print_function
import os
import csv
import json
import math
from datetime import datetime
import io

# Provide Py2/Py3 compatibility for type checks when run outside IronPython
try:
    basestring
except NameError:
    basestring = (str,)

# ------------------ ANSI COLOR HELPERS ------------------
class Ansi(object):
    RESET = "\033[0m"
    BOLD = "\033[1m"
    GREEN = "\033[92m"
    CYAN = "\033[96m"
    YELLOW = "\033[93m"
    MAGENTA = "\033[95m"
    RED = "\033[91m"

# Global active configuration (used for orientation & constraints)
ACTIVE_CONFIG = None

# =============================================================================
# SECTION 1: DATA STRUCTURES & BASIC CONFIGURATION
# =============================================================================
class OpeningClearances(object):
    """Minimum clearances around openings (inches)."""
    def __init__(self, jamb_min=6, header_min=8, sill_min=6):
        self.jamb_min = float(jamb_min)
        self.header_min = float(header_min)
        self.sill_min = float(sill_min)

class Panel(object):
    def __init__(self, x=0, y=0, w=0, h=0, name="", cutouts=None):
        self.x = float(x)
        self.y = float(y)
        self.w = float(w)
        self.h = float(h)
        self.name = name or ""
        self.cutouts = cutouts or []

class Opening(object):
    """Opening with clearance zones."""
    def __init__(self, oid, otype, x, y, w, h, clearances_template):
        self.id = str(oid)
        self.type = str(otype)
        self.x = float(x)      # Left edge in inches
        self.y = float(y)      # Bottom edge (sill) in inches
        self.w = float(w)      # Width in inches
        self.h = float(h)      # Height in inches
        
        # [CRITICAL FIX] Create independent copies of clearances.
        self.original_clearances = OpeningClearances(
            clearances_template.jamb_min,
            clearances_template.header_min,
            clearances_template.sill_min
        )
        self.clearances = OpeningClearances(
            clearances_template.jamb_min,
            clearances_template.header_min,
            clearances_template.sill_min
        )
        
        self.force_blocker = False

    @property
    def left_clearance_zone(self):
        # Never allow clearance zones to extend past wall start
        return max(0.0, self.x - self.clearances.jamb_min)


    @property
    def right_clearance_zone(self):
        """X coordinate where jamb clearance ends (right side)."""
        return self.x + self.w + self.clearances.jamb_min

    @property
    def top_clearance_zone(self):
        """Y coordinate where header clearance ends (top side)."""
        return self.y + self.h + self.clearances.header_min

    @property
    def bottom_clearance_zone(self):
        # Never allow clearance zones to go below wall base
        return max(0.0, self.y - self.clearances.sill_min)


# ------------------ PANEL DIMENSION CONFIGURATION (LEGACY CONSTANTS) ------------------
PANEL_WIDTH_MIN = 24
PANEL_HEIGHT_MIN = 24
LONG_MAX = 348
SHORT_MAX = 138
DIMENSION_INCREMENT = 1


def snap_down(value, inc):
    try:
        value = float(value)
        inc = float(inc)
        return (value // inc) * inc
    except Exception:
        return 0

def snap_up(value, inc):
    try:
        value = float(value)
        inc = float(inc)
        return ((value + inc - 1) // inc) * inc
    except Exception:
        return 0


# =============================================================================
# CONFIGURATION SYSTEM
# =============================================================================
class PanelConstraints(object):
    def __init__(self, min_width=24.0, max_width=138, min_height=24.0, max_height=348.0,
                 short_max=138, long_max=348.0, dimension_increment=1.0, panel_spacing=0.125):
        self.min_width = float(min_width)
        self.max_width = float(max_width)
        self.min_height = float(min_height)
        self.max_height = float(max_height)
        self.short_max = float(short_max)
        self.long_max = float(long_max)
        self.dimension_increment = float(dimension_increment)
        self.panel_spacing = float(panel_spacing)


class OptimizationStrategy(object):
    def __init__(self, prioritize_coverage=True, allow_vertical_stacking=True,
                 prefer_full_height_panels=True, fill_above_storefronts=True,
                 panel_orientation="vertical"):
        self.prioritize_coverage = bool(prioritize_coverage)
        self.allow_vertical_stacking = bool(allow_vertical_stacking)
        self.prefer_full_height_panels = bool(prefer_full_height_panels)
        self.fill_above_storefronts = bool(fill_above_storefronts)
        self.panel_orientation = str(panel_orientation)

class OptimizerConfig(object):
    def __init__(self, project_name="Default Project", panel_constraints=None,
                 door_clearances=None, window_clearances=None,
                 storefront_clearances=None, optimization_strategy=None):
        self.project_name = project_name
        self.panel_constraints = panel_constraints or PanelConstraints()
        self.door_clearances = door_clearances or OpeningClearances()
        self.window_clearances = window_clearances or OpeningClearances()
        self.storefront_clearances = storefront_clearances or OpeningClearances()
        self.optimization_strategy = optimization_strategy or OptimizationStrategy()

    def to_dict(self):
        return {
            "project_name": self.project_name,
            # Simplified for brevity
        }

    @classmethod
    def from_dict(cls, data):
        pc = data.get("panel_constraints", {})
        dc = data.get("door_clearances", {})
        wc = data.get("window_clearances", {})
        sc = data.get("storefront_clearances", {})
        os_ = data.get("optimization_strategy", {})
        return cls(
            project_name=data.get("project_name", "Default Project"),
            panel_constraints=PanelConstraints(
                pc.get("min_width", 24.0), pc.get("max_width", 348.0),
                pc.get("min_height", 24.0), pc.get("max_height", 144.0),
                pc.get("short_max", 138), pc.get("long_max", 348.0),
                pc.get("dimension_increment", 1.0),
                pc.get("panel_spacing", 0.125)
            ),
            door_clearances=OpeningClearances(dc.get("jamb_min", 6), dc.get("header_min", 8), dc.get("sill_min", 6)),
            window_clearances=OpeningClearances(wc.get("jamb_min", 6), wc.get("header_min", 8), wc.get("sill_min", 6)),
            storefront_clearances=OpeningClearances(sc.get("jamb_min", 0.75), sc.get("header_min", 0.75), sc.get("sill_min", 0.75)),
            optimization_strategy=OptimizationStrategy(
                os_.get("prioritize_coverage", True),
                os_.get("allow_vertical_stacking", True),
                os_.get("prefer_full_height_panels", True),
                os_.get("fill_above_storefronts", True),
                os_.get("panel_orientation", "vertical")
            )
        )


    def save(self, filepath):
        try:
            f = io.open(filepath, "w", newline="")
        except TypeError:
            f = open(filepath, "w")

        with f:
            json.dump(self.to_dict(), f, indent=2)

        print("{} Config saved: {}{}".format(Ansi.GREEN, filepath, Ansi.RESET))

    @classmethod
    def load(cls, filepath):
        with open(filepath, "r") as f:
            data = json.load(f)
        return cls.from_dict(data)


def get_preset_configs():
    presets = {}
    presets["vertical"] = OptimizerConfig(
        project_name="Vertical Panels",
        panel_constraints=PanelConstraints(
            min_width=24, max_width=138, min_height=24, max_height=348.0,
            short_max=138, long_max=348.0, dimension_increment=1, panel_spacing=0.125
        ),
        door_clearances=OpeningClearances(6, 8, 6),
        window_clearances=OpeningClearances(4, 6, 4),
        storefront_clearances=OpeningClearances(0.75, 0.75, 0.75),
        optimization_strategy=OptimizationStrategy(True, True, True, True, "vertical")
    )
    presets["horizontal"] = OptimizerConfig(
        project_name="Horizontal Panels",
        panel_constraints=PanelConstraints(
            min_width=12, max_width=348.0, min_height=12, max_height=138,
            short_max=138, long_max=348.0, dimension_increment=1, panel_spacing=0.125
        ),
        door_clearances=OpeningClearances(6, 8, 6),
        window_clearances=OpeningClearances(6, 8, 6),
        storefront_clearances=OpeningClearances(0.75, 0.75, 0.75),
        optimization_strategy=OptimizationStrategy(False, True, False, True, "horizontal")
    )
    return presets


# =============================================================================
# SECTION 2: DATA LOADING & VALIDATION (CSV-based)
# =============================================================================

def read_csv_rows(path):
    if not os.path.exists(path): return []
    rows = []
    with open(path, "r") as f:
        reader = csv.DictReader(f)
        for row in reader: rows.append(row)
    return rows


def load_walls_from_csv(walls_csv):
    if not os.path.exists(walls_csv):
        raise IOError("Walls CSV not found: {}".format(walls_csv))
    rows = read_csv_rows(walls_csv)
    print(Ansi.CYAN + "[INFO] Loaded {} walls from CSV".format(len(rows)) + Ansi.RESET)
    return rows


def load_openings_from_csv(openings_csv):
    if not os.path.exists(openings_csv):
        print(Ansi.YELLOW + "[WARN] Openings CSV not found." + Ansi.RESET)
        return []
    rows = read_csv_rows(openings_csv)
    norm_rows = []
    for r in rows:
        nr = {}
        for k, v in r.items():
            nk = k.strip() if isinstance(k, basestring) else k
            nr[nk] = v
        norm_rows.append(nr)
    print(Ansi.CYAN + "[INFO] Loaded {} openings from CSV".format(len(norm_rows)) + Ansi.RESET)
    return norm_rows


def _is_empty(v):
    if v is None: return True
    if isinstance(v, basestring):
        s = v.strip()
        return s == "" or s.lower() == "nan" or s.lower() == "none"
    try:
        return math.isnan(float(v))
    except Exception:
        return False


def safe_float(v, default=0.0):
    try:
        if _is_empty(v): return default
        return float(v)
    except Exception: return default


def get_wall_id(wall_row):
    for col in ["WallId", "ElementId", "Id"]:
        if col in wall_row and not _is_empty(wall_row.get(col)):
            val = wall_row.get(col)
            try: return str(int(float(val)))
            except: return str(val)
    if "Name" in wall_row and not _is_empty(wall_row.get("Name")):
        return str(wall_row.get("Name"))
    return "unknown"


def get_wall_dimensions(wall_row):
    try:
        length_ft = safe_float(wall_row.get("Length(ft)", 0))
        height_ft = safe_float(wall_row.get("UnconnectedHeight(ft)", 0))
        if length_ft <= 0 or height_ft <= 0: return None
        return (float((length_ft * 12)), float((height_ft * 12)))
    except Exception: return None


def get_wall_openings(wall_id, openings_rows, door_clearances, window_clearances, storefront_clearances):
    if not openings_rows: return []
    try:
        wall_id_int = int(float(wall_id))
    except Exception: return []
    
    wall_openings = [r for r in openings_rows if safe_float(r.get("HostWallId"), None) == wall_id_int]
    if not wall_openings: return []

    openings = []
    for row in wall_openings:
        width_ft = safe_float(row.get("Width(ft)", 0))
        height_ft = safe_float(row.get("Height(ft)", 0))
        sill_ft = safe_float(row.get("SillHeight(ft)", 0))
        
        if width_ft <= 0 or height_ft <= 0: continue
        
        left_ft = safe_float(row.get("LeftEdgeAlongWall(ft)", 0))
        if left_ft == 0 and "PositionAlongWall(ft)" in row:
             pos = safe_float(row.get("PositionAlongWall(ft)", 0))
             if pos != 0: left_ft = pos - (width_ft/2.0)

        x_in = float(left_ft * 12)
        y_in = float(sill_ft * 12)
        w_in = float(width_ft * 12)
        h_in =float(height_ft * 12)
        
        opening_type = str(row.get("OpeningType", "Unknown"))
        otype_lower = opening_type.lower()

        if "door" in otype_lower:
            clearances = door_clearances
        elif ("storefront" in otype_lower) or ("curtain" in otype_lower):
            clearances = storefront_clearances
            opening_type = "Storefront/Curtain"
        else:
            clearances = window_clearances

        openings.append(Opening(row.get("OpeningId", ""), opening_type, x_in, y_in, w_in, h_in, clearances))
    
    return openings


def adjust_panels_for_small_openings(panels, openings, constraints, dim_inc):
    SMALL_W = 72.0 
    SMALL_H = 120.0
    spacing = constraints.panel_spacing

    for opening in openings:
        if not (opening.w < SMALL_W and opening.h < SMALL_H): continue

        opening_right = opening.x + opening.w
        band_panels = [
            p for p in panels
            if not (p.y + p.h <= opening.y or p.y >= opening.y + opening.h)
        ]
        band_panels.sort(key=lambda p: p.x)

        for i in range(len(band_panels) - 1):
            left = band_panels[i]
            right = band_panels[i + 1]
            actual_seam = left.x + left.w + spacing

            if abs(right.x - actual_seam) > 1.0: continue
            if not (opening.x < actual_seam < opening_right): continue

            new_seam = snap_down(opening.x, dim_inc)
            if new_seam <= left.x: continue

            delta = actual_seam - new_seam
            new_left_fab_w = left.w - delta
            new_right_x = new_seam
            new_right_fab_w = (right.x + right.w) - new_right_x

            if new_left_fab_w < constraints.min_width or new_right_fab_w < constraints.min_width:
                continue

            left.w = float(new_left_fab_w)
            right.x = float(new_right_x)
            right.w = float(new_right_fab_w)
            break


# =============================================================================
# SECTION 3: COLLISION DETECTION
# =============================================================================

def panels_overlap(p1, p2):
    return not (p1.x + p1.w <= p2.x or p2.x + p2.w <= p1.x or
                p1.y + p1.h <= p2.y or p2.y + p2.h <= p1.y)

def is_storefront_like(opening):
    otype = (opening.type or "").lower()
    return ("storefront" in otype) or ("curtain" in otype)

def classify_openings_dynamic(openings, constraints):
    """
    Decides if an opening is a CUTOUT (bridged) or BLOCKER (stop).
    [FIXED] Storefronts are ALWAYS blockers now.
    """
    max_panel_w = constraints.max_width
    spacing = constraints.panel_spacing

    for op in openings:
        otype = op.type.lower()
        is_storefront = "storefront" in otype or "curtain" in otype
        
        # [FIX] If it's a Storefront, it is ALWAYS a blocker.
        if is_storefront:
            op.force_blocker = True
            op.clearances.jamb_min = spacing
            op.clearances.header_min = spacing
            op.clearances.sill_min = spacing
            print("    [BLOCKER] Storefront {} forced to block. Gap set to {}.".format(op.id, spacing))
            continue

        # For normal windows/doors, check size
        required_span = op.w + op.original_clearances.jamb_min * 2
        
        if required_span <= max_panel_w:
            # Fits -> Cutout
            op.force_blocker = False
            op.clearances.jamb_min = op.original_clearances.jamb_min
            op.clearances.header_min = op.original_clearances.header_min
            op.clearances.sill_min = op.original_clearances.sill_min
        else:
            # Too wide -> Blocker
            op.force_blocker = True
            op.clearances.jamb_min = spacing
            op.clearances.header_min = spacing
            op.clearances.sill_min = spacing
            print("    [BLOCKER] Opening {} (Width={:.1f}\") > Max Panel. Gap set to {}.".format(op.id, op.w, spacing))

def is_blocking_storefront(opening, constraints):
    return opening.force_blocker

def is_cutout_opening(opening, constraints):
    return not opening.force_blocker

def panel_overlaps_clearance(panel, openings, constraints, allow_intentional=False):
    p_right = panel.x + panel.w
    p_top = panel.y + panel.h

    for opening in openings:
        if allow_intentional: continue
        if is_cutout_opening(opening, constraints): continue

        if not (
            p_right <= opening.left_clearance_zone or
            panel.x >= opening.right_clearance_zone or
            p_top <= opening.bottom_clearance_zone or
            panel.y >= opening.top_clearance_zone
        ):
            return True
    return False


def fill_vertical_gap(region_x_start, region_x_end, gap_y_start, gap_y_end,
                      opening_left, opening_right, panels, panel_counter,
                      constraints, all_openings, label,
                      is_storefront=False):
    PANEL_WIDTH_MIN = constraints.min_width
    PANEL_HEIGHT_MIN = constraints.min_height
    SHORT_MAX = constraints.short_max
    LONG_MAX = constraints.long_max
    DIMENSION_INCREMENT = constraints.dimension_increment
    spacing = float(getattr(constraints, "panel_spacing", 0.0) or 0.0)

    if is_storefront:
        panel_x_start = region_x_start + spacing
        panel_x_end   = region_x_end - spacing
    else:
        panel_x_start = max(opening_left + spacing, region_x_start)
        panel_x_end   = min(opening_right - spacing, region_x_end)

    gap_width  = panel_x_end - panel_x_start
    gap_height = gap_y_end - gap_y_start

    if gap_width < PANEL_WIDTH_MIN or gap_height < PANEL_HEIGHT_MIN:
        return panel_counter

    y_cursor = gap_y_start
    while y_cursor < gap_y_end:
        remaining_height = gap_y_end - y_cursor
        if remaining_height < PANEL_HEIGHT_MIN: break

        panel_h = snap_down(remaining_height, DIMENSION_INCREMENT)
        if panel_h < PANEL_HEIGHT_MIN: break

        max_width = SHORT_MAX if panel_h > SHORT_MAX else LONG_MAX
        x_cursor = panel_x_start
        row_placed = False

        while x_cursor < panel_x_end:
            remaining_width = panel_x_end - x_cursor
            if remaining_width < PANEL_WIDTH_MIN: break

            panel_w = min(remaining_width, max_width)
            panel_w = snap_down(panel_w, DIMENSION_INCREMENT)

            leftover = remaining_width - panel_w
            if leftover > 0 and leftover < (PANEL_WIDTH_MIN + spacing):
                panel_w = snap_down(remaining_width, DIMENSION_INCREMENT)

            if panel_w < PANEL_WIDTH_MIN or not is_valid_panel(panel_w, panel_h, constraints): break

            candidate = Panel(x_cursor, y_cursor, panel_w, panel_h, "P{:02d}".format(panel_counter))

            if any(panels_overlap(candidate, p) for p in panels): break
            if panel_overlaps_clearance(candidate, all_openings, constraints, allow_intentional=True): break

            candidate.cutouts = calculate_panel_cutouts(candidate, all_openings)
            panels.append(candidate)
            panel_counter += 1
            row_placed = True
            x_cursor += (panel_w + spacing)

        if row_placed:
            y_cursor += (panel_h + spacing)
        else:
            break

    return panel_counter


def calculate_segment_layout(start_x, target_x, max_w, min_w, inc, spacing):
    total_dist = target_x - start_x
    if total_dist < min_w: return total_dist
    if total_dist <= max_w: return snap_down(total_dist, inc)

    n_panels = 1
    while True:
        total_spacing = (n_panels - 1) * spacing
        available_width = total_dist - total_spacing
        candidate_width = available_width / n_panels

        if candidate_width <= max_w:
            if candidate_width < min_w: return max_w
            return snap_down(candidate_width, inc)
        n_panels += 1
        if n_panels > 100: return max_w


def place_panels_sequential(wall_width, wall_height, openings, constraints, orientation="vertical"):
    """
    Fixed panel placement with Lookahead Logic + Seam Validation.
    """
    orientation = str(orientation or "vertical").lower()
    horizontal_mode = (orientation == "horizontal")

    # 1. Run Dynamic Classification
    classify_openings_dynamic(openings, constraints)

    # Bind constraints
    PANEL_WIDTH_MIN = constraints.min_width
    PANEL_HEIGHT_MIN = constraints.min_height
    SHORT_MAX = constraints.short_max
    LONG_MAX = constraints.long_max
    DIMENSION_INCREMENT = constraints.dimension_increment
    spacing = float(getattr(constraints, "panel_spacing", 0.0) or 0.0)

    panels = []
    panel_counter = 1

    sorted_openings = sorted(openings, key=lambda o: o.x)
    blocking_storefronts = [o for o in sorted_openings if is_blocking_storefront(o, constraints)]
    regular_openings = [o for o in sorted_openings if o not in blocking_storefronts]

    # BUILD X-REGIONS
    regions = []
    if not blocking_storefronts:
        regions.append({
            'x_start': 0, 'x_end': wall_width,
            'y_start': 0, 'y_end': wall_height,
            'openings': regular_openings
        })
    else:
        storefronts_sorted = sorted(blocking_storefronts, key=lambda sf: sf.left_clearance_zone)
        x_boundaries = [0]
        for sf in storefronts_sorted:
            x_boundaries.extend([sf.left_clearance_zone, sf.right_clearance_zone])
        x_boundaries.append(wall_width)
        x_boundaries = sorted(list(set(x_boundaries)))

        for i in range(len(x_boundaries) - 1):
            x_start, x_end = x_boundaries[i], x_boundaries[i + 1]
            if (x_end - x_start) < PANEL_WIDTH_MIN: continue

            blocked = any(
                not (x_end <= sf.left_clearance_zone or x_start >= sf.right_clearance_zone)
                for sf in storefronts_sorted
            )
            if not blocked:
                region_openings_list = [
                    o for o in regular_openings
                    if not (o.right_clearance_zone <= x_start or o.left_clearance_zone >= x_end)
                ]
                regions.append({
                    'x_start': x_start, 'x_end': x_end,
                    'y_start': 0, 'y_end': wall_height,
                    'openings': region_openings_list
                })

    # PROCESS EACH REGION
    for region in regions:
        region_openings = region['openings']
        bands = []
        if horizontal_mode:
            cy = 0
            while cy < wall_height:
                rem_h = wall_height - cy
                bh = snap_down(min(rem_h, SHORT_MAX), DIMENSION_INCREMENT)
                if bh >= PANEL_HEIGHT_MIN:
                    bands.append((cy, cy + bh))
                    cy += bh
                else: break
        else:
            bands = [(region['y_start'], region['y_end'])]

        for y_start, y_end in bands:
            band_height = y_end - y_start
            max_width_for_band = SHORT_MAX if band_height > SHORT_MAX else LONG_MAX
            x_cursor = max(0.0, region['x_start'])

            while x_cursor < region['x_end']:
                remaining_wall = region['x_end'] - x_cursor
                if remaining_wall < PANEL_WIDTH_MIN: break

                future_openings = [
                    o for o in region_openings
                    if (o.left_clearance_zone > x_cursor + 0.01)
                    and not (o.top_clearance_zone <= y_start or o.bottom_clearance_zone >= y_end)
                    and not is_cutout_opening(o, constraints)
                ]
                next_opening = min(future_openings, key=lambda o: o.left_clearance_zone) if future_openings else None

                hard_stop_x = region['x_end']
                target_is_opening = False

                if next_opening:
                    bridge_dist = next_opening.right_clearance_zone - x_cursor
                    can_bridge = (bridge_dist <= max_width_for_band and bridge_dist >= PANEL_WIDTH_MIN)

                    if can_bridge:
                        panel_w = snap_down(bridge_dist, DIMENSION_INCREMENT)
                        candidate = Panel(x_cursor, y_start, panel_w, band_height, "P{:02d}".format(panel_counter))
                        candidate.cutouts = calculate_panel_cutouts(candidate, region_openings)
                        panels.append(candidate)
                        panel_counter += 1
                        x_cursor += (panel_w + spacing)
                        continue
                    else:
                        hard_stop_x = next_opening.left_clearance_zone
                        target_is_opening = True

                dist_to_stop = hard_stop_x - x_cursor
                if dist_to_stop < PANEL_WIDTH_MIN:
                    if target_is_opening: x_cursor = next_opening.right_clearance_zone
                    else: break
                    continue

                panel_w = calculate_segment_layout(x_cursor, hard_stop_x, max_width_for_band, PANEL_WIDTH_MIN, DIMENSION_INCREMENT, spacing)
                candidate_right = x_cursor + panel_w
                for op in region_openings:
                    if (op.left_clearance_zone + 0.1) < candidate_right < (op.right_clearance_zone - 0.1):
                        dist_to_left_jamb = op.left_clearance_zone - x_cursor
                        if dist_to_left_jamb >= PANEL_WIDTH_MIN:
                            panel_w = snap_down(dist_to_left_jamb, DIMENSION_INCREMENT)
                        else:
                            dist_to_right_jamb = op.right_clearance_zone - x_cursor
                            width_to_clear = snap_up(dist_to_right_jamb, DIMENSION_INCREMENT)
                            if width_to_clear <= max_width_for_band: panel_w = width_to_clear
                            else: panel_w = snap_down(max_width_for_band, DIMENSION_INCREMENT)
                        break

                if not is_valid_panel(panel_w, band_height, constraints): break

                candidate = Panel(x_cursor, y_start, panel_w, band_height, "P{:02d}".format(panel_counter))
                if panel_overlaps_clearance(candidate, region_openings, constraints, allow_intentional=False):
                    print("    [WARN] Panel overlaps hard clearance")

                candidate.cutouts = calculate_panel_cutouts(candidate, region_openings)
                panels.append(candidate)
                panel_counter += 1
                x_cursor += (panel_w + spacing)

                if target_is_opening and abs(x_cursor - (hard_stop_x + spacing)) < 1.0:
                    x_cursor = next_opening.right_clearance_zone

        # FILL VERTICAL GAPS
        gap_openings = [o for o in region_openings if not is_cutout_opening(o, constraints)]
        for opening in gap_openings:
            if not is_storefront_like(opening) and opening.bottom_clearance_zone <= 0 and opening.top_clearance_zone >= wall_height:
                continue

            if opening.bottom_clearance_zone > 0:
                gap_height = opening.bottom_clearance_zone - 0
                if gap_height >= PANEL_HEIGHT_MIN:
                    panel_counter = fill_vertical_gap(
                        region['x_start'], region['x_end'],
                        0, opening.bottom_clearance_zone,
                        opening.left_clearance_zone, opening.right_clearance_zone,
                        panels, panel_counter, constraints, sorted_openings,
                        "below"
                    )

            if opening.top_clearance_zone < wall_height:
                gap_height = wall_height - opening.top_clearance_zone
                if gap_height >= PANEL_HEIGHT_MIN:
                    panel_counter = fill_vertical_gap(
                        region['x_start'], region['x_end'],
                        opening.top_clearance_zone, wall_height,
                        opening.left_clearance_zone, opening.right_clearance_zone,
                        panels, panel_counter, constraints, sorted_openings,
                        "above",
                        is_storefront_like(opening)
                    )

        region_panels = [p for p in panels if (region['x_start'] <= p.x < region['x_end'])]
        adjust_panels_for_small_openings(region_panels, region_openings, constraints, DIMENSION_INCREMENT)

    # EXTRA: Fill ABOVE AND BELOW BLOCKING storefront spans
    extra_filled = 0
    for sf in blocking_storefronts:
        # Fill ABOVE
        if sf.top_clearance_zone < wall_height:
            gap_height = wall_height - sf.top_clearance_zone
            if gap_height >= PANEL_HEIGHT_MIN:
                before_count = len(panels)
                panel_counter = fill_vertical_gap(
                    sf.left_clearance_zone, sf.right_clearance_zone,
                    sf.top_clearance_zone, wall_height,
                    sf.left_clearance_zone, sf.right_clearance_zone,
                    panels, panel_counter, constraints, sorted_openings,
                    "above",
                    True
                )
                extra_filled += len(panels) - before_count
        
        # [NEW] Fill BELOW
        if sf.bottom_clearance_zone > 0:
            gap_height = sf.bottom_clearance_zone - 0
            if gap_height >= PANEL_HEIGHT_MIN:
                before_count = len(panels)
                panel_counter = fill_vertical_gap(
                    sf.left_clearance_zone, sf.right_clearance_zone,
                    0, sf.bottom_clearance_zone,
                    sf.left_clearance_zone, sf.right_clearance_zone,
                    panels, panel_counter, constraints, sorted_openings,
                    "below",
                    True
                )
                extra_filled += len(panels) - before_count

    return panels


def calculate_panel_cutouts(panel, openings):
    cutouts = []
    p_left = panel.x
    p_right = panel.x + panel.w
    p_bottom = panel.y
    p_top = panel.y + panel.h

    for opening in openings:
        if opening.force_blocker: continue

        hole_left = opening.left_clearance_zone
        hole_right = opening.right_clearance_zone
        hole_bottom = opening.bottom_clearance_zone
        hole_top = opening.top_clearance_zone

        inter_left = max(p_left, hole_left)
        inter_right = min(p_right, hole_right)
        inter_bottom = max(p_bottom, hole_bottom)
        inter_top = min(p_top, hole_top)

        if inter_right > inter_left and inter_top > inter_bottom:
            cutout_x = inter_left - p_left
            cutout_y = inter_bottom - p_bottom
            cutout_w = inter_right - inter_left
            cutout_h = inter_top - inter_bottom

            cutout_info = {
                "id": opening.id,
                "type": opening.type,
                "x_in": float(cutout_x),
                "y_in": float(cutout_y),
                "width_in": float(cutout_w),
                "height_in": float(cutout_h)
            }
            cutouts.append(cutout_info)

    return cutouts

def process_wall(wall_id, wall_width, wall_height, openings):
    global ACTIVE_CONFIG
    
    if ACTIVE_CONFIG is None:
        presets = get_preset_configs()
        ACTIVE_CONFIG = presets.get("horizontal")

    orientation = str(ACTIVE_CONFIG.optimization_strategy.panel_orientation or "vertical").lower()
    
    panels = place_panels_sequential(
        wall_width, wall_height, openings,
        ACTIVE_CONFIG.panel_constraints, orientation
    )

    records = []
    for panel in panels:
        records.append({
            "panel_name": panel.name,
            "panel_type": "{}x{}".format(panel.w, panel.h),
            "wall_id": wall_id,
            "x_in": panel.x,
            "y_in": panel.y,
            "width_in": panel.w,
            "height_in": panel.h,
            "area_in2": panel.w * panel.h,
            "rotation_deg": 0.0,
            "x_ref": "start",
            "cutouts_json": json.dumps(panel.cutouts) if panel.cutouts else ""
        })
    
    print(Ansi.GREEN + " Result: {} panels generated".format(len(panels)) + Ansi.RESET)
    return records


def process_all_walls(walls_rows, openings_rows, output_dir,
                      door_clearances, window_clearances, storefront_clearances,
                      config=None, orientation="vertical", output_filename="optimized_panel_placement.csv"):
    global ACTIVE_CONFIG
    if config is not None: ACTIVE_CONFIG = config
    elif ACTIVE_CONFIG is None:
        presets = get_preset_configs()
        ACTIVE_CONFIG = presets.get(orientation, presets["vertical"])
    
    all_panel_records = []
    for wall_row in walls_rows:
        wall_id = get_wall_id(wall_row)
        dims = get_wall_dimensions(wall_row)
        if dims is None: continue
        
        wall_width, wall_height = dims
        openings = get_wall_openings(
            wall_id, openings_rows,
            door_clearances, window_clearances, storefront_clearances
        )
        panel_records = process_wall(wall_id, wall_width, wall_height, openings)
        all_panel_records.extend(panel_records)
    
    if not all_panel_records: return None, None
    
    panels_csv = os.path.join(output_dir, output_filename)
    fieldnames = [
        "panel_name", "panel_type", "wall_id",
        "x_in", "y_in", "width_in", "height_in",
        "area_in2", "rotation_deg", "x_ref", "cutouts_json"
    ]
    panels_path = write_csv(panels_csv, all_panel_records, fieldnames)
    
    config_path = None
    if panels_path and ACTIVE_CONFIG:
        config_path = os.path.join(output_dir, "config_used.json")
        try:
            if not os.path.exists(output_dir): os.makedirs(output_dir)
            ACTIVE_CONFIG.save(config_path)
        except: pass
    
    return panels_path, config_path

def write_csv(path, rows, fieldnames=None):
    if not rows: return None
    if fieldnames is None: fieldnames = list(rows[0].keys())
    try: f = open(path, "w", newline="")
    except TypeError: f = open(path, "w")
    with f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows: writer.writerow(r)
    return path

def is_valid_panel(w, h, constraints):
    try: w, h = float(w), float(h)
    except: return False
    if w < constraints.min_width or h < constraints.min_height: return False
    if w > constraints.long_max or h > constraints.long_max: return False
    if w > constraints.short_max and h > constraints.short_max: return False
    return True



def find_next_opening_in_range(x_start, x_end, y_start, y_end, openings):
    """Find the leftmost opening that intersects with the given range."""
    candidates = []
    for o in openings:
        # Check if opening intersects horizontally
        if o.right_clearance_zone <= x_start or o.left_clearance_zone >= x_end:
            continue
        # Check if opening intersects vertically
        if o.top_clearance_zone <= y_start or o.bottom_clearance_zone >= y_end:
            continue
        candidates.append(o)
    
    if not candidates:
        return None
    return min(candidates, key=lambda o: o.left_clearance_zone)


def determine_panel_width_with_opening(x_cursor, x_end, y_start, y_end, max_width, openings, constraints):
    """
    Determine panel width considering openings ahead.
    Returns: (panel_width, should_include_opening, opening_obj or None)
    """
    PANEL_WIDTH_MIN = constraints.min_width
    DIMENSION_INCREMENT = constraints.dimension_increment
    
    # Find next opening in our path
    next_opening = find_next_opening_in_range(x_cursor, x_end, y_start, y_end, openings)
    
    if next_opening is None:
        # No opening ahead - use remaining space or max_width
        available = x_end - x_cursor
        panel_w = min(available, max_width)
        panel_w = snap_down(panel_w, DIMENSION_INCREMENT)
        return (panel_w, False, None)
    
    # Opening exists - decide whether to include or trim before it
    opening_left = next_opening.left_clearance_zone
    opening_right = next_opening.right_clearance_zone
    opening_width = opening_right - opening_left
    
    # Distance to opening start
    dist_to_opening = opening_left - x_cursor
    
    # Can we include the entire opening?
    width_with_opening = opening_right - x_cursor
    
    if dist_to_opening < PANEL_WIDTH_MIN:
        # Opening is very close - we should include it if possible
        if width_with_opening <= max_width:
            # Include the opening
            panel_w = snap_down(width_with_opening, DIMENSION_INCREMENT)
            if panel_w >= PANEL_WIDTH_MIN:
                return (panel_w, True, next_opening)
    
    # Opening is far enough - check if we should trim before it
    if dist_to_opening >= PANEL_WIDTH_MIN and dist_to_opening <= max_width:
        # Trim before the opening
        panel_w = snap_down(dist_to_opening, DIMENSION_INCREMENT)
        if panel_w >= PANEL_WIDTH_MIN:
            return (panel_w, False, next_opening)
    
    # Opening is too far or including it exceeds max_width
    # Use max_width or remaining space
    available = x_end - x_cursor
    panel_w = min(available, max_width)
    panel_w = snap_down(panel_w, DIMENSION_INCREMENT)
    return (panel_w, False, next_opening)



# =============================================================================
# SECTION 6: VISUALIZATION (optional)
# =============================================================================

def visualize_wall_layout(wall_id, panels_csv, openings_csv, walls_csv, output_image=None):
    try:
        import plotly.graph_objects as go
    except Exception:
        print(Ansi.YELLOW + "[VIS] Plotly not installed, skipping" + Ansi.RESET)
        return
    panels_rows = read_csv_rows(panels_csv)
    openings_rows = read_csv_rows(openings_csv)
    walls_rows = read_csv_rows(walls_csv)
    try:
        wall_id_int = int(float(wall_id))
    except Exception:
        wall_id_int = None
    wall_row = None
    for r in walls_rows:
        for col in ["WallId", "ElementId", "Id"]:
            try:
                if col in r and int(float(r.get(col))) == wall_id_int:
                    wall_row = r
                    break
            except Exception:
                pass
        if wall_row:
            break
    if wall_row is None:
        print(Ansi.YELLOW + "[VIS] Wall {} not found".format(wall_id) + Ansi.RESET)
        return
    wall_width = safe_float(wall_row.get("Length(ft)", 0)) * 12.0
    wall_height = safe_float(wall_row.get("UnconnectedHeight(ft)", 0)) * 12.0
    wall_panels = [p for p in panels_rows if str(p.get("wall_id")) == str(wall_id)]
    wall_openings = [o for o in openings_rows if safe_float(o.get("HostWallId"), None) == wall_id_int]
    fig = go.Figure()
    fig.add_shape(type="rect", x0=0, y0=0, x1=wall_width, y1=wall_height,
                  line=dict(color="black", width=3), fillcolor="lightgray", opacity=0.1)
    colors = ['rgba(65,105,225,0.3)', 'rgba(30,144,255,0.3)', 'rgba(100,149,237,0.3)']
    for i, panel in enumerate(wall_panels):
        x_in = float(panel.get("x_in", 0)); y_in = float(panel.get("y_in", 0))
        w_in = float(panel.get("width_in", 0)); h_in = float(panel.get("height_in", 0))
        fig.add_shape(type="rect", x0=x_in, y0=y_in, x1=x_in + w_in, y1=y_in + h_in,
                      line=dict(color="blue", width=2), fillcolor=colors[i % len(colors)])
        fig.add_annotation(x=x_in + w_in/2.0, y=y_in + h_in/2.0,
                           text="<b>{}</b><br/>{}\"x{}\"".format(panel.get('panel_name', ''), w_in, h_in),
                           showarrow=False, font=dict(size=10), bgcolor="white", opacity=0.8)
    for opening in wall_openings:
        left_ft = safe_float(opening.get("LeftEdgeAlongWall(ft)", 0))
        width_ft = safe_float(opening.get("Width(ft)", 0))
        sill_ft = safe_float(opening.get("SillHeight(ft)", 0))
        height_ft = safe_float(opening.get("Height(ft)", 0))
        if width_ft <= 0 or height_ft <= 0:
            continue
        left_in = left_ft * 12.0
        width_in = width_ft * 12.0
        sill_in = sill_ft * 12.0
        height_in = height_ft * 12.0
        opening_type = str(opening.get("OpeningType", "")).lower()
        if "door" in opening_type:
            color = "red"; rgb = "255,0,0"; label = "Door"
        elif ("storefront" in opening_type) or ("curtain" in opening_type):
            color = "darkgreen"; rgb = "0,100,0"; label = "Storefront"
        else:
            color = "purple"; rgb = "128,0,128"; label = "Window"
        fig.add_shape(type="rect",
                      x0=left_in - 6, y0=sill_in - 6,
                      x1=left_in + width_in + 6, y1=sill_in + height_in + 8,
                      line=dict(color="orange", width=1, dash="dash"), fillcolor="rgba(255,165,0,0.1)")
        fig.add_shape(type="rect",
                      x0=left_in, y0=sill_in,
                      x1=left_in + width_in, y1=sill_in + height_in,
                      line=dict(color=color, width=2), fillcolor="rgba({},{})".format(rgb, "0.4"))
        fig.add_annotation(x=left_in + width_in/2.0, y=sill_in + height_in/2.0,
                           text="{}<br/>{}\"x{}\"".format(label, float(width_in), float(height_in)),
                           showarrow=False, font=dict(size=9, color="white"), bgcolor=color, opacity=0.9)
    fig.update_layout(title="Wall {} - Sequential Panel Layout (Doors, Windows & Storefronts)".format(wall_id),
                      xaxis=dict(range=[0, wall_width], title="Length (inches)", showgrid=True),
                      yaxis=dict(range=[0, wall_height], title="Height (inches)", showgrid=True, scaleanchor="x"),
                      width=1400, height=600, showlegend=False, plot_bgcolor='white')
    if output_image:
        try:
            fig.write_image(output_image)
            print(Ansi.CYAN + "[VIS] Saved: {}".format(output_image) + Ansi.RESET)
        except Exception:
            fig.show()
    else:
        fig.show()


def visualize_all_walls(panels_csv, openings_csv, walls_csv, output_dir, save_as_image=True):
    panels_rows = read_csv_rows(panels_csv)
    wall_ids = sorted(set([r.get("wall_id") for r in panels_rows]))
    print(Ansi.MAGENTA + "\n[VIS] Generating {} visualizations...".format(len(wall_ids)) + Ansi.RESET)
    for wid in wall_ids:
        output_image = os.path.join(output_dir, "wall_{}_layout.png".format(wid)) if save_as_image else None
        visualize_wall_layout(wid, panels_csv, openings_csv, walls_csv, output_image)


# =============================================================================
# SECTION 7A: INTERACTIVE CONFIG CREATOR (ORIENTATION-FOCUSED)
# =============================================================================
def create_simple_config():
    """Interactive configuration creator with parameter preview/editing."""
    # IronPython vs CPython input
    try:
        get_input = raw_input  # type: ignore
    except NameError:
        get_input = input

    print("\n" + "=" * 60)
    print("  PANEL OPTIMIZER CONFIGURATION")
    print("=" * 60)

    print("\nSelect configuration preset:")
    print("1. Vertical Panels   (tall, narrow - best for high-rise)")
    print("2. Horizontal Panels (wide, short - best for retail/commercial)")
    print("3. Custom            (define all parameters)")

    choice = get_input("\nChoice (1-3) [default: 1]: ").strip() or "1"

    presets = get_preset_configs()

    if choice == "1":
        config = presets["vertical"]
        print("\n{}=== VERTICAL PANEL PRESET ==={}".format(Ansi.CYAN, Ansi.RESET))
        print_config_summary(config)
        pc = config.panel_constraints
        print("  Spacing:      {}\"".format(pc.panel_spacing))


        # Ask for confirmation first
        confirm = get_input("\nUse these preset values? (y/n/edit) [y]: ").strip().lower()
        if confirm == "n":
            print("{}Cancelled. Returning to menu...{}".format(Ansi.YELLOW, Ansi.RESET))
            return create_simple_config()  # Start over
        elif confirm == "edit" or confirm == "e":
            config = edit_panel_constraints(config)

    elif choice == "2":
        config = presets["horizontal"]
        print("\n{}=== HORIZONTAL PANEL PRESET ==={}".format(Ansi.CYAN, Ansi.RESET))
        print_config_summary(config)

        # Ask for confirmation first
        confirm = get_input("\nUse these preset values? (y/n/edit) [y]: ").strip().lower()
        if confirm == "n":
            print("{}Cancelled. Returning to menu...{}".format(Ansi.YELLOW, Ansi.RESET))
            return create_simple_config()  # Start over
        elif confirm == "edit" or confirm == "e":
            config = edit_panel_constraints(config)

    else:  # choice == "3"
        print("\n{}=== CUSTOM CONFIGURATION ==={}".format(Ansi.CYAN, Ansi.RESET))
        config = create_custom_config()

    # Optional: project name override
    project_name = get_input("\nProject name [{}]: ".format(config.project_name)).strip()
    if project_name:
        config.project_name = project_name

    return config


def print_config_summary(config):
    """Display configuration parameters."""
    pc = config.panel_constraints
    print("\nPanel Constraints:")
    print("  Orientation:  {}".format(config.optimization_strategy.panel_orientation))
    print("  Min Width:    {}\"".format(pc.min_width))
    print("  Max Width:    {}\"".format(pc.max_width))
    print("  Min Height:   {}\"".format(pc.min_height))
    print("  Max Height:   {}\"".format(pc.max_height))
    print("  Short Max:    {}\"".format(pc.short_max))
    print("  Long Max:     {}\"".format(pc.long_max))
    print("  Increment:    {}\"".format(pc.dimension_increment))

    print("\nClearances:")
    print("  Doors:       jamb={}\" header={}\" sill={}\"".format(
        config.door_clearances.jamb_min,
        config.door_clearances.header_min,
        config.door_clearances.sill_min))
    print("  Windows:     jamb={}\" header={}\" sill={}\"".format(
        config.window_clearances.jamb_min,
        config.window_clearances.header_min,
        config.window_clearances.sill_min))
    print("  Storefronts: jamb={}\" header={}\" sill={}\"".format(
        config.storefront_clearances.jamb_min,
        config.storefront_clearances.header_min,
        config.storefront_clearances.sill_min))

def edit_panel_constraints(config):
    """Enhanced parameter editor with grouping, validation, and full customization."""
    try:
        get_input = raw_input  # type: ignore
    except NameError:
        get_input = input

    print("\n{}=== PARAMETER CUSTOMIZATION ==={}".format(Ansi.YELLOW, Ansi.RESET))
    print("\nWhat would you like to edit?")
    print("1. Panel Dimensions (width/height limits)")
    print("2. Clearances (doors, windows, storefronts)")
    print("3. Panel Orientation")
    print("4. All Parameters")
    print("5. Done (keep current values)")

    while True:
        choice = get_input("\nChoice (1-5) [5]: ").strip() or "5"

        if choice == "1":
            config = edit_panel_dimensions(config)
        elif choice == "2":
            config = edit_clearances(config)
        elif choice == "3":
            config = edit_orientation(config)
        elif choice == "4":
            config = edit_panel_dimensions(config)
            config = edit_clearances(config)
            config = edit_orientation(config)
            break
        elif choice == "5":
            break
        else:
            print("{}Invalid choice. Please enter 1-5.{}".format(Ansi.RED, Ansi.RESET))
            continue

        # After each edit, ask if they want to edit more
        if choice in ["1", "2", "3"]:
            more = get_input("\nEdit another section? (y/n) [n]: ").strip().lower()
            if more != "y":
                break

    print("\n{}Final Configuration:{}".format(Ansi.GREEN, Ansi.RESET))
    print_config_summary(config)

    confirm = get_input("\nUse this configuration? (y/n) [y]: ").strip().lower()
    if confirm == "n":
        print("{}Discarding changes...{}".format(Ansi.YELLOW, Ansi.RESET))
        # Return original preset
        presets = get_preset_configs()
        if "vertical" in config.project_name.lower():
            return presets["vertical"]
        else:
            return presets["horizontal"]

    return config


def edit_panel_dimensions(config):
    """Edit panel dimension constraints with validation."""
    try:
        get_input = raw_input  # type: ignore
    except NameError:
        get_input = input

    print("\n{}--- PANEL DIMENSIONS ---{}".format(Ansi.CYAN, Ansi.RESET))
    print("(Press Enter to keep current value)")

    pc = config.panel_constraints

    # Min Width
    while True:
        val = get_input("  Min Width [{}\"]: ".format(pc.min_width)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val <= 0:
                print("    {}Error: Must be positive{}".format(Ansi.RED, Ansi.RESET))
                continue
            if new_val >= pc.max_width:
                print("    {}Error: Must be less than Max Width ({}\"){}".format(
                    Ansi.RED, pc.max_width, Ansi.RESET))
                continue
            pc.min_width = new_val
            break
        except ValueError:
            print("    {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Max Width
    while True:
        val = get_input("  Max Width [{}\"]: ".format(pc.max_width)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val <= pc.min_width:
                print("    {}Error: Must be greater than Min Width ({}\"){}".format(
                    Ansi.RED, pc.min_width, Ansi.RESET))
                continue
            if new_val > pc.long_max:
                print("    {}Warning: Exceeds Long Max ({}\")-consider adjusting Long Max too{}".format(
                    Ansi.YELLOW, pc.long_max, Ansi.RESET))
            pc.max_width = new_val
            break
        except ValueError:
            print("    {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Min Height
    while True:
        val = get_input("  Min Height [{}\"]: ".format(pc.min_height)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val <= 0:
                print("    {}Error: Must be positive{}".format(Ansi.RED, Ansi.RESET))
                continue
            if new_val >= pc.max_height:
                print("    {}Error: Must be less than Max Height ({}\"){}".format(
                    Ansi.RED, pc.max_height, Ansi.RESET))
                continue
            pc.min_height = new_val
            break
        except ValueError:
            print("    {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Max Height
    while True:
        val = get_input("  Max Height [{}\"]: ".format(pc.max_height)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val <= pc.min_height:
                print("    {}Error: Must be greater than Min Height ({}\"){}".format(
                    Ansi.RED, pc.min_height, Ansi.RESET))
                continue
            if new_val > pc.long_max:
                print("    {}Warning: Exceeds Long Max ({}\")-consider adjusting Long Max too{}".format(
                    Ansi.YELLOW, pc.long_max, Ansi.RESET))
            pc.max_height = new_val
            break
        except ValueError:
            print("    {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Short Max
    while True:
        val = get_input("  Short Max (one dimension must be <= this) [{}\"]: ".format(pc.short_max)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val <= 0:
                print("    {}Error: Must be positive{}".format(Ansi.RED, Ansi.RESET))
                continue
            if new_val > pc.long_max:
                print("    {}Error: Must be <= Long Max ({}\"){}".format(
                    Ansi.RED, pc.long_max, Ansi.RESET))
                continue
            pc.short_max = new_val
            break
        except ValueError:
            print("    {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Long Max
    while True:
        val = get_input("  Long Max (absolute maximum for either dimension) [{}\"]: ".format(pc.long_max)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val < pc.short_max:
                print("    {}Error: Must be >= Short Max ({}\"){}".format(
                    Ansi.RED, pc.short_max, Ansi.RESET))
                continue
            pc.long_max = new_val
            break
        except ValueError:
            print("    {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Dimension Increment
    while True:
        val = get_input("  Dimension Increment (snap grid) [{}\"]: ".format(pc.dimension_increment)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val <= 0:
                print("    {}Error: Must be positive{}".format(Ansi.RED, Ansi.RESET))
                continue
            if new_val > 12:
                print("    {}Warning: Large increment ({}\")-panels may not fit well{}".format(
                    Ansi.YELLOW, new_val, Ansi.RESET))
            pc.dimension_increment = new_val
            break
        except ValueError:
            print("    {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    print("  {}✓ Panel dimensions updated{}".format(Ansi.GREEN, Ansi.RESET))
    return config


def edit_clearances(config):
    """Edit clearance values for openings."""
    try:
        get_input = raw_input  # type: ignore
    except NameError:
        get_input = input

    print("\n{}--- CLEARANCES (inches) ---{}".format(Ansi.CYAN, Ansi.RESET))
    print("(Press Enter to keep current value)")

    # Door Clearances
    print("\n  {}DOOR CLEARANCES:{}".format(Ansi.BOLD, Ansi.RESET))
    config.door_clearances = edit_opening_clearance(
        "Door", config.door_clearances)

    # Window Clearances
    print("\n  {}WINDOW CLEARANCES:{}".format(Ansi.BOLD, Ansi.RESET))
    config.window_clearances = edit_opening_clearance(
        "Window", config.window_clearances)

    # Storefront Clearances
    print("\n  {}STOREFRONT CLEARANCES:{}".format(Ansi.BOLD, Ansi.RESET))
    config.storefront_clearances = edit_opening_clearance(
        "Storefront", config.storefront_clearances)

    print("  {}✓ Clearances updated{}".format(Ansi.GREEN, Ansi.RESET))
    return config


def edit_opening_clearance(opening_type, clearances):
    """Edit clearances for a specific opening type."""
    try:
        get_input = raw_input  # type: ignore
    except NameError:
        get_input = input

    # Jamb
    while True:
        val = get_input("    Jamb (left/right) [{}\"]: ".format(clearances.jamb_min)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val < 0:
                print("      {}Error: Cannot be negative{}".format(Ansi.RED, Ansi.RESET))
                continue
            if new_val > 24:
                print("      {}Warning: Large clearance ({}\")-may reduce coverage{}".format(
                    Ansi.YELLOW, new_val, Ansi.RESET))
            clearances.jamb_min = new_val
            break
        except ValueError:
            print("      {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Header
    while True:
        val = get_input("    Header (top) [{}\"]: ".format(clearances.header_min)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val < 0:
                print("      {}Error: Cannot be negative{}".format(Ansi.RED, Ansi.RESET))
                continue
            if new_val > 24:
                print("      {}Warning: Large clearance ({}\")-may reduce coverage{}".format(
                    Ansi.YELLOW, new_val, Ansi.RESET))
            clearances.header_min = new_val
            break
        except ValueError:
            print("      {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    # Sill
    while True:
        val = get_input("    Sill (bottom) [{}\"]: ".format(clearances.sill_min)).strip()
        if not val:
            break
        try:
            new_val = float(val)
            if new_val < 0:
                print("      {}Error: Cannot be negative{}".format(Ansi.RED, Ansi.RESET))
                continue
            if new_val > 24:
                print("      {}Warning: Large clearance ({}\")-may reduce coverage{}".format(
                    Ansi.YELLOW, new_val, Ansi.RESET))
            clearances.sill_min = new_val
            break
        except ValueError:
            print("      {}Error: Invalid number{}".format(Ansi.RED, Ansi.RESET))

    return clearances


def edit_orientation(config):
    """Change panel orientation."""
    try:
        get_input = raw_input  # type: ignore
    except NameError:
        get_input = input

    print("\n{}--- PANEL ORIENTATION ---{}".format(Ansi.CYAN, Ansi.RESET))
    current = config.optimization_strategy.panel_orientation
    print("  Current: {}".format(current))
    print("\n  1. Vertical (tall panels)")
    print("  2. Horizontal (wide panels)")

    choice = get_input("\nChoice (1-2) or Enter to keep [{}]: ".format(current)).strip()

    if choice == "1":
        config.optimization_strategy.panel_orientation = "vertical"
        config.optimization_strategy.prefer_full_height_panels = True
        print("  {}✓ Changed to VERTICAL{}".format(Ansi.GREEN, Ansi.RESET))
    elif choice == "2":
        config.optimization_strategy.panel_orientation = "horizontal"
        config.optimization_strategy.prefer_full_height_panels = False
        print("  {}✓ Changed to HORIZONTAL{}".format(Ansi.GREEN, Ansi.RESET))
    else:
        print("  Keeping current orientation: {}".format(current))

    return config

def create_custom_config():
    """Create fully custom configuration using the enhanced editing workflow."""
    try:
        get_input = raw_input  # type: ignore
    except NameError:
        get_input = input

    print("\n{}=== CUSTOM CONFIGURATION ==={}".format(Ansi.CYAN, Ansi.RESET))
    print("Starting with default values. You'll customize each section.")

    # Start with a default base config
    presets = get_preset_configs()
    config = presets["vertical"]  # Use vertical as starting point
    config.project_name = "Custom Configuration"

    # Orientation first
    print("\n{}Step 1: Panel Orientation{}".format(Ansi.BOLD, Ansi.RESET))
    print("  1. Vertical (tall panels)")
    print("  2. Horizontal (wide panels)")

    orient_choice = get_input("\nChoice (1-2) [1]: ").strip() or "1"
    if orient_choice == "2":
        config = presets["horizontal"]
        config.project_name = "Custom Configuration"
        config.optimization_strategy.panel_orientation = "horizontal"
        print("  {}✓ Starting with horizontal preset{}".format(Ansi.GREEN, Ansi.RESET))
    else:
        config.optimization_strategy.panel_orientation = "vertical"
        print("  {}✓ Starting with vertical preset{}".format(Ansi.GREEN, Ansi.RESET))

    # Show starting values
    print("\n{}Starting Configuration:{}".format(Ansi.CYAN, Ansi.RESET))
    print_config_summary(config)

    # Panel Dimensions
    print("\n{}Step 2: Panel Dimensions{}".format(Ansi.BOLD, Ansi.RESET))
    customize = get_input("Customize panel dimensions? (y/n) [y]: ").strip().lower()
    if customize != "n":
        config = edit_panel_dimensions(config)

    # Clearances
    print("\n{}Step 3: Clearances{}".format(Ansi.BOLD, Ansi.RESET))
    customize = get_input("Customize clearances? (y/n) [y]: ").strip().lower()
    if customize != "n":
        config = edit_clearances(config)

    # Final review
    print("\n{}=== FINAL CUSTOM CONFIGURATION ==={}".format(Ansi.GREEN, Ansi.RESET))
    print_config_summary(config)

    confirm = get_input("\nUse this configuration? (y/n) [y]: ").strip().lower()
    if confirm == "n":
        print("{}Cancelled. Using default vertical preset.{}".format(Ansi.YELLOW, Ansi.RESET))
        return presets["vertical"]

    return config




# =============================================================================
# SECTION 7: MAIN ENTRY POINT
# =============================================================================
def main():
    """
    OPTIONAL standalone CLI entry point for debugging only.
    The UI script should provide input/output folders.
    """

    global ACTIVE_CONFIG

    print(Ansi.YELLOW + "[INFO] No input directory provided. "
                        "This CLI mode is only for manual debugging." + Ansi.RESET)

    try:
        input_dir = raw_input("Enter input folder path: ").strip()
    except NameError:
        input_dir = input("Enter input folder path: ").strip()

    if not input_dir or not os.path.isdir(input_dir):
        print(Ansi.RED + "[ERROR] Invalid folder. Exiting." + Ansi.RESET)
        return

    OUTPUT_DIR = input_dir   # Save results next to input files unless user changes it

    GENERATE_VISUALIZATIONS = True
    SAVE_VISUALIZATIONS_AS_PNG = True

    # Load or create config inside input_dir
    config_file = os.path.join(input_dir, "optimizer_config.json")
    if os.path.exists(config_file):
        print(Ansi.CYAN + "[CONFIG] Loading: {}".format(config_file) + Ansi.RESET)
        config = OptimizerConfig.load(config_file)
    else:
        print(Ansi.YELLOW + "[CONFIG] No configuration found. Creating new..." + Ansi.RESET)
        config = create_simple_config()
        config.save(config_file)

    ACTIVE_CONFIG = config

    used_config_path = os.path.join(input_dir, "config_used.json")
    config.save(used_config_path)
    print(" Saved run config to: {}".format(used_config_path))

    walls_csv = os.path.join(input_dir, "walls.csv")
    openings_csv = os.path.join(input_dir, "wall_openings.csv")

    if not os.path.exists(walls_csv):
        print(Ansi.RED + "[ERROR] walls.csv not found. Exiting." + Ansi.RESET)
        return

    openings_rows = []
    if os.path.exists(openings_csv):
        openings_rows = load_openings_from_csv(openings_csv)
    else:
        print(Ansi.YELLOW + "[WARN] wall_openings.csv missing. Continuing without openings." + Ansi.RESET)

    walls_rows = load_walls_from_csv(walls_csv)

    panels_path, config_path = process_all_walls(
        walls_rows, openings_rows, input_dir,
        config.door_clearances,
        config.window_clearances,
        config.storefront_clearances
    )


    # Copy config next to placement file
    if panels_path:
        dst = os.path.join(os.path.dirname(panels_path), "config_used.json")
        if dst != used_config_path:
            import shutil
            shutil.copy(used_config_path, dst)
            print("Copied config to: {}".format(dst))

    if GENERATE_VISUALIZATIONS and panels_path:
        try:
            visualize_all_walls(
                panels_path,
                openings_csv,
                walls_csv,
                input_dir,
                save_as_image=SAVE_VISUALIZATIONS_AS_PNG
            )
            print(" Visualization complete.")
        except Exception as e:
            print("[VIS ERROR]", e)


if __name__ == "__main__":
    # Do NOT run automatically inside pyRevit.
    # Only executes if someone manually runs calculator.py from command line.
    main()