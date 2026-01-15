# -*- coding: utf-8 -*-
"""
REVIT PANEL PLACEMENT - FIXED VERSION
Reads the optimized panel placement CSV and places panels in Revit.

FIXES:
- Uses consistent perpendicular calculation for wall normal
- Respects wall flip state
- Proper core face offset calculation
- Enhanced debug output

- Supports optional CSV columns:
    rotation_deg : float, rotation around vertical axis in degrees
    x_ref        : "start" or "end" ("right" treated as "end")
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector, Wall, Transaction, XYZ, Line,
    FamilySymbol, BuiltInCategory, BuiltInParameter, Transform, ElementId,
    DirectShape, ElementTransformUtils, FamilyPlacementType
)

try:
    from Autodesk.Revit.DB.Structure import StructuralType
except:
    from Autodesk.Revit.DB import Structure
    StructuralType = Structure.StructuralType

from pyrevit import revit
import csv
import os
import json
import math

import clr
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import FolderBrowserDialog, DialogResult

doc = revit.doc

# ========== SETTINGS ==========
DEFAULT_INPUT_DIR = None
PANELS_FILE = "optimized_panel_placement.csv"
USE_FOLDER_PICKER = True
PANEL_FAMILY_NAME = None

SHOW_CUTOUTS = True
CUTOUT_THICKNESS_IN = 2.0
CUTOUT_DEPTH_IN = 3.0
PANEL_FRONT_OFFSET_IN = 0.0
ALLOW_TYPE_PARAM_CHANGE = True

# KEY SETTING: Most Revit families use center-based insertion
# Set to False if your family uses corner-based insertion
FAMILY_ORIGIN_IS_CENTER = False

# How x_in is interpreted when CSV does not specify x_ref:
#   "start" -> 0 at wall_start, increasing along wall_dir
#   "end"   -> 0 at wall_end, increasing opposite wall_dir
PANEL_COORD_DEFAULT_REF = "start"

# If True, script will read rotation_deg from CSV (if present)
USE_CSV_ROTATION = True

# Runtime overrides (set via UI in main)
X_REF_OVERRIDE = None
ROTATION_OVERRIDE_DEG = None

WIDTH_PARAM_CANDIDATES = [
    "Width", "Panel Width", "W", "Overall Width",
    "Length", "L"
]

HEIGHT_PARAM_CANDIDATES = [
    "Height", "Panel Height", "H", "Overall Height",
    "Thickness", "Depth"
]

DISABLE_WALL_FLIP = False
TEST_MODE_LIMIT_PANELS = 0  # 0 = place all panels

# +1 = same side as DirectShape, -1 = opposite side
PANEL_SIDE_SIGN = 1

# Extend wall ends for joined/mitered walls (common ~half wall thickness “missing”)
USE_WALL_ENDCAP_EXTENSION = True



# ========== UTILITIES ==========
def _pick_input_folder(default_dir=None):
    """Ask user for the folder that contains the panel CSV."""
    try:
        fbd = FolderBrowserDialog()
        fbd.Description = "Select the folder containing '{0}'".format(PANELS_FILE)
        if default_dir:
            try:
                if os.path.isdir(default_dir):
                    fbd.SelectedPath = default_dir
            except:
                pass
        result = fbd.ShowDialog()
        if result == DialogResult.OK and fbd.SelectedPath and os.path.isdir(fbd.SelectedPath):
            return str(fbd.SelectedPath)
    except Exception as e:
        print("Folder picker failed: {0}".format(e))
    return None


def norm_id(val):
    try:
        return str(int(float(val)))
    except:
        return str(val).strip()


def get_wall_by_id(wall_id):
    try:
        elem_id = int(float(wall_id))
        element = doc.GetElement(ElementId(elem_id))
        if isinstance(element, Wall):
            return element
    except Exception as e:
        print("[WARN] Could not get wall {0}: {1}".format(wall_id, e))
    return None


def _feet(val_inch):
    return float(val_inch) / 12.0


# ========== ELEMENT INFO HELPERS ==========
def get_element_name(element):
    try:
        name_param = element.get_Parameter(BuiltInParameter.SYMBOL_NAME)
        if name_param:
            return name_param.AsString()
    except:
        pass
    try:
        return element.Name
    except:
        pass
    return "Unknown"


def get_family_name(family_symbol):
    try:
        return family_symbol.Family.Name
    except:
        try:
            return family_symbol.FamilyName
        except:
            return "Unknown Family"


def get_all_family_symbols():
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    families_dict = {}
    for symbol in collector:
        family_name = get_family_name(symbol)
        families_dict.setdefault(family_name, []).append(symbol)
    return families_dict


def select_family_interactive():
    from pyrevit import forms
    families_dict = get_all_family_symbols()
    family_names = sorted(families_dict.keys())
    family_names.insert(0, "< Use DirectShape (3D Solid Panels) >")
    selected_family = forms.SelectFromList.show(
        family_names, title="Select Panel Placement Method",
        button_name="Select", multiselect=False
    )
    if not selected_family:
        return None, False
    if selected_family == "< Use DirectShape (3D Solid Panels) >":
        return None, True
    symbols = families_dict[selected_family]
    if len(symbols) == 1:
        return symbols[0], False
    symbol_names = [get_element_name(s) for s in symbols]
    selected_type = forms.SelectFromList.show(
        symbol_names,
        title="Select Family Type for '{0}'".format(selected_family),
        button_name="Select",
        multiselect=False
    )
    if not selected_type:
        return symbols[0], False
    for symbol in symbols:
        if get_element_name(symbol) == selected_type:
            return symbol, False
    return symbols[0], False


def get_panel_family_symbol(family_name):
    if family_name is None:
        return select_family_interactive()
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    for symbol in collector:
        if get_family_name(symbol) == family_name:
            return symbol, False
    return None, False


def ensure_symbol_active(symbol):
    try:
        if not symbol.IsActive:
            symbol.Activate()
        return True
    except Exception as e:
        print("[WARN] Could not activate symbol '{0}': {1}".format(get_element_name(symbol), e))
        return False


# Global cache so we only compute once
REF_LEVEL_ELEVATION = None

def get_reference_level_elevation():
    """
    Returns the elevation (ft) of the global reference level used for vertical origin.
    - Prefer a level named 'Level 1'
    - Otherwise use the lowest level by elevation
    """
    global REF_LEVEL_ELEVATION
    if REF_LEVEL_ELEVATION is not None:
        return REF_LEVEL_ELEVATION

    try:
        levels = list(
            FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Levels)
            .WhereElementIsNotElementType()
        )
    except:
        levels = []

    if not levels:
        REF_LEVEL_ELEVATION = 0.0
        return REF_LEVEL_ELEVATION

    # Try to find a level literally named "Level 1"
    lvl1 = None
    for lvl in levels:
        try:
            if (lvl.Name or "").strip().lower() == "level 1":
                lvl1 = lvl
                break
        except:
            continue

    if lvl1 is not None:
        REF_LEVEL_ELEVATION = lvl1.Elevation
        print("Using 'Level 1' as reference level (elev = {0:.2f} ft)".format(REF_LEVEL_ELEVATION))
        return REF_LEVEL_ELEVATION

    # Otherwise, pick the lowest level elevation
    min_lvl = min(levels, key=lambda l: getattr(l, "Elevation", 0.0))
    REF_LEVEL_ELEVATION = getattr(min_lvl, "Elevation", 0.0)
    print("Using lowest level '{0}' as reference level (elev = {1:.2f} ft)".format(
        getattr(min_lvl, "Name", "Unknown"),
        REF_LEVEL_ELEVATION
    ))
    return REF_LEVEL_ELEVATION


# ========== PARAMETER SETTERS ==========
def _find_param_by_candidates(element, candidates):
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(nm.lower() == cand.lower() for cand in candidates):
                return p
        except:
            continue
    lower_cands = [c.lower() for c in candidates]
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(c in nm.lower() for c in lower_cands):
                return p
        except:
            continue
    return None


def set_size_parameters(inst, width_in, height_in, symbol=None):
    width_ft = _feet(width_in)
    height_ft = _feet(height_in)

    w_param = _find_param_by_candidates(inst, WIDTH_PARAM_CANDIDATES)
    h_param = _find_param_by_candidates(inst, HEIGHT_PARAM_CANDIDATES)

    changed = False

    try:
        if w_param and not w_param.IsReadOnly:
            w_param.Set(width_ft)
            changed = True
        if h_param and not h_param.IsReadOnly:
            h_param.Set(height_ft)
            changed = True
    except Exception as e:
        print("[WARN] Setting instance parameters failed: {0}".format(e))

    if not changed and ALLOW_TYPE_PARAM_CHANGE and symbol:
        try:
            wtp = _find_param_by_candidates(symbol, WIDTH_PARAM_CANDIDATES)
            htp = _find_param_by_candidates(symbol, HEIGHT_PARAM_CANDIDATES)
            if wtp and not wtp.IsReadOnly:
                wtp.Set(width_ft)
                changed = True
            if htp and not htp.IsReadOnly:
                htp.Set(height_ft)
                changed = True
            if changed:
                print("[INFO] Width/Height set on TYPE '{0}' (affects all instances).".format(
                    get_element_name(symbol)))
        except Exception as e:
            print("[WARN] Setting type parameters failed: {0}".format(e))

    if not changed:
        print("[WARN] Could not set Width/Height. Ensure your family exposes adjustable parameters.")


# ========== LEVEL HELPERS ==========
def get_wall_base_level(wall):
    try:
        lvl_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        if lvl_id and lvl_id.IntegerValue > 0:
            lvl = doc.GetElement(lvl_id)
            return lvl
    except:
        pass
    try:
        lvl = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).FirstElement()
        return lvl
    except:
        return None


# ========== GEOMETRY HELPERS (FIXED) ==========

def get_core_face_offset(wall):
    """
    Returns offset (ft) from the wall location line to the exterior core face.
    For basic walls without compound structure, returns 0.
    """
    try:
        cs = wall.WallType.GetCompoundStructure()
        if cs:
            # Get core layer boundaries
            core_start = cs.GetCoreStartOffset()  # ft from location line
            core_end = cs.GetCoreEndOffset()      # ft from location line
            
            # Return the exterior face offset (closest to zero, or most negative)
            core_offset = min(core_start, core_end)
            
            print("    [CORE] start={0:.4f} ft, end={1:.4f} ft, using={2:.4f} ft".format(
                core_start, core_end, core_offset))
            
            return core_offset
    except Exception as e:
        print("    [CORE] No compound structure or error: {0}".format(e))
    
    return 0.0

def compute_panel_base_point(wall, panel):
    """
    FIXED VERSION: Computes panel insertion point with consistent wall normal,
    and compensates for wall-join / miter endcap offset (common ~half wall thickness).

    Returns: (base_point, wall_dir, wall_normal)
    """
    loc_curve = wall.Location.Curve
    wall_start = loc_curve.GetEndPoint(0)
    wall_end   = loc_curve.GetEndPoint(1)
    wall_dir   = (wall_end - wall_start).Normalize()

    # ===== FIX 1: Use consistent perpendicular calculation =====
    up = XYZ(0, 0, 1)
    wall_normal = wall_dir.CrossProduct(up).Normalize()

    # ===== FIX 2: Respect wall flip state =====
    try:
        if wall.Flipped:
            wall_normal = wall_normal.Negate()
            print("    [FLIP] Wall is flipped, reversed normal")
    except:
        pass

    print("    [GEOMETRY] Wall direction: ({0:.3f}, {1:.3f}, {2:.3f})".format(
        wall_dir.X, wall_dir.Y, wall_dir.Z))
    print("    [GEOMETRY] Wall normal: ({0:.3f}, {1:.3f}, {2:.3f})".format(
        wall_normal.X, wall_normal.Y, wall_normal.Z))

    # --- Wall base from bounding box ---
    bb = wall.get_BoundingBox(None)
    wall_base_z = bb.Min.Z

    # --- Read panel coordinates (inches) ---
    x_in = float(panel.get("x_in", 0.0) or 0.0)
    y_in = float(panel.get("y_in", 0.0) or 0.0)

    # --- X reference along the wall ---
    x_ref = (panel.get("x_ref", PANEL_COORD_DEFAULT_REF) or PANEL_COORD_DEFAULT_REF).lower().strip()
    if X_REF_OVERRIDE in ("start", "end"):
        x_ref = X_REF_OVERRIDE

    x_ft = _feet(x_in)

    # ===== FIX 0: compensate for wall join / miter endcap =====
    # Revit wall location curve endpoints can be inset from the visible wall face ends
    # by ~half the wall thickness (e.g., 8" wall -> 4" inset). Extend endpoints.
    wall_start_used = wall_start
    wall_end_used   = wall_end
    try:
        # Toggle with your global if you have it; otherwise treat as always on.
        use_endcap = True
        try:
            use_endcap = bool(USE_WALL_ENDCAP_EXTENSION)
        except:
            use_endcap = True

        if use_endcap:
            half_thk = wall.Width / 2.0  # feet
            wall_start_used = wall_start - wall_dir * half_thk
            wall_end_used   = wall_end   + wall_dir * half_thk
            print("    [ENDCAP] half_thickness = {0:.4f} ft ({1:.2f}\")".format(
                half_thk, half_thk * 12.0))
    except Exception as e:
        print("    [ENDCAP] Could not apply endcap extension: {0}".format(e))

    # --- Use extended endpoints for XY along-wall placement ---
    if x_ref == "start":
        pt_xy = wall_start_used + wall_dir * x_ft
    else:
        pt_xy = wall_end_used - wall_dir * x_ft

    # --- Y from bottom (CSV coordinate) ---
    y_ft = _feet(y_in)
    base_point_before = XYZ(pt_xy.X, pt_xy.Y, wall_base_z + y_ft)

    print("    [COORD] x_in={0:.2f}\", y_in={1:.2f}\" (from {2})".format(
        x_in, y_in, x_ref))
    print("    [POINT] Before offset: ({0:.3f}, {1:.3f}, {2:.3f})".format(
        base_point_before.X, base_point_before.Y, base_point_before.Z))

    # ===== FIX 3: Apply offsets correctly =====
    core_offset = get_core_face_offset(wall)
    front_offset = _feet(PANEL_FRONT_OFFSET_IN) * float(PANEL_SIDE_SIGN)
    total_offset = core_offset + front_offset

    print("    [OFFSET] Core: {0:.4f} ft ({1:.2f}\")".format(core_offset, core_offset * 12))
    print("    [OFFSET] Front: {0:.4f} ft ({1:.2f}\")".format(front_offset, front_offset * 12))
    print("    [OFFSET] Total: {0:.4f} ft ({1:.2f}\")".format(total_offset, total_offset * 12))
    print("    [OFFSET] Side sign: {0}".format(PANEL_SIDE_SIGN))

    base_point = base_point_before + (wall_normal * total_offset)

    print("    [POINT] After offset: ({0:.3f}, {1:.3f}, {2:.3f})".format(
        base_point.X, base_point.Y, base_point.Z))

    return base_point, wall_dir, wall_normal


# ========== FAMILY PLACEMENT ==========
def place_panel_family(wall, panel, symbol):
    """
    Places a panel family instance at the correct location.
    - Uses x_ref ('start' / 'end') to interpret x_in
    - Optionally rotates by rotation_deg from CSV
    """
    if not ensure_symbol_active(symbol):
        return None

    # --- Compute geometry & insertion point ---
    try:
        base_point, wall_dir, wall_normal = compute_panel_base_point(wall, panel)
    except Exception as e:
        print("[ERROR] Could not compute base point: {0}".format(e))
        import traceback
        print(traceback.format_exc())
        return None

    # --- Place the family instance ---
    inst = None
    try:
        # Try host-based on wall first
        inst = doc.Create.NewFamilyInstance(
            base_point, symbol, wall, StructuralType.NonStructural
        )
        print("  [PLACE] Hosted instance: {0}".format(panel.get("panel_name", "")))
    except Exception as e:
        print("  [PLACE] Hosted placement failed: {0}".format(e))

    if inst is None:
        level = get_wall_base_level(wall)
        try:
            if level:
                inst = doc.Create.NewFamilyInstance(
                    base_point, symbol, level, StructuralType.NonStructural
                )
            else:
                inst = doc.Create.NewFamilyInstance(
                    base_point, symbol, StructuralType.NonStructural
                )
            print("  [PLACE] Non-hosted instance: {0}".format(panel.get("panel_name", "")))
        except Exception as e:
            print("[ERROR] Non-hosted placement failed: {0}".format(e))
            return None

    if inst is None:
        print("[ERROR] Family placement failed for panel '{0}'".format(
            panel.get("panel_name", "")))
        return None

    # --- Force geometry update ---
    try:
        doc.Regenerate()
    except:
        pass

    # --- Rotate around vertical axis ---
    global ROTATION_OVERRIDE_DEG
    rot_deg = 0.0

    if ROTATION_OVERRIDE_DEG is not None:
        rot_deg = ROTATION_OVERRIDE_DEG
    elif USE_CSV_ROTATION:
        try:
            rot_deg = float(panel.get("rotation_deg", 0.0) or 0.0)
        except:
            rot_deg = 0.0

    if abs(rot_deg) > 1e-3:
        try:
            try:
                inst_location = inst.Location.Point
            except:
                inst_location = base_point

            rotation_axis = Line.CreateBound(
                inst_location,
                inst_location + XYZ(0, 0, 10.0)
            )
            rotation_angle = math.radians(rot_deg)

            ElementTransformUtils.RotateElement(
                doc, inst.Id, rotation_axis, rotation_angle
            )

            print("    [ROTATE] {0:.1f}° around vertical axis".format(rot_deg))
        except Exception as e:
            print("[WARN] Rotation failed: {0}".format(e))

    # --- Set size parameters ---
    set_size_parameters(inst, panel["width_in"], panel["height_in"], symbol=symbol)

    # --- Set panel name / mark ---
    name_param = _find_param_by_candidates(inst, ["Name", "Panel Name", "Mark"])
    try:
        if name_param and not name_param.IsReadOnly:
            name_param.Set(panel.get("panel_name", "") or "")
    except:
        pass

    return inst


# ========== DIRECTSHAPE PANEL ==========
def create_panel_as_direct_shape(wall, panel):
    """Creates a 3D box representing the panel (DirectShape)."""
    try:
        base_point, wall_dir, wall_normal = compute_panel_base_point(wall, panel)
        
        w_in = float(panel.get('width_in', 0.0) or 0.0)
        h_in = float(panel.get('height_in', 0.0) or 0.0)
        
        w_ft = _feet(w_in)
        h_ft = _feet(h_in)

        thickness = 1.0 / 12.0
        front = wall_normal * 0.01
        back = wall_normal * (0.01 + thickness)

        v1 = base_point + front
        v2 = v1 + wall_dir * w_ft
        v3 = v2 + XYZ(0, 0, h_ft)
        v4 = v1 + XYZ(0, 0, h_ft)

        v5 = base_point + back
        v6 = v5 + wall_dir * w_ft
        v7 = v6 + XYZ(0, 0, h_ft)
        v8 = v5 + XYZ(0, 0, h_ft)

        lines = [
            Line.CreateBound(v1, v2), Line.CreateBound(v2, v3),
            Line.CreateBound(v3, v4), Line.CreateBound(v4, v1),
            Line.CreateBound(v5, v6), Line.CreateBound(v6, v7),
            Line.CreateBound(v7, v8), Line.CreateBound(v8, v5),
            Line.CreateBound(v1, v5), Line.CreateBound(v2, v6),
            Line.CreateBound(v3, v7), Line.CreateBound(v4, v8)
        ]
        
        ds = DirectShape.CreateElement(doc, ElementId(int(BuiltInCategory.OST_GenericModel)))
        ds.SetShape(lines)
        ds.Name = panel.get('panel_name', 'PanelSolid') or 'PanelSolid'
        
        print("  [DIRECTSHAPE] Created: {0}".format(ds.Name))
        return ds
        
    except Exception as e:
        print("[ERROR] DirectShape creation failed: {0}".format(str(e)))
        import traceback
        print(traceback.format_exc())
        return None


# ========== CUTOUT VISUALIZATION ==========
def create_cutout_visualization(wall, panel, cutout_data, symbol, use_directshape):
    """
    Place a cutout marker (family instance or DirectShape).
    
    Conventions:
      - panel x_in, y_in are GLOBAL (from wall START/END and from BOTTOM)
      - cutout_data x_in, y_in are LOCAL to panel's bottom-left corner
    """

    # -------- FAMILY-BASED CUTOUT --------
    if not use_directshape and symbol is not None:
        try:
            cutout_local_x = float(cutout_data.get("x_in", 0.0) or 0.0)
            cutout_local_y = float(cutout_data.get("y_in", 0.0) or 0.0)
            cutout_w = float(cutout_data.get("width_in", 0.0) or 0.0)
            cutout_h = float(cutout_data.get("height_in", 0.0) or 0.0)

            # Calculate global coordinates
            cutout_global_x = float(panel.get("x_in", 0.0) or 0.0) + cutout_local_x
            cutout_global_y = float(panel.get("y_in", 0.0) or 0.0) + cutout_local_y

            # Create fake panel dict to reuse place_panel_family()
            fake_panel = {
                "panel_name": "CUT_{0}".format(cutout_data.get("id", "")),
                "panel_type": panel.get("panel_type", ""),
                "wall_id": panel.get("wall_id", ""),
                "x_in": cutout_global_x,
                "y_in": cutout_global_y,
                "width_in": cutout_w,
                "height_in": cutout_h,
                "rotation_deg": panel.get("rotation_deg", 0.0),
                "x_ref": panel.get("x_ref", PANEL_COORD_DEFAULT_REF),
                "cutouts": []
            }

            print("    [CUTOUT] Family: local x={0:.2f}\", global x={1:.2f}\"".format(
                cutout_local_x, cutout_global_x))

            inst = place_panel_family(wall, fake_panel, symbol)
            if inst is not None:
                print("    [CUTOUT] Placed family instance")
            return inst

        except Exception as e:
            print("[WARN] Cutout family placement failed: {0}".format(e))
            # Fall through to DirectShape

    # -------- DIRECTSHAPE CUTOUT --------
    try:
        base_point, wall_dir, wall_normal = compute_panel_base_point(wall, panel)
        
        # Read cutout dimensions
        cutout_local_x = float(cutout_data.get("x_in", 0.0) or 0.0)
        cutout_local_y = float(cutout_data.get("y_in", 0.0) or 0.0)
        cutout_w_in = float(cutout_data.get("width_in", 0.0) or 0.0)
        cutout_h_in = float(cutout_data.get("height_in", 0.0) or 0.0)
        
        # Offset from panel base point
        cutout_offset_x = _feet(cutout_local_x)
        cutout_offset_y = _feet(cutout_local_y)
        cutout_w_ft = _feet(cutout_w_in)
        cutout_h_ft = _feet(cutout_h_in)
        
        # Cutout base point
        cutout_base = base_point + wall_dir * cutout_offset_x + XYZ(0, 0, cutout_offset_y)
        
        margin = _feet(CUTOUT_THICKNESS_IN) / 2.0
        depth = _feet(CUTOUT_DEPTH_IN)
        
        # Build cutout box
        v1 = cutout_base - wall_dir * margin - XYZ(0, 0, margin)
        v2 = v1 + wall_dir * (cutout_w_ft + 2 * margin)
        v3 = v2 + XYZ(0, 0, cutout_h_ft + 2 * margin)
        v4 = v1 + XYZ(0, 0, cutout_h_ft + 2 * margin)

        back_offset = wall_normal * depth
        v5 = v1 + back_offset
        v6 = v2 + back_offset
        v7 = v3 + back_offset
        v8 = v4 + back_offset

        lines = [
            Line.CreateBound(v1, v2), Line.CreateBound(v2, v3),
            Line.CreateBound(v3, v4), Line.CreateBound(v4, v1),
            Line.CreateBound(v5, v6), Line.CreateBound(v6, v7),
            Line.CreateBound(v7, v8), Line.CreateBound(v8, v5),
            Line.CreateBound(v1, v5), Line.CreateBound(v2, v6),
            Line.CreateBound(v3, v7), Line.CreateBound(v4, v8)
        ]

        ds = DirectShape.CreateElement(doc, ElementId(int(BuiltInCategory.OST_GenericModel)))
        ds.SetShape(lines)

        cutout_type = cutout_data.get('type', 'unknown')
        cutout_id = cutout_data.get('id', 'Unknown')
        ds.Name = "Cutout_{0}_{1}".format(cutout_type, cutout_id)

        print("    [CUTOUT] DirectShape created")
        return ds

    except Exception as e:
        print("[ERROR] DirectShape cutout failed: {0}".format(str(e)))
        return None



# ========== MAIN EXECUTION ==========
def main():
    print("\n" + "=" * 70)
    print("REVIT PANEL PLACEMENT - ADAPTIVE VERSION")
    print("=" * 70 + "\n")

    # --- Pick input folder ---
    if USE_FOLDER_PICKER:
        input_dir = _pick_input_folder(DEFAULT_INPUT_DIR)
        if not input_dir:
            print("[ERROR] No folder selected. Aborting panel placement.")
            return
    else:
        if DEFAULT_INPUT_DIR and os.path.isdir(DEFAULT_INPUT_DIR):
            input_dir = DEFAULT_INPUT_DIR
        else:
            input_dir = os.getcwd()

    panels_path = os.path.join(input_dir, PANELS_FILE)
    if not os.path.exists(panels_path):
        print("[ERROR] Panels file not found: {0}".format(panels_path))
        return

    print("Using panels file: {0}".format(panels_path))

    # --- Load panels from CSV ---
    panels = []
    with open(panels_path, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            cutouts_json = row.get("cutouts_json", "")
            cutouts = []
            if cutouts_json:
                try:
                    cutouts = json.loads(cutouts_json)
                except:
                    pass

            # Some exports might use 'panel_typ' instead of 'panel_type'
            panel_type_val = row.get("panel_type", "") or row.get("panel_typ", "")

            try:
                rotation_deg_val = float(row.get("rotation_deg", 0) or 0)
            except:
                rotation_deg_val = 0.0

            x_ref_val = (row.get("x_ref", "") or PANEL_COORD_DEFAULT_REF).lower()

            panels.append({
                "panel_name": row.get("panel_name", ""),
                "panel_type": panel_type_val,
                "wall_id": norm_id(row.get("wall_id", "")),
                "x_in": float(row.get("x_in", 0) or 0),
                "y_in": float(row.get("y_in", 0) or 0),
                "width_in": float(row.get("width_in", 0) or 0),
                "height_in": float(row.get("height_in", 0) or 0),
                "rotation_deg": rotation_deg_val,
                "x_ref": x_ref_val,
                "cutouts": cutouts
            })

    print("Loaded {0} panels to place".format(len(panels)))
    total_cutouts = sum(len(p['cutouts']) for p in panels)
    print("Total cutouts across all panels: {0}".format(total_cutouts))
    # --- Select family or DirectShape ---
    family_symbol, use_directshape = get_panel_family_symbol(PANEL_FAMILY_NAME)
    if family_symbol is None and not use_directshape:
        print("\n[ERROR] No family selected!")
        return

    from pyrevit import forms
    global X_REF_OVERRIDE, ROTATION_OVERRIDE_DEG

    if not use_directshape:
        print("\nUsing family: {0}".format(get_family_name(family_symbol)))
        print("Type: {0}".format(get_element_name(family_symbol)))
        print("Insertion mode: {0}".format(
            "CENTER-BASED" if FAMILY_ORIGIN_IS_CENTER else "CORNER-BASED"))
        print("Default x_ref: {0}".format(PANEL_COORD_DEFAULT_REF))

        # --- UI: choose x_ref behavior for this run ---
        xref_options = [
            "Use CSV / default (per panel)",
            "Force from wall START (left side)",
            "Force from wall END (right side)"
        ]
        xref_choice = forms.SelectFromList.show(
            xref_options,
            title="X origin for panel layout",
            button_name="Use",
            multiselect=False
        )

        if xref_choice == xref_options[1]:
            X_REF_OVERRIDE = "start"
            print("X_REF_OVERRIDE: start (measure x_in from wall START)")
        elif xref_choice == xref_options[2]:
            X_REF_OVERRIDE = "end"
            print("X_REF_OVERRIDE: end (measure x_in from wall END)")
        else:
            X_REF_OVERRIDE = None
            print("X_REF_OVERRIDE: None (use per-panel x_ref / default)")

        # --- UI: choose rotation behavior for this run ---
        rot_options = [
            "Use CSV rotation_deg",
            "Force 0° (no rotation)",
            "Force 90°",
            "Force -90°"
        ]
        rot_choice = forms.SelectFromList.show(
            rot_options,
            title="Panel rotation behavior",
            button_name="Use",
            multiselect=False
        )

        if rot_choice == rot_options[1]:
            ROTATION_OVERRIDE_DEG = 0.0
        elif rot_choice == rot_options[2]:
            ROTATION_OVERRIDE_DEG = 90.0
        elif rot_choice == rot_options[3]:
            ROTATION_OVERRIDE_DEG = -90.0
        else:
            ROTATION_OVERRIDE_DEG = None

        print("ROTATION_OVERRIDE_DEG: {0}".format(
            "CSV (per panel)" if ROTATION_OVERRIDE_DEG is None else ROTATION_OVERRIDE_DEG))
    else:
        # DirectShape mode: overrides don't matter
        X_REF_OVERRIDE = None
        ROTATION_OVERRIDE_DEG = None

    panels_by_wall = {}
    for panel in panels:
        panels_by_wall.setdefault(panel["wall_id"], []).append(panel)

    t = Transaction(doc, "Place Panels - Adaptive")
    t.Start()
    try:
        placed_panels = 0
        failed_panels = 0
        placed_cutouts = 0

        for wall_id, wall_panels in panels_by_wall.items():
            wall = get_wall_by_id(wall_id)
            if wall is None:
                print("[WARN] Wall {0} not found.".format(wall_id))
                failed_panels += len(wall_panels)
                continue

            wall_curve = wall.Location.Curve
            wall_length_in = wall_curve.Length * 12.0

            print("\n" + "=" * 70)
            print("WALL {0}: {1} panels".format(wall_id, len(wall_panels)))
            print("  Wall length: {0:.1f} in".format(wall_length_in))
            print("=" * 70)

            if TEST_MODE_LIMIT_PANELS > 0:
                wall_panels = wall_panels[:TEST_MODE_LIMIT_PANELS]
                print("  [TEST MODE] Limiting to first {0} panels".format(len(wall_panels)))

            for panel in wall_panels:
                print("\n  Placing '{0}':".format(panel.get('panel_name', 'Unnamed')))

                if use_directshape:
                    inst = create_panel_as_direct_shape(wall, panel)
                else:
                    inst = place_panel_family(wall, panel, family_symbol)

                if inst:
                    placed_panels += 1
                    print("    SUCCESS")
                else:
                    failed_panels += 1
                    print("    FAILED")

                if SHOW_CUTOUTS:
                    for cutout in panel['cutouts']:
                        cutout_vis = create_cutout_visualization(
                            wall,
                            panel,
                            cutout,
                            family_symbol,
                            use_directshape
                        )
                        if cutout_vis:
                            placed_cutouts += 1

        t.Commit()
        print("\n" + "=" * 70)
        print("PLACEMENT COMPLETE")
        print("=" * 70)
        print("Panels placed: {0}".format(placed_panels))
        print("Cutouts visualized: {0}".format(placed_cutouts))
        print("Panels failed: {0}".format(failed_panels))
        print("\nIf panels are still misaligned, try:")
        print(" - Adjusting FAMILY_ORIGIN_IS_CENTER")
        print(" - Setting x_ref ('start'/'end') and rotation_deg in the CSV")
        print(" - Changing X origin / rotation options in the dialog")
    except Exception as e:
        t.RollBack()
        print("[ERROR] Transaction failed: {0}".format(str(e)))
        import traceback
        print(traceback.format_exc())


if __name__ == "__main__":
    main()
