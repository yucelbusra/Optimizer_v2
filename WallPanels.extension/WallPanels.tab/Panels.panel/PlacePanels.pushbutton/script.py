# -*- coding: utf-8 -*-
"""
REVIT PANEL PLACEMENT - STRUCTURAL CORE CENTER ALIGNMENT
Reads the optimized panel placement CSV and places panels in Revit.

FIXES:
- Calculates the CENTER of the Structural Core layer.
- Aligns the Panel's Center to the Core's Center.
- Ensures correct placement regardless of Wall Location Line (Finish Face, Centerline, etc).
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
ALLOW_TYPE_PARAM_CHANGE = True

# --- DEPTH SETTINGS ---
PANEL_THICKNESS_IN = 4.0
FAMILY_ORIGIN_LOCATION = "Center" 
MANUAL_DEPTH_OFFSET_IN = 0.0

# --- COORDINATE SETTINGS ---
PANEL_COORD_DEFAULT_REF = "start"
USE_CSV_ROTATION = True

# Runtime overrides
X_REF_OVERRIDE = None
ROTATION_OVERRIDE_DEG = None

WIDTH_PARAM_CANDIDATES = ["Width", "Panel Width", "W", "Overall Width", "Length", "L"]
HEIGHT_PARAM_CANDIDATES = ["Height", "Panel Height", "H", "Overall Height", "Thickness", "Depth"]

# Disable endcap extension to match exact drawing points
USE_WALL_ENDCAP_EXTENSION = False
PANEL_SIDE_SIGN = 1


# ========== UTILITIES ==========
def _pick_input_folder(default_dir=None):
    try:
        fbd = FolderBrowserDialog()
        fbd.Description = "Select the folder containing '{0}'".format(PANELS_FILE)
        if default_dir and os.path.isdir(default_dir):
            fbd.SelectedPath = default_dir
        result = fbd.ShowDialog()
        if result == DialogResult.OK and fbd.SelectedPath:
            return str(fbd.SelectedPath)
    except: pass
    return None

def norm_id(val):
    try: return str(int(float(val)))
    except: return str(val).strip()

def get_wall_by_id(wall_id):
    try:
        elem_id = int(float(wall_id))
        element = doc.GetElement(ElementId(elem_id))
        if isinstance(element, Wall): return element
    except: pass
    return None

def _feet(val_inch):
    return float(val_inch) / 12.0

# ========== GEOMETRY CORE ==========

def get_wall_geometry_normalized(wall):
    """
    Returns consistent geometry and the CENTER offset of the structural core.
    Fixes the 'Surface Placement' bug by calculating layer widths manually.
    """
    lc = wall.Location.Curve
    p0 = lc.GetEndPoint(0)
    p1 = lc.GetEndPoint(1)
    
    # 1. True Exterior Normal
    normal = wall.Orientation
    
    # 2. Determine "Visual Right" direction
    up = XYZ(0,0,1)
    visual_right_dir = normal.CrossProduct(up)
    
    # 3. Project p0/p1
    dot0 = p0.DotProduct(visual_right_dir)
    dot1 = p1.DotProduct(visual_right_dir)
    
    if dot0 < dot1:
        visual_left = p0
        visual_right = p1
    else:
        visual_left = p1
        visual_right = p0
        
    normalized_dir = (visual_right - visual_left).Normalize()
    
    # 4. Find CORE CENTER Offset correctly
    # Offset is defined as distance to move ALONG THE NORMAL to get to Core Center.
    core_center_offset = 0.0
    
    try:
        w_type = wall.WallType
        cs = w_type.GetCompoundStructure()
        
        if cs:
            # A. Calculate Core Center relative to Exterior Face
            total_width = cs.GetWidth()
            layers = cs.GetLayers()
            
            ext_thickness = 0.0
            core_thickness = 0.0
            
            # Iterate layers to find core boundaries
            # Layers are ordered Exterior -> Interior
            for i, layer in enumerate(layers):
                if cs.IsCoreLayer(i):
                    core_thickness += layer.Width
                elif cs.GetCoreBoundaryLayerIndex(0) > i: 
                    # This layer is before the core (Exterior side)
                    ext_thickness += layer.Width
            
            # Distance from Ext Face to Core Center
            core_center_from_ext_face = ext_thickness + (core_thickness / 2.0)
            
            # B. Determine where the Location Line is relative to Exterior Face
            # WALL_KEY_REF_PARAM values: 0=Wall Ctr, 2=Fin Face Ext, 3=Fin Face Int, etc.
            param = wall.get_Parameter(BuiltInParameter.WALL_KEY_REF_PARAM)
            loc_param_int = param.AsInteger() if param else 0
            
            loc_line_from_ext_face = 0.0
            
            if loc_param_int == 0: # Wall Centerline
                loc_line_from_ext_face = total_width / 2.0
            elif loc_param_int == 2: # Finish Face Exterior
                loc_line_from_ext_face = 0.0
            elif loc_param_int == 3: # Finish Face Interior
                loc_line_from_ext_face = total_width
            elif loc_param_int == 1: # Core Centerline
                loc_line_from_ext_face = core_center_from_ext_face
            elif loc_param_int == 4: # Core Face Exterior
                loc_line_from_ext_face = ext_thickness
            elif loc_param_int == 5: # Core Face Interior
                loc_line_from_ext_face = ext_thickness + core_thickness
                
            # C. Calculate Final Offset
            # If LocLine is at 0 (Ext) and Core is at 1, we need to move -1 (Inwards/Opposite to Normal).
            # Offset = LocLine - CorePosition
            core_center_offset = loc_line_from_ext_face - core_center_from_ext_face
            
    except Exception as e: 
        print("Warning - Core Calc Failed: {}".format(e))
        pass

    return visual_left, visual_right, normalized_dir, normal, core_center_offset


def compute_panel_base_point(wall, panel, extra_z_offset_in=0.0):
    """
    Calculates insertion point aligning Panel CENTER to Wall CORE CENTER.
    """
    
    vis_left, vis_right, wall_dir, wall_normal, core_center_off = get_wall_geometry_normalized(wall)
    
    # --- Endcap Extension (Disabled) ---
    left_used = vis_left
    right_used = vis_right
    if USE_WALL_ENDCAP_EXTENSION:
        half_thk = wall.Width / 2.0
        left_used = vis_left - (wall_dir * half_thk)
        right_used = vis_right + (wall_dir * half_thk)

    # --- Read Panel Data ---
    x_in = float(panel.get("x_in", 0.0) or 0.0)
    y_in = float(panel.get("y_in", 0.0) or 0.0)
    
    x_ref = (panel.get("x_ref", PANEL_COORD_DEFAULT_REF) or PANEL_COORD_DEFAULT_REF).lower().strip()
    if X_REF_OVERRIDE == "start": x_ref = "start"
    if X_REF_OVERRIDE == "end": x_ref = "end"

    x_ft = _feet(x_in)
    y_ft = _feet(y_in)

    # --- XY Location ---
    if x_ref == "start":
        pt_xy = left_used + (wall_dir * x_ft)
    else:
        pt_xy = right_used - (wall_dir * x_ft)

    # --- Z Location ---
    bb = wall.get_BoundingBox(None)
    base_z = bb.Min.Z + y_ft
    base_point_loc = XYZ(pt_xy.X, pt_xy.Y, base_z)
    
    # --- DEPTH ALIGNMENT LOGIC (CORE CENTER) ---
    # Goal: Panel Center = Core Center
    
    calculated_offset = core_center_off
    p_thickness_ft = _feet(PANEL_THICKNESS_IN)
    
    # Adjust based on where the Family Origin is defined
    if FAMILY_ORIGIN_LOCATION.lower() == "center":
        # Origin is Center. Core is Center. No adjustment needed.
        pass
    elif FAMILY_ORIGIN_LOCATION.lower() == "front":
        # Origin is Front. We want Center at Core Center.
        # So we move Origin BACK by half thickness.
        calculated_offset -= (p_thickness_ft / 2.0)
    elif FAMILY_ORIGIN_LOCATION.lower() == "back":
        # Origin is Back. We want Center at Core Center.
        # So we move Origin FORWARD by half thickness.
        calculated_offset += (p_thickness_ft / 2.0)

    # Manual Nudge + Visual Pop-out
    calculated_offset += _feet(MANUAL_DEPTH_OFFSET_IN)
    calculated_offset += _feet(extra_z_offset_in)

    # Final Point
    final_point = base_point_loc + (wall_normal * calculated_offset)
    
    return final_point, wall_dir, wall_normal


# ========== PLACEMENT ==========
def place_panel_family(wall, panel, symbol, extra_z_offset_in=0.0):
    if not ensure_symbol_active(symbol): return None
    
    try:
        pt, w_dir, w_norm = compute_panel_base_point(wall, panel, extra_z_offset_in)
    except Exception as e:
        print("[ERROR] Geometry calc failed: {0}".format(e))
        return None

    inst = None
    
    # 1. Place Instance
    try:
        inst = doc.Create.NewFamilyInstance(pt, symbol, wall, StructuralType.NonStructural)
        if extra_z_offset_in > 0:
            print("  [CUTOUT] Placed visualization pop-out.")
        else:
            print("  [PLACE] Hosted: {0}".format(panel.get("panel_name", "")))
    except: pass
        
    if not inst:
        try:
            lvl = get_wall_base_level(wall)
            if lvl:
                inst = doc.Create.NewFamilyInstance(pt, symbol, lvl, StructuralType.NonStructural)
            else:
                inst = doc.Create.NewFamilyInstance(pt, symbol, StructuralType.NonStructural)
            print("  [PLACE] Non-hosted: {0}".format(panel.get("panel_name", "")))
        except Exception as e:
            print("[ERROR] Placement failed: {0}".format(e))
            return None

    doc.Regenerate()
    
    # 2. ORIENTATION LOGIC
    # We want the panel facing EXTERIOR (Positive Dot Product)
    if inst and inst.CanFlipFacing:
        try:
            inst_facing = inst.FacingOrientation
            dot = w_norm.DotProduct(inst_facing)
            if dot < -0.01:
                inst.flipFacing()
        except: pass

    # 3. Rotation
    rot_deg = 0.0
    if ROTATION_OVERRIDE_DEG is not None:
        rot_deg = ROTATION_OVERRIDE_DEG
    elif USE_CSV_ROTATION:
        try: rot_deg = float(panel.get("rotation_deg", 0.0) or 0.0)
        except: pass

    if abs(rot_deg) > 0.001:
        try:
            axis = Line.CreateBound(pt, pt + XYZ(0,0,10))
            ElementTransformUtils.RotateElement(doc, inst.Id, axis, math.radians(rot_deg))
        except: pass

    # 4. Params
    set_size_parameters(inst, panel["width_in"], panel["height_in"], symbol)
    try:
        p = _find_param_by_candidates(inst, ["Name", "Panel Name", "Mark"])
        if p and not p.IsReadOnly: p.Set(panel.get("panel_name",""))
    except: pass
    
    return inst

# [Standard Helpers]
def get_element_name(element):
    try:
        p = element.get_Parameter(BuiltInParameter.SYMBOL_NAME)
        if p: return p.AsString()
    except: pass
    try: return element.Name
    except: return "Unknown"

def get_family_name(symbol):
    try: return symbol.Family.Name
    except: return "Unknown"

def get_all_family_symbols():
    collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
    families_dict = {}
    for symbol in collector:
        family_name = get_family_name(symbol)
        families_dict.setdefault(family_name, []).append(symbol)
    return families_dict

def get_panel_family_symbol(family_name):
    from pyrevit import forms
    if family_name:
        collector = FilteredElementCollector(doc).OfClass(FamilySymbol)
        for s in collector:
            if get_family_name(s) == family_name: return s, False
    families_dict = get_all_family_symbols()
    family_names = sorted(families_dict.keys())
    family_names.insert(0, "< Use DirectShape (3D Solid Panels) >")
    selected_family = forms.SelectFromList.show(family_names, title="Select Panel Placement Method", button_name="Select", multiselect=False)
    if not selected_family: return None, False
    if selected_family == "< Use DirectShape (3D Solid Panels) >": return None, True
    symbols = families_dict[selected_family]
    if len(symbols) == 1: return symbols[0], False
    symbol_names = [get_element_name(s) for s in symbols]
    selected_type = forms.SelectFromList.show(symbol_names, title="Select Family Type", button_name="Select", multiselect=False)
    if not selected_type: return symbols[0], False
    for symbol in symbols:
        if get_element_name(symbol) == selected_type: return symbol, False
    return symbols[0], False

def ensure_symbol_active(symbol):
    try:
        if not symbol.IsActive: symbol.Activate()
        return True
    except: return False

def _find_param_by_candidates(element, candidates):
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(nm.lower() == cand.lower() for cand in candidates): return p
        except: continue
    lower_cands = [c.lower() for c in candidates]
    for p in element.Parameters:
        try:
            nm = p.Definition.Name
            if nm and any(c in nm.lower() for c in lower_cands): return p
        except: continue
    return None

def set_size_parameters(inst, width_in, height_in, symbol=None):
    width_ft = _feet(width_in)
    height_ft = _feet(height_in)
    changed = False
    w_param = _find_param_by_candidates(inst, WIDTH_PARAM_CANDIDATES)
    h_param = _find_param_by_candidates(inst, HEIGHT_PARAM_CANDIDATES)
    try:
        if w_param and not w_param.IsReadOnly:
            w_param.Set(width_ft)
            changed = True
        if h_param and not h_param.IsReadOnly:
            h_param.Set(height_ft)
            changed = True
    except: pass
    if not changed and ALLOW_TYPE_PARAM_CHANGE and symbol:
        try:
            wtp = _find_param_by_candidates(symbol, WIDTH_PARAM_CANDIDATES)
            htp = _find_param_by_candidates(symbol, HEIGHT_PARAM_CANDIDATES)
            if wtp and not wtp.IsReadOnly: wtp.Set(width_ft)
            if htp and not htp.IsReadOnly: htp.Set(height_ft)
        except: pass

def get_wall_base_level(wall):
    try:
        lvl_id = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsElementId()
        if lvl_id and lvl_id.IntegerValue > 0: return doc.GetElement(lvl_id)
    except: pass
    try: return FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).FirstElement()
    except: return None

def create_panel_as_direct_shape(wall, panel):
    try:
        pt, w_dir, w_norm = compute_panel_base_point(wall, panel)
        w_ft = _feet(panel.get("width_in",0))
        h_ft = _feet(panel.get("height_in",0))
        thk = 1.0/12.0
        v1 = pt + (w_norm * 0.01)
        v2 = v1 + (w_dir * w_ft)
        v3 = v2 + XYZ(0,0,h_ft)
        v4 = v1 + XYZ(0,0,h_ft)
        v5 = pt + (w_norm * (0.01+thk))
        v6 = v5 + (w_dir * w_ft)
        v7 = v6 + XYZ(0,0,h_ft)
        v8 = v5 + XYZ(0,0,h_ft)
        lines = [Line.CreateBound(v1,v2), Line.CreateBound(v2,v3), Line.CreateBound(v3,v4), Line.CreateBound(v4,v1),
                 Line.CreateBound(v5,v6), Line.CreateBound(v6,v7), Line.CreateBound(v7,v8), Line.CreateBound(v8,v5),
                 Line.CreateBound(v1,v5), Line.CreateBound(v2,v6), Line.CreateBound(v3,v7), Line.CreateBound(v4,v8)]
        ds = DirectShape.CreateElement(doc, ElementId(int(BuiltInCategory.OST_GenericModel)))
        ds.SetShape(lines)
        ds.Name = panel.get("panel_name", "PanelSolid")
        print("  [DS] Created: {0}".format(ds.Name))
        return ds
    except Exception as e:
        print("DS Fail: {0}".format(e))
        return None

def create_cutout_visualization(wall, panel, cutout_data, symbol, use_ds):
    if not use_ds and symbol:
        try:
            c_x = float(cutout_data.get("x_in",0))
            c_y = float(cutout_data.get("y_in",0))
            g_x = float(panel.get("x_in",0)) + c_x
            g_y = float(panel.get("y_in",0)) + c_y
            fake_panel = panel.copy()
            fake_panel.update({
                "panel_name": "CUT_" + str(cutout_data.get("id","")),
                "x_in": g_x, "y_in": g_y,
                "width_in": cutout_data.get("width_in",0),
                "height_in": cutout_data.get("height_in",0),
                "cutouts": []
            })
            
            # [FIX] Visual pop-out for cutouts
            place_panel_family(wall, fake_panel, symbol, extra_z_offset_in=2.0)
            return True
        except: pass
    return False

# ========== MAIN ==========
def main():
    print("--- PANEL PLACEMENT: CORE CENTER ALIGNMENT ---")
    
    if USE_FOLDER_PICKER:
        path = _pick_input_folder(DEFAULT_INPUT_DIR)
        if not path: return
        panels_path = os.path.join(path, PANELS_FILE)
    else:
        path = DEFAULT_INPUT_DIR or os.getcwd()
        panels_path = os.path.join(path, PANELS_FILE)

    if not os.path.exists(panels_path):
        print("CSV not found: " + panels_path)
        return

    panels = []
    with open(panels_path, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            try: cutouts = json.loads(row.get("cutouts_json","[]"))
            except: cutouts = []
            p = {
                "wall_id": norm_id(row.get("wall_id")),
                "x_in": row.get("x_in"),
                "y_in": row.get("y_in"),
                "width_in": row.get("width_in"),
                "height_in": row.get("height_in"),
                "x_ref": row.get("x_ref"),
                "panel_name": row.get("panel_name"),
                "rotation_deg": row.get("rotation_deg"),
                "cutouts": cutouts
            }
            panels.append(p)

    print("Loaded {0} panels.".format(len(panels)))
    
    sym, use_ds = get_panel_family_symbol(PANEL_FAMILY_NAME)
    if not sym and not use_ds: return

    from pyrevit import forms
    global X_REF_OVERRIDE, ROTATION_OVERRIDE_DEG
    
    if not use_ds:
        print("Using Family: " + get_family_name(sym))
        
        xref_ops = ["Use CSV Default", "Force Start (Left)", "Force End (Right)"]
        res = forms.SelectFromList.show(xref_ops, button_name="Set X Ref", multiselect=False)
        if res == xref_ops[1]: X_REF_OVERRIDE = "start"
        elif res == xref_ops[2]: X_REF_OVERRIDE = "end"
        
        rot_ops = ["Use CSV Rotation", "Force 0", "Force 90", "Force -90", "Force 180"]
        res = forms.SelectFromList.show(rot_ops, button_name="Set Rotation", multiselect=False)
        if res == rot_ops[1]: ROTATION_OVERRIDE_DEG = 0.0
        elif res == rot_ops[2]: ROTATION_OVERRIDE_DEG = 90.0
        elif res == rot_ops[3]: ROTATION_OVERRIDE_DEG = -90.0
        elif res == rot_ops[4]: ROTATION_OVERRIDE_DEG = 180.0

    # Group by wall
    panels_map = {}
    for p in panels:
        panels_map.setdefault(p["wall_id"], []).append(p)

    t = Transaction(doc, "Place Panels")
    t.Start()
    
    count = 0
    for wid, wall_panels in panels_map.items():
        wall = get_wall_by_id(wid)
        if not wall:
            print("Wall {0} not found.".format(wid))
            continue
            
        print("\n--- Wall {0} ---".format(wid))
        for p in wall_panels:
            if use_ds:
                res = create_panel_as_direct_shape(wall, p)
            else:
                res = place_panel_family(wall, p, sym)
            if res: count += 1
            if SHOW_CUTOUTS:
                for c in p["cutouts"]:
                    create_cutout_visualization(wall, p, c, sym, use_ds)
            
    t.Commit()
    print("\nDone. Placed {0} panels.".format(count))

if __name__ == "__main__":
    main()