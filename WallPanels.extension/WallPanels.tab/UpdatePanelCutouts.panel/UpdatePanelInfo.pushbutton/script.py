# -*- coding: utf-8 -*-
"""
UPDATE PANEL INFO - SELECT ONE PANEL, EXPORT ALL ON SAME WALL
User selects one panel, script finds all panels and exports their current state
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector, Wall, Transaction, XYZ, Line,
    FamilySymbol, BuiltInCategory, BuiltInParameter, ElementId,
    DirectShape, Options, GeometryElement
)
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons
from pyrevit import revit, forms
import csv
import os
import json

import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import SaveFileDialog, DialogResult

doc = revit.doc
uidoc = revit.uidoc

# ========== SETTINGS ==========
# Default location for saving
DEFAULT_OUTPUT_DIR = r"C:\Users\byucel\OneDrive - RNGD\Desktop\revit_walls"


def _pick_save_csv_path(default_dir, default_name):
    """
    Ask user for an output CSV path (folder + file name) via Windows file dialog.
    Returns full path or None if cancelled.
    """
    try:
        dlg = SaveFileDialog()
        dlg.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        dlg.Title = "Save updated panel information"
        dlg.DefaultExt = "csv"
        dlg.AddExtension = True
        
        # Set initial directory
        if default_dir and os.path.isdir(default_dir):
            dlg.InitialDirectory = default_dir
        else:
            dlg.InitialDirectory = os.path.expanduser("~\\Desktop")

        # Set default filename
        if default_name:
            dlg.FileName = default_name
        else:
            dlg.FileName = "updated_panel_info.csv"

        result = dlg.ShowDialog()
        
        if result == DialogResult.OK and dlg.FileName:
            selected_path = str(dlg.FileName)
            print("User selected path: {}".format(selected_path))
            return selected_path
        else:
            print("User cancelled file selection")
            return None
            
    except Exception as e:
        print("SaveFileDialog error: {0}".format(e))
        import traceback
        traceback.print_exc()
        forms.alert("Error opening file dialog: {}".format(str(e)))
        return None


def get_host_wall(element):
    """Try to find the wall that hosts this element."""
    try:
        # Try to get host
        if hasattr(element, 'Host'):
            host = element.Host
            if isinstance(host, Wall):
                return host
        
        # Alternative: check bounding box and find nearest wall
        bbox = element.get_BoundingBox(None)
        if bbox:
            center = (bbox.Min + bbox.Max) * 0.5
            
            # Find walls near this point
            walls = FilteredElementCollector(doc).OfClass(Wall).ToElements()
            min_dist = float('inf')
            nearest_wall = None
            
            for wall in walls:
                loc = wall.Location
                if hasattr(loc, 'Curve'):
                    curve = loc.Curve
                    # Get distance from center to wall curve
                    result = curve.Project(center)
                    if result:
                        dist = result.Distance
                        if dist < min_dist:
                            min_dist = dist
                            nearest_wall = wall
            
            if min_dist < 2.0:  # Within 2 feet
                return nearest_wall
        
        return None
        
    except Exception as e:
        print("Error finding host wall: {}".format(str(e)))
        return None


def get_all_directshapes_near_wall(wall):
    """Get all DirectShape elements near the specified wall."""
    if not wall:
        # If no wall, get all DirectShapes
        collector = FilteredElementCollector(doc).OfClass(DirectShape)
        return [ds for ds in collector if ds.Name and not ds.Name.startswith("Cutout_")]
    
    try:
        # Get wall location
        loc_curve = wall.Location.Curve
        wall_start = loc_curve.GetEndPoint(0)
        wall_end = loc_curve.GetEndPoint(1)
        
        # Get all DirectShapes
        collector = FilteredElementCollector(doc).OfClass(DirectShape)
        nearby_panels = []
        
        for ds in collector:
            name = ds.Name
            # Skip cutouts
            if not name or name.startswith("Cutout_"):
                continue
            
            # Check if near the wall
            bbox = ds.get_BoundingBox(None)
            if bbox:
                center = (bbox.Min + bbox.Max) * 0.5
                result = loc_curve.Project(center)
                if result and result.Distance < 2.0:  # Within 2 feet
                    nearby_panels.append(ds)
        
        return nearby_panels
        
    except Exception as e:
        print("Error finding panels near wall: {}".format(str(e)))
        # Fallback: return all DirectShapes
        collector = FilteredElementCollector(doc).OfClass(DirectShape)
        return [ds for ds in collector if ds.Name and not ds.Name.startswith("Cutout_")]


def extract_panel_info_from_element(panel_element, wall=None):
    """Extract current dimensions and position from a placed panel."""
    try:
        # Get the geometry
        options = Options()
        geom = panel_element.get_Geometry(options)
        
        if not geom:
            return None
        
        # Get bounding box to determine dimensions
        bbox = panel_element.get_BoundingBox(None)
        if not bbox:
            return None
        
        min_pt = bbox.Min
        max_pt = bbox.Max
        
        # Calculate dimensions
        width_ft = abs(max_pt.X - min_pt.X)
        height_ft = abs(max_pt.Z - min_pt.Z)
        depth_ft = abs(max_pt.Y - min_pt.Y)
        
        # Use the two largest dimensions
        dims = sorted([width_ft, height_ft, depth_ft], reverse=True)
        width_ft = dims[0]
        height_ft = dims[1]
        
        width_in = width_ft * 12.0
        height_in = height_ft * 12.0
        
        # Position
        x_in = min_pt.X * 12.0
        y_in = min_pt.Z * 12.0
        
        # Wall info
        wall_id = ""
        if wall:
            wall_id = str(wall.Id.IntegerValue)
        
        panel_data = {
            "panel_name": panel_element.Name or "Panel",
            "panel_type": "{}x{}".format(int(round(width_in)), int(round(height_in))),
            "wall_id": wall_id,
            "element_id": str(panel_element.Id.IntegerValue),
            "x_in": round(x_in, 2),
            "y_in": round(y_in, 2),
            "width_in": round(width_in, 2),
            "height_in": round(height_in, 2),
            "area_in2": round(width_in * height_in, 2),
        }
        
        return panel_data
        
    except Exception as e:
        print("Error extracting panel info from {}: {}".format(
            panel_element.Id.IntegerValue, str(e)))
        return None


def export_panels_from_selection():
    """Let user select one panel, then export all panels on same wall."""
    
    # Instruction message
    forms.alert(
        "SELECT ONE PANEL\n\n"
        "Click on any panel to select it.\n"
        "All panels on the same wall will be exported."
    )
    
    # Ask user to select one panel
    try:
        import Autodesk
        selection = uidoc.Selection.PickObject(
            Autodesk.Revit.UI.Selection.ObjectType.Element,
            "Select any panel"
        )
        
        selected_elem = doc.GetElement(selection.ElementId)
        
        if not selected_elem:
            forms.alert("No element selected!")
            return
        
        print("Selected element: {} (ID: {})".format(
            selected_elem.Name, selected_elem.Id.IntegerValue))
        
        # Find the wall this panel is on
        host_wall = get_host_wall(selected_elem)
        
        if host_wall:
            print("Found host wall: {} (ID: {})".format(
                host_wall.Name, host_wall.Id.IntegerValue))
        else:
            print("No host wall found, will export all panels")
        
        # Get all panels on this wall (or all panels if no wall found)
        all_panels = get_all_directshapes_near_wall(host_wall)
        
        print("Found {} total panels".format(len(all_panels)))
        
        if not all_panels:
            forms.alert("No panels found!")
            return
        
        # Show confirmation
        wall_info = "on wall {}".format(host_wall.Id.IntegerValue) if host_wall else "in the model"
        confirm_msg = "Found {} panels {}.\n\nExport all to CSV?".format(
            len(all_panels), wall_info)
        
        if not forms.alert(confirm_msg, yes=True, no=True):
            print("Export cancelled by user")
            return
        
        # Extract info from each panel
        all_rows = []
        for panel in all_panels:
            panel_data = extract_panel_info_from_element(panel, host_wall)
            if panel_data:
                all_rows.append(panel_data)
        
        if not all_rows:
            forms.alert("Could not extract information from panels!")
            return
        
        print("Successfully extracted info from {} panels".format(len(all_rows)))
        
        # Ask user where to save
        output_path = _pick_save_csv_path(DEFAULT_OUTPUT_DIR, "updated_panel_info.csv")
        
        if not output_path:
            forms.alert("Export cancelled.")
            return
        
        # Define CSV columns
        fieldnames = [
            "panel_name", "panel_type", "wall_id", "element_id",
            "x_in", "y_in", "width_in", "height_in", "area_in2"
        ]
        
        # Write CSV
        try:
            with open(output_path, "w", newline='') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(all_rows)
            
            forms.alert(
                "✓ Successfully exported {} panels to:\n{}".format(
                    len(all_rows), output_path
                )
            )
            print("Successfully exported to: {}".format(output_path))
            
        except Exception as e:
            forms.alert("Error writing CSV file:\n{}".format(str(e)))
            print("Export error: {}".format(str(e)))
        
    except Exception as e:
        print("Selection cancelled or error: {}".format(str(e)))
        return


# ========== MAIN ==========

def main():
    print("\n" + "=" * 70)
    print("UPDATE PANEL INFO - SELECT & EXPORT")
    print("=" * 70 + "\n")
    
    export_panels_from_selection()
    
    print("\nExport complete!")


if __name__ == "__main__":
    main()