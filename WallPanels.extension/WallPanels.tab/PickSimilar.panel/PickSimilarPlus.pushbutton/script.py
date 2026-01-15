# -*- coding: utf-8 -*-
"""
PICK SIMILAR PLUS
-----------------
pyRevit tool to select elements similar to a picked element with options:
    a) Entire model
    b) Same axis (same X or Y)
    c) Same elevation (same Z)

"Similar" is defined as:
    - Same Category
    - Same Type (GetTypeId)

"Same axis":
    - Elements whose bounding box center has either:
        |X - X_seed| < AXIS_TOL  OR  |Y - Y_seed| < AXIS_TOL

"Same elevation":
    - |Z - Z_seed| < ELEV_TOL
"""

from pyrevit import revit, DB, forms

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId
)
from Autodesk.Revit.UI.Selection import ObjectType

# >>> ADD THIS IMPORT <<<
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

# ---------------------------------------
# CONFIG
# ---------------------------------------
# Internal units are feet. 1" ~ 0.0833 ft
AXIS_TOL = 0.1   # ~1.2" tolerance for same axis
ELEV_TOL = 0.1   # ~1.2" tolerance for same elevation


def get_element_center(elem):
    """Return the center (XYZ) of the element's bounding box in model coords."""
    try:
        bbox = elem.get_BoundingBox(None)
        if not bbox:
            return None
        min_pt = bbox.Min
        max_pt = bbox.Max
        return DB.XYZ(
            0.5 * (min_pt.X + max_pt.X),
            0.5 * (min_pt.Y + max_pt.Y),
            0.5 * (min_pt.Z + max_pt.Z)
        )
    except Exception:
        return None


def collect_similar_elements(seed):
    """Collect all elements in the model that are similar to the seed."""
    seed_cat = seed.Category
    seed_typeid = seed.GetTypeId()

    if seed_cat is None or seed_typeid == ElementId.InvalidElementId:
        return []

    collector = (FilteredElementCollector(doc)
                 .WhereElementIsNotElementType()
                 .OfCategoryId(seed_cat.Id))

    similar = []
    for e in collector:
        try:
            if e.Id == seed.Id:
                continue
            if e.GetTypeId() == seed_typeid:
                similar.append(e)
        except Exception:
            continue

    return similar


def select_elements(elem_list):
    """Set current selection to the provided elements."""
    # Need a .NET List[ElementId], not a Python list
    id_list = List[ElementId]()
    for e in elem_list:
        id_list.Add(e.Id)

    uidoc.Selection.SetElementIds(id_list)


def main():
    # ----------------------------
    # 1) Pick seed element
    # ----------------------------
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            "Pick an element to 'Pick Similar'"
        )
    except Exception:
        # User cancelled
        return

    seed = doc.GetElement(ref.ElementId)
    if seed is None:
        forms.alert("No element selected.", exitscript=True)

    seed_center = get_element_center(seed)
    if seed_center is None:
        forms.alert("Could not get bounding box for the selected element.", exitscript=True)

    # ----------------------------
    # 2) Choose mode
    # ----------------------------
    modes = [
        "Entire model",
        "Same axis (X/Y)",
        "Same elevation (Z)"
    ]

    mode = forms.CommandSwitchWindow.show(
        modes,
        message="Pick Similar Options"
    )

    if not mode:
        # User cancelled
        return

    # ----------------------------
    # 3) Collect similar elements
    # ----------------------------
    similar_elems = collect_similar_elements(seed)

    if not similar_elems:
        forms.alert("No similar elements found in the model.", exitscript=True)

    # ----------------------------
    # 4) Filter based on mode
    # ----------------------------
    if mode == "Entire model":
        # Include the seed as well to make it clear
        final_selection = [seed] + similar_elems

    else:
        filtered = []
        for e in similar_elems:
            c = get_element_center(e)
            if c is None:
                continue

            if mode == "Same axis (X/Y)":
                same_x = abs(c.X - seed_center.X) < AXIS_TOL
                same_y = abs(c.Y - seed_center.Y) < AXIS_TOL
                if same_x or same_y:
                    filtered.append(e)

            elif mode == "Same elevation (Z)":
                if abs(c.Z - seed_center.Z) < ELEV_TOL:
                    filtered.append(e)

        if not filtered:
            forms.alert("No similar elements found for the chosen condition.", exitscript=True)

        final_selection = [seed] + filtered

    # ----------------------------
    # 5) Apply selection
    # ----------------------------
    select_elements(final_selection)


if __name__ == "__main__":
    main()
