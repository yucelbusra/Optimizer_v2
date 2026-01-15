"""
Places largest possible wall panels on walls defined in walls.csv, 
considering openings from wall_openings.csv.
"""


from __future__ import print_function
import os
from datetime import datetime

# --- .NET UI imports ---
import clr
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from System.Windows.Forms import (
    Application, FolderBrowserDialog, DialogResult, Form,
    Label, RadioButton, Button, FormBorderStyle, FormStartPosition,
    MessageBox, MessageBoxButtons, MessageBoxIcon, TextBox, IWin32Window
)
from System.Drawing import Point, Size

# --- Additional refs to bind dialogs to Revit main window ---
clr.AddReference('System')
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from System import IntPtr
from Autodesk.Revit.UI import UIApplication

# --- Import optimizer module sitting next to this script ---
try:
    import panel_calculator as opt
except Exception as e:
    raise Exception("Failed to import panel_calculator.py: {0}".format(e))

# =========================
# Helpers: Revit window owner
# =========================
class WindowWrapper(IWin32Window):
    def __init__(self, handle):
        self._hwnd = handle
    @property
    def Handle(self):
        return self._hwnd

def get_revit_owner():
    """Return an IWin32Window wrapper of Revit's main window; None on failure."""
    try:
        # In pyRevit, __revit__ is already a UIApplication
        uiapp = __revit__
        hwnd = uiapp.MainWindowHandle
        return WindowWrapper(IntPtr(hwnd))
    except Exception as e:
        print("Warning: could not retrieve Revit main window handle: {0}".format(e))
        return None

# =========================
# UI: Folder Picker (click)
# =========================

def pick_data_folder():
    """Show a FolderBrowserDialog; default to Desktop, fallback to Home; return selected path or script dir."""
    owner = get_revit_owner()
    dialog = FolderBrowserDialog()
    dialog.Description = "Select folder containing walls.csv and wall_openings.csv"
    # Default: Desktop; fallback: Home
    initial_dir = os.path.join(os.path.expanduser("~"), "Desktop")
    if not os.path.exists(initial_dir):
        initial_dir = os.path.expanduser("~")
    dialog.SelectedPath = initial_dir
    result = dialog.ShowDialog(owner) if owner else dialog.ShowDialog()
    if result == DialogResult.OK and dialog.SelectedPath and os.path.isdir(dialog.SelectedPath):
        return dialog.SelectedPath
    else:
        # If user cancels, fall back to the folder where this script lives
        return os.path.dirname(os.path.abspath(__file__))

# =====================================
# UI: Orientation Picker (radio buttons)
# =====================================

class OrientationDialog(Form):
    def __init__(self):
        # Basic form setup
        self.Text = "Select Panel Orientation"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(600, 320)
        self.TopMost = True

        # Label
        lbl = Label()
        lbl.Text = "Select panel orientation:"
        lbl.Location = Point(12, 12)
        lbl.AutoSize = True
        self.Controls.Add(lbl)

        # Radio buttons (default: Vertical)
        self.rbVertical = RadioButton()
        self.rbVertical.Text = "Vertical"
        self.rbVertical.Checked = True
        self.rbVertical.Location = Point(24, 40)
        self.Controls.Add(self.rbVertical)

        self.rbHorizontal = RadioButton()
        self.rbHorizontal.Text = "Horizontal"
        self.rbHorizontal.Location = Point(24, 66)
        self.Controls.Add(self.rbHorizontal)

        # OK / Cancel buttons
        btnOK = Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = Point(180, 130)

        btnCancel = Button()
        btnCancel.Text = "Cancel"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = Point(260, 130)

        self.AcceptButton = btnOK
        self.CancelButton = btnCancel
        self.Controls.Add(btnOK)
        self.Controls.Add(btnCancel)


def pick_orientation():
    """Show OrientationDialog and return 'vertical' or 'horizontal'; None if canceled."""
    owner = get_revit_owner()
    dlg = OrientationDialog()
    result = dlg.ShowDialog(owner) if owner else dlg.ShowDialog()
    if result == DialogResult.OK:
        if dlg.rbVertical.Checked:
            return "vertical"
        else:
            return "horizontal"
    else:
        return None

# =====================================
# NEW: UI: Panel Type Picker
# =====================================

class PanelTypeDialog(Form):
    def __init__(self):
        # Basic form setup
        self.Text = "Select Panel Type"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(600, 320)
        self.TopMost = True

        # Label
        lbl = Label()
        lbl.Text = "Select panel type:"
        lbl.Location = Point(12, 12)
        lbl.AutoSize = True
        self.Controls.Add(lbl)

        # Radio buttons (default: Backer)
        self.rbBacker = RadioButton()
        self.rbBacker.Text = "Backer Panels (1/8\" spacing)"
        self.rbBacker.Checked = True
        self.rbBacker.Location = Point(24, 40)
        self.rbBacker.AutoSize = True
        self.Controls.Add(self.rbBacker)

        self.rbFullyFinished = RadioButton()
        self.rbFullyFinished.Text = "Fully Finished Panels (3/4\" spacing)"
        self.rbFullyFinished.Location = Point(24, 66)
        self.rbFullyFinished.AutoSize = True
        self.Controls.Add(self.rbFullyFinished)

        # OK / Cancel buttons
        btnOK = Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = Point(180, 130)

        btnCancel = Button()
        btnCancel.Text = "Cancel"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = Point(260, 130)

        self.AcceptButton = btnOK
        self.CancelButton = btnCancel
        self.Controls.Add(btnOK)
        self.Controls.Add(btnCancel)


def pick_panel_type():
    """Show PanelTypeDialog and return spacing value (0.125 or 0.75); None if canceled."""
    owner = get_revit_owner()
    dlg = PanelTypeDialog()
    result = dlg.ShowDialog(owner) if owner else dlg.ShowDialog()
    if result == DialogResult.OK:
        if dlg.rbBacker.Checked:
            return 0.125  # 1/8" for backer panels
        else:
            return 0.75   # 3/4" for fully finished panels
    else:
        return None

# =====================================
# UI: Project Name Input Dialog
# =====================================

class ProjectNameDialog(Form):
    def __init__(self):
        # Basic form setup
        self.Text = "Project Name"
        self.StartPosition = FormStartPosition.CenterScreen
        self.FormBorderStyle = FormBorderStyle.FixedDialog
        self.MaximizeBox = False
        self.MinimizeBox = False
        self.ClientSize = Size(600, 320)
        self.TopMost = True

        # Label
        lbl = Label()
        lbl.Text = "Enter project name for output folder:"
        lbl.Location = Point(12, 12)
        lbl.AutoSize = True
        self.Controls.Add(lbl)

        # TextBox for project name
        self.txtProjectName = TextBox()
        self.txtProjectName.Text = "PanelOptimization"
        self.txtProjectName.Location = Point(12, 40)
        self.txtProjectName.Size = Size(365, 20)
        self.Controls.Add(self.txtProjectName)

        # OK / Cancel buttons
        btnOK = Button()
        btnOK.Text = "OK"
        btnOK.DialogResult = DialogResult.OK
        btnOK.Location = Point(210, 80)

        btnCancel = Button()
        btnCancel.Text = "Cancel"
        btnCancel.DialogResult = DialogResult.Cancel
        btnCancel.Location = Point(290, 80)

        self.AcceptButton = btnOK
        self.CancelButton = btnCancel
        self.Controls.Add(btnOK)
        self.Controls.Add(btnCancel)


def get_project_name():
    """Show ProjectNameDialog and return project name; None if canceled."""
    owner = get_revit_owner()
    while True:
        dlg = ProjectNameDialog()
        result = dlg.ShowDialog(owner) if owner else dlg.ShowDialog()
        if result == DialogResult.OK:
            project_name = dlg.txtProjectName.Text.strip()
            if project_name:
                return project_name
            else:
                # Show error and loop to ask again
                MessageBox.Show(
                    "Project name cannot be empty.",
                    "Invalid Input",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                )
        else:
            return None

# =================
# Utility functions
# =================

def _ensure_dir(path):
    if not os.path.exists(path):
        try:
            os.makedirs(path)
        except Exception as e:
            print("Failed to create directory {0}: {1}".format(path, e))

def _backup_file(path):
    if os.path.exists(path):
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup = path.replace(".json", "_backup_{0}.json".format(ts))
        try:
            import shutil
            shutil.copy(path, backup)
        except Exception as e:
            print("Failed to backup file {0}: {1}".format(path, e))

def _sanitize_folder_name(name):
    """Remove invalid characters from folder name."""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        name = name.replace(char, '_')
    return name

# =====
# Main
# =====
def main():
    try:
        Application.EnableVisualStyles()
    except Exception as e:
        print("EnableVisualStyles failed: {0}".format(e))

    # 1) Pick input folder (click)
    input_dir = pick_data_folder()
    _ensure_dir(input_dir)

    # 2) Get project name (type)
    project_name = get_project_name()
    if project_name is None:
        MessageBox.Show(
            "Operation canceled.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        return

    # 3) Pick orientation (click)
    orientation = pick_orientation()
    if orientation is None:
        MessageBox.Show(
            "Operation canceled.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        return

    # 4) NEW: Pick panel type to determine spacing
    panel_spacing = pick_panel_type()
    if panel_spacing is None:
        MessageBox.Show(
            "Operation canceled.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        return

    # 5) Create timestamped output directory
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_project_name = _sanitize_folder_name(project_name)
    output_folder_name = "{0}_{1}".format(safe_project_name, timestamp)
    output_dir = os.path.join(input_dir, output_folder_name)
    _ensure_dir(output_dir)

    # 6) Build config based on orientation choice
    presets = opt.get_preset_configs()
    if orientation == "custom":
        MessageBox.Show(
            "Custom configuration selected.\n\n" +
            "You'll be asked to configure parameters in the console window.",
            "Custom Configuration",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
        config = opt.create_custom_config()
        config.project_name = project_name
    else:
        config = presets[orientation]
        config.project_name = project_name
    
    # NEW: Set panel spacing based on user selection
    config.panel_constraints.panel_spacing = panel_spacing
    panel_type_name = "Backer" if panel_spacing == 0.125 else "Fully Finished"

    # Show parameters in console and ask for confirmation
    print("\n" + "=" * 70)
    print("SELECTED CONFIGURATION: {} ({})".format(project_name, orientation.upper()))
    print("Panel Type: {} ({}\")".format(panel_type_name, panel_spacing))
    print("=" * 70)
    opt.print_config_summary(config)
    try:
        user_input = raw_input  # IronPython
    except NameError:
        user_input = input  # CPython
    print("\nOptions:")
    print("1. Continue with these settings")
    print("2. Edit parameters")
    print("3. Cancel")
    choice = user_input("\nChoice (1-3) [1]: ").strip() or "1"
    if choice == "2":
        config = opt.edit_panel_constraints(config)
        config.project_name = project_name  # Restore project name
        config.panel_constraints.panel_spacing = panel_spacing  # Restore spacing
    elif choice == "3":
        print("Cancelled by user.")
        return

    # 7) Resolve CSV paths (in INPUT directory)
    walls_csv = os.path.join(input_dir, "walls.csv")
    openings_csv = os.path.join(input_dir, "wall_openings.csv")

    # Validate inputs with message boxes
    if not os.path.exists(walls_csv):
        MessageBox.Show(
            "Could not find walls.csv in:\n{0}".format(input_dir),
            "Missing Input",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        )
        return

    if not os.path.exists(openings_csv):
        MessageBox.Show(
            "Could not find wall_openings.csv in:\n{0}\n\nProceeding without openings.".format(input_dir),
            "Missing Input",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )

    # 8) Run optimizer (output to OUTPUT directory)
    walls_rows = opt.load_walls_from_csv(walls_csv)
    openings_rows = opt.load_openings_from_csv(openings_csv)
    
    # UPDATED: process_all_walls now returns (panels_path, config_path)
    panels_path, config_path = opt.process_all_walls(
        walls_rows, openings_rows, output_dir,  # OUTPUT to timestamped folder
        config.door_clearances,
        config.window_clearances,
        config.storefront_clearances,
        config,  # Pass full config object
        orientation  # Pass orientation
    )

    # Config is automatically saved by process_all_walls
    if config_path:
        print("Configuration saved to: {}".format(config_path))
    else:
        print("WARNING: Configuration was not saved")

    # 9) Done message
    if panels_path and os.path.exists(panels_path):
        MessageBox.Show(
            "Optimization complete.\n\nPanel Type: {0}\nSpacing: {1}\"\n\nExported panels to:\n{2}".format(
                panel_type_name, panel_spacing, output_dir),
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information
        )
    else:
        MessageBox.Show(
            "No panels generated.\nPlease check inputs and configuration.",
            "Greedy Optimizer",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )

# Entrypoint for pyRevit button
if __name__ == "__main__":
    main()