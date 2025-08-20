import sys
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from scipy.stats import pearsonr
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QListWidget,
    QFileDialog,
    QMessageBox,
    QCheckBox,
    QFormLayout,
    QLineEdit,
    QGroupBox,
    QMenuBar,
    QMenu,
    QTabWidget,
    QListWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
)
from PyQt6.QtCore import Qt, pyqtSignal, QSettings, QEvent
from PyQt6.QtGui import QFont, QAction, QMouseEvent

# Import custom modules
from modules.baseline import (
    baseline_correction,
    get_baseline_methods,
    get_baseline_with_raw,
)
from modules.plotting import (
    setup_originlab_style,
    format_originlab_plot,
    create_originlab_legend,
)
from modules.file_converter import (
    convert_jws_with_fallback,
    load_ylk_file,
    save_ylk_file,
    ylk_to_dataframe,
)
from modules.data_processing import preprocess_data, calculate_correlation_matrix


class BaselineCreationTab(QWidget):
    def __init__(self, ylk_data, filename, parent=None):
        super().__init__(parent)
        self.ylk_data = ylk_data.copy()
        self.filename = filename
        self.parent_analyzer = parent
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Info section
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel(f"File: {self.filename}"))
        range_info = self.ylk_data.get("range", [0, 0])
        info_layout.addWidget(
            QLabel(f"Range: {range_info[0]:.1f} - {range_info[1]:.1f} cm⁻¹")
        )
        layout.addLayout(info_layout)

        # Parameters section
        params_group = QGroupBox("ALS Parameters")
        params_layout = QFormLayout(params_group)

        # Load existing baseline parameters as defaults if they exist
        existing_params = self.ylk_data.get("metadata", {}).get("baseline_params", {})

        # Ensure lambda is positive and properly formatted
        existing_lambda = existing_params.get("lambda", 1e5)
        if existing_lambda <= 0:
            existing_lambda = 1e5
        default_lambda = f"{existing_lambda:g}"  # Use :g format to handle scientific notation properly

        # Ensure p is valid
        existing_p = existing_params.get("p", 0.01)
        if existing_p <= 0 or existing_p >= 1:
            existing_p = 0.01
        default_p = str(existing_p)

        default_smooth = existing_params.get("smooth", False)

        self.lambda_edit = QLineEdit(default_lambda)
        params_layout.addRow("Lambda (smoothness):", self.lambda_edit)

        self.p_edit = QLineEdit(default_p)
        params_layout.addRow("P (asymmetry):", self.p_edit)

        self.smooth_checkbox = QCheckBox("Apply smoothing")
        self.smooth_checkbox.setChecked(default_smooth)
        params_layout.addRow("", self.smooth_checkbox)

        layout.addWidget(params_group)

        # Plot section
        plot_group = QGroupBox("Baseline Preview")
        plot_layout = QVBoxLayout(plot_group)

        self.figure = Figure(figsize=(10, 6))
        self.canvas = FigureCanvas(self.figure)
        self.ax = self.figure.add_subplot(111)
        plot_layout.addWidget(self.canvas)

        # Enable right-click context menu on canvas
        self.canvas.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.canvas.customContextMenuRequested.connect(self.show_plot_context_menu)

        # Enable mouse events for draggable anchors
        self.canvas.mpl_connect("button_press_event", self.on_mouse_press)
        self.canvas.mpl_connect("button_release_event", self.on_mouse_release)
        self.canvas.mpl_connect("motion_notify_event", self.on_mouse_move)

        # Enable keyboard events for anchor deletion
        self.canvas.mpl_connect("key_press_event", self.on_key_press)

        # Make sure the canvas can receive keyboard focus
        self.canvas.setFocusPolicy(Qt.FocusPolicy.ClickFocus)

        # Store anchors for manual baseline adjustment
        self.anchors = []
        self.selected_anchor = None
        self.dragging = False

        # Load existing anchor points if they exist
        existing_anchors = existing_params.get("anchors", [])
        if existing_anchors:
            self.anchors = existing_anchors.copy()

        layout.addWidget(plot_group)

        # Toggle button for baseline/corrected view
        toggle_layout = QHBoxLayout()
        self.view_toggle_btn = QPushButton("Show Corrected Data")
        self.view_toggle_btn.setCheckable(True)
        self.view_toggle_btn.clicked.connect(self.toggle_view)
        toggle_layout.addWidget(self.view_toggle_btn)

        layout.addLayout(toggle_layout)

        # Control buttons
        button_layout = QHBoxLayout()

        save_btn = QPushButton("Save Baseline")
        save_btn.clicked.connect(self.save_baseline)
        button_layout.addWidget(save_btn)

        close_btn = QPushButton("Close Tab")
        close_btn.clicked.connect(self.close_tab)
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)

        # Connect parameter changes to real-time update
        self.lambda_edit.textChanged.connect(self.update_preview)
        self.p_edit.textChanged.connect(self.update_preview)
        self.smooth_checkbox.stateChanged.connect(self.update_preview)

        # Initial plot
        self.update_preview()

    def get_parameters(self):
        """Get ALS parameters from UI"""
        try:
            lambda_val = float(self.lambda_edit.text())
            p_val = float(self.p_edit.text())
            smooth_val = self.smooth_checkbox.isChecked()
            return lambda_val, p_val, smooth_val
        except ValueError:
            return 1e5, 0.01, False

    def on_mouse_press(self, event):
        """Handle mouse press events for anchor selection and manipulation"""
        if event.inaxes != self.ax:
            return

        # Only handle left-click for anchor selection (ignore right-click)
        if event.button != 1:  # 1 = left mouse button
            return

        # Check if we clicked near an anchor point
        if self.anchors:
            # Convert mouse coordinates to display coordinates
            mouse_display_x, mouse_display_y = self.ax.transData.transform(
                [event.xdata, event.ydata]
            )

            # Find the closest anchor within a certain pixel radius
            min_pixel_distance = float("inf")
            closest_anchor_idx = None

            for i, (anchor_x, anchor_y) in enumerate(self.anchors):
                # Convert anchor coordinates to display coordinates
                anchor_display_x, anchor_display_y = self.ax.transData.transform(
                    [anchor_x, anchor_y]
                )

                # Calculate pixel distance
                pixel_distance = (
                    (mouse_display_x - anchor_display_x) ** 2
                    + (mouse_display_y - anchor_display_y) ** 2
                ) ** 0.5
                if (
                    pixel_distance < min_pixel_distance and pixel_distance < 10
                ):  # 10 pixel threshold
                    min_pixel_distance = pixel_distance
                    closest_anchor_idx = i

            if closest_anchor_idx is not None:
                self.selected_anchor = closest_anchor_idx
                self.dragging = True
                self.update_preview()  # Redraw to show selection

    def on_mouse_release(self, event):
        """Handle mouse release events"""
        if self.dragging:
            self.dragging = False
            self.update_preview()  # Redraw to update anchor appearance

    def on_key_press(self, event):
        """Handle keyboard events"""
        if event.key == "delete" or event.key == "backspace":
            # Delete the selected anchor if any
            if self.selected_anchor is not None:
                self.remove_single_anchor(self.selected_anchor)
        elif event.key == "escape":
            # Deselect anchor on Escape
            self.selected_anchor = None
            self.dragging = False
            self.update_preview()

    def on_mouse_move(self, event):
        """Handle mouse move events for dragging anchors"""
        if event.inaxes != self.ax or not self.dragging or self.selected_anchor is None:
            return

        # Update the position of the selected anchor with boundary checks
        if event.xdata is not None and event.ydata is not None:
            # Use fixed axis limits to prevent auto-expansion
            if hasattr(self, "_fixed_xlim") and hasattr(self, "_fixed_ylim"):
                x_min, x_max = self._fixed_xlim
                y_min, y_max = self._fixed_ylim
            else:
                # Fallback to current limits if fixed limits not set
                x_min, x_max = self.ax.get_xlim()
                y_min, y_max = self.ax.get_ylim()

            # Clamp coordinates within fixed graph bounds
            constrained_x = max(x_min, min(x_max, event.xdata))
            constrained_y = max(y_min, min(y_max, event.ydata))

            self.anchors[self.selected_anchor] = (constrained_x, constrained_y)
            self.update_preview()  # Redraw with updated anchor position

    def show_plot_context_menu(self, position):
        """Show context menu on right-click on plot"""
        # Get the mouse position in data coordinates
        x_data, y_data = self.ax.transData.inverted().transform(
            [position.x(), position.y()]
        )

        # Create context menu
        menu = QMenu(self)

        # Add anchor action
        add_anchor_action = QAction("Add Anchor Point", self)
        add_anchor_action.triggered.connect(lambda: self.add_anchor(x_data, y_data))
        menu.addAction(add_anchor_action)

        # Clear anchors action (only show if there are anchors)
        if self.anchors:
            clear_anchors_action = QAction("Clear All Anchors", self)
            clear_anchors_action.triggered.connect(self.clear_anchors)
            menu.addAction(clear_anchors_action)

        # Add instructions for anchor deletion
        if self.anchors:
            menu.addSeparator()
            info_action = QAction(
                "Tip: Click anchor to select, then press Delete to remove", self
            )
            info_action.setEnabled(False)  # Make it non-clickable, just informational
            menu.addAction(info_action)

        menu.exec(self.canvas.mapToGlobal(position))

    def remove_single_anchor(self, anchor_idx):
        """Remove a single anchor by index"""
        if 0 <= anchor_idx < len(self.anchors):
            self.anchors.pop(anchor_idx)
            # Reset selected anchor if it was the removed one
            if self.selected_anchor == anchor_idx:
                self.selected_anchor = None
            elif self.selected_anchor is not None and self.selected_anchor > anchor_idx:
                # Adjust selected anchor index if it's after the removed anchor
                self.selected_anchor -= 1
            self.update_preview()

    def add_anchor(self, x, y):
        """Add an anchor point at the specified coordinates"""
        # If no existing anchors, try to snap to ALS baseline
        if not self.anchors:
            try:
                lambda_val, p_val, smooth_val = self.get_parameters()
                raw_data = self.ylk_data.get("raw_data", {})
                x_data = np.array(raw_data.get("x", []))
                y_data = np.array(raw_data.get("y", []))

                if len(x_data) > 0 and len(y_data) > 0:
                    from modules.baseline import get_baseline_with_raw

                    # Calculate baseline with smoothing if enabled
                    _, als_baseline, _ = get_baseline_with_raw(
                        x_data,
                        y_data,
                        method="als",
                        lam=lambda_val,
                        p=p_val,
                        smooth=smooth_val,
                    )

                    # Find closest x point and get corresponding baseline y
                    closest_idx = np.argmin(np.abs(x_data - x))
                    baseline_y = als_baseline[closest_idx]

                    # Snap to baseline y-coordinate
                    y = baseline_y
            except Exception:
                pass  # Use original y if baseline calculation fails

        self.anchors.append((x, y))
        self.update_preview()

    def clear_anchors(self):
        """Clear all anchor points"""
        self.anchors = []
        self.selected_anchor = None
        self.dragging = False
        self.update_preview()

    def draw_anchors(self):
        """Draw anchor points on the plot with selection feedback"""
        if self.anchors:
            for i, (anchor_x, anchor_y) in enumerate(self.anchors):
                # Draw anchor point
                color = "red"
                size = 50
                marker = "o"

                # Highlight selected anchor
                if i == self.selected_anchor:
                    color = "orange"
                    size = 60  # Smaller selected anchor size
                    marker = "s"  # Square for selected anchors

                self.ax.scatter(
                    anchor_x, anchor_y, color=color, s=size, marker=marker, zorder=5
                )

    def toggle_view(self):
        """Toggle between baseline and corrected view"""
        if self.view_toggle_btn.isChecked():
            self.view_toggle_btn.setText("Show Baseline View")
        else:
            self.view_toggle_btn.setText("Show Corrected Data")
        self.update_preview()

    def _apply_anchor_adjustments(self, x_data, als_baseline):
        """Apply anchor adjustments to ALS baseline with smooth transitions"""
        adjusted_baseline = als_baseline.copy()

        if not self.anchors:
            return adjusted_baseline

        # Calculate sigma based on data range / 50
        data_range = np.max(x_data) - np.min(x_data)
        sigma = data_range / 50.0  # Smaller influence area

        # For each anchor, apply a smooth adjustment
        for anchor_x, anchor_y in self.anchors:
            # Find the closest point in the data
            closest_idx = np.argmin(np.abs(x_data - anchor_x))
            closest_x = x_data[closest_idx]

            # Calculate the adjustment needed at this point
            current_baseline_y = als_baseline[closest_idx]
            adjustment = anchor_y - current_baseline_y

            # Apply smooth adjustment using a Gaussian-like function
            distances = np.abs(x_data - closest_x)
            weights = np.exp(-0.5 * (distances / sigma) ** 2)

            # Apply weighted adjustment
            adjusted_baseline += adjustment * weights

        return adjusted_baseline

    def update_preview(self):
        """Update preview with current parameters (real-time)"""
        lambda_val, p_val, smooth_val = self.get_parameters()

        raw_data = self.ylk_data.get("raw_data", {})
        x_data = np.array(raw_data.get("x", []))
        y_data = np.array(raw_data.get("y", []))

        if len(x_data) == 0 or len(y_data) == 0:
            # Clear plot and show error message
            self.ax.clear()
            self.ax.text(
                0.5,
                0.5,
                "No data available",
                transform=self.ax.transAxes,
                ha="center",
                va="center",
                fontsize=14,
            )
            self.ax.set_title("Error: No data to display")
            self.canvas.draw()
            return

        if len(x_data) != len(y_data):
            # Clear plot and show error message
            self.ax.clear()
            self.ax.text(
                0.5,
                0.5,
                "Data dimension mismatch",
                transform=self.ax.transAxes,
                ha="center",
                va="center",
                fontsize=14,
            )
            self.ax.set_title("Error: Invalid data dimensions")
            self.canvas.draw()
            return

        # Store current axis limits to prevent auto-expansion during anchor dragging
        if hasattr(self, "_fixed_xlim") and hasattr(self, "_fixed_ylim"):
            stored_xlim = self._fixed_xlim
            stored_ylim = self._fixed_ylim
        else:
            # Initial setup - use data range with some padding
            x_range = np.max(x_data) - np.min(x_data)
            y_range = np.max(y_data) - np.min(y_data)
            x_padding = x_range * 0.02
            y_padding = y_range * 0.1
            stored_xlim = (np.min(x_data) - x_padding, np.max(x_data) + x_padding)
            stored_ylim = (np.min(y_data) - y_padding, np.max(y_data) + y_padding)
            self._fixed_xlim = stored_xlim
            self._fixed_ylim = stored_ylim

        self.ax.clear()

        try:
            from modules.baseline import get_baseline_with_raw

            # Always use original raw data for calculation
            original_y_data = np.array(self.ylk_data["raw_data"]["y"])

            # Calculate baseline with smoothing applied if requested
            processed_data, als_baseline, _ = get_baseline_with_raw(
                x_data,
                original_y_data,
                method="als",
                lam=lambda_val,
                p=p_val,
                smooth=smooth_val,
            )

            # Apply anchor adjustments to ALS baseline
            adjusted_baseline = self._apply_anchor_adjustments(x_data, als_baseline)

            if self.view_toggle_btn.isChecked():
                # Show corrected data - use original raw data minus baseline
                corrected_values = original_y_data - adjusted_baseline
                # print(f"Debug: Corrected data range: {np.min(corrected_values):.4f} to {np.max(corrected_values):.4f}")
                # print(f"Debug: Original data range: {np.min(original_y_data):.4f} to {np.max(original_y_data):.4f}")
                # print(f"Debug: Baseline range: {np.min(adjusted_baseline):.4f} to {np.max(adjusted_baseline):.4f}")

                self.ax.plot(
                    x_data, corrected_values, "g-", label="Corrected", linewidth=1.2
                )
                self.ax.set_title(
                    f'Baseline-Corrected: {self.ylk_data.get("name", "Unknown")}'
                )
            else:
                # Show baseline view (raw + baseline)
                self.ax.plot(
                    x_data,
                    original_y_data,
                    "b-",
                    label="Raw Data",
                    linewidth=1.2,
                    alpha=0.7,
                )
                self.ax.plot(
                    x_data, adjusted_baseline, "r--", label="Baseline", linewidth=1.5
                )
                self.ax.set_title(
                    f'Baseline View: {self.ylk_data.get("name", "Unknown")}'
                )

        except Exception as e:
            # If ALS calculation fails, show raw data
            print(f"Error in baseline calculation: {e}")
            original_y_data = np.array(self.ylk_data["raw_data"]["y"])
            self.ax.plot(x_data, original_y_data, "b-", label="Raw Data", linewidth=1.2)
            self.ax.set_title(
                f'Raw Data: {self.ylk_data.get("name", "Unknown")} (Baseline calc failed: {str(e)})'
            )

        # Draw anchor points if any (only in baseline view)
        if not self.view_toggle_btn.isChecked():
            self.draw_anchors()

        self.ax.set_xlabel("Wavenumber (cm⁻¹)")
        self.ax.set_ylabel("Absorbance")
        if self.parent_analyzer and self.parent_analyzer.show_legend:
            self.ax.legend()
        self.ax.grid(True, alpha=0.3)

        if self.parent_analyzer and self.parent_analyzer.reverse_x_axis:
            self.ax.invert_xaxis()

        # Restore fixed axis limits to prevent auto-expansion
        self.ax.set_xlim(stored_xlim)

        # For corrected data, calculate appropriate Y limits instead of using stored ones
        if self.view_toggle_btn.isChecked():
            # Let matplotlib auto-scale Y axis for corrected data
            self.ax.relim()
            self.ax.autoscale_view(scalex=False, scaley=True)
        else:
            # Use stored limits for baseline view
            self.ax.set_ylim(stored_ylim)

        self.canvas.draw()

    def _plot_with_als_baseline(
        self, x_data, y_data, lambda_val, p_val, als_baseline=None
    ):
        """Helper method to plot with ALS baseline"""
        try:
            # Use pre-calculated baseline if provided, otherwise calculate it
            if als_baseline is None:
                from modules.baseline import baseline_als

                als_baseline = baseline_als(y_data, lam=lambda_val, p=p_val)

            corrected_values = y_data - als_baseline

            if self.view_toggle_btn.isChecked():
                # Show corrected data only
                self.ax.plot(
                    x_data, corrected_values, "g-", label="Corrected", linewidth=1.2
                )
                self.ax.set_title(
                    f'Baseline-Corrected: {self.ylk_data.get("name", "Unknown")}'
                )
            else:
                # Show baseline view (raw + baseline)
                self.ax.plot(
                    x_data, y_data, "b-", label="Raw Data", linewidth=1.2, alpha=0.7
                )
                self.ax.plot(
                    x_data,
                    np.asarray(als_baseline),
                    "r--",
                    label="ALS Baseline",
                    linewidth=1.5,
                )
                self.ax.set_title(
                    f'Baseline View: {self.ylk_data.get("name", "Unknown")}'
                )
        except Exception as e:
            # If calculation fails, show raw data
            self.ax.plot(x_data, y_data, "b-", label="Raw Data", linewidth=1.2)
            self.ax.set_title(
                f'Raw Data: {self.ylk_data.get("name", "Unknown")} (Baseline calc failed)'
            )

    def save_baseline(self):
        lambda_val, p_val, smooth_val = self.get_parameters()

        raw_data = self.ylk_data.get("raw_data", {})
        x_data = np.array(raw_data.get("x", []))
        y_data = np.array(raw_data.get("y", []))

        if len(x_data) == 0 or len(y_data) == 0:
            QMessageBox.warning(self, "Error", "No data available to create baseline")
            return

        if len(x_data) != len(y_data):
            QMessageBox.warning(self, "Error", "Data dimensions do not match")
            return

        if len(x_data) < 10:  # Minimum data points for meaningful baseline
            QMessageBox.warning(
                self,
                "Error",
                "Insufficient data points for baseline calculation (minimum 10 required)",
            )
            return

        try:
            # Calculate baseline with smoothing if enabled
            from modules.baseline import get_baseline_with_raw

            # Always use original raw data for baseline calculation
            original_y_data = np.array(self.ylk_data["raw_data"]["y"])

            _, als_baseline, _ = get_baseline_with_raw(
                x_data,
                original_y_data,
                method="als",
                lam=lambda_val,
                p=p_val,
                smooth=smooth_val,
            )

            # Apply anchor adjustments to ALS baseline
            baseline_values = self._apply_anchor_adjustments(x_data, als_baseline)

            # Save baseline parameters (include both ALS and anchor data)
            baseline_params = {
                "method": "als_with_anchors" if self.anchors else "als",
                "lambda": lambda_val,
                "p": p_val,
                "smooth": smooth_val,
            }

            # Include anchor data if present
            if self.anchors:
                baseline_params["anchors"] = self.anchors

            # Update YLK data structure
            self.ylk_data["baseline"] = {
                "x": x_data.tolist(),
                "y": baseline_values.tolist(),
            }

            # Save parameters used
            if "metadata" not in self.ylk_data:
                self.ylk_data["metadata"] = {}
            self.ylk_data["metadata"]["baseline_params"] = baseline_params

            # Find the original YLK file path and save using the new structure
            if self.parent_analyzer is not None:
                ylk_file_path = None
                folder_data = None
                ylk_data_index = None

                # Search through all folders to find the file
                for folder_path, folder_info in self.parent_analyzer.folders.items():
                    for i, file_path in enumerate(folder_info["files"]):
                        if (
                            os.path.basename(file_path).replace(".ylk", "")
                            == self.filename
                        ):
                            ylk_file_path = file_path
                            folder_data = folder_info
                            ylk_data_index = i
                            break
                    if ylk_file_path:
                        break

                if (
                    ylk_file_path
                    and folder_data is not None
                    and ylk_data_index is not None
                ):
                    if save_ylk_file(ylk_file_path, self.ylk_data):
                        full_filename = os.path.basename(ylk_file_path)
                        QMessageBox.information(
                            self, "Success", f"Baseline saved to {full_filename}"
                        )
                        # Update the folder's YLK data
                        folder_data["ylk_data"][ylk_data_index] = self.ylk_data

                        # Update any selected data that corresponds to this file
                        for i, file_key in enumerate(
                            self.parent_analyzer.selected_files
                        ):
                            folder_path, filename = file_key
                            if filename == self.filename:
                                df = ylk_to_dataframe(self.ylk_data)
                                if df is not None:
                                    self.parent_analyzer.selected_data[i] = df
                                break

                    # Automatically close the tab after saving
                    self.close_tab()
                else:
                    QMessageBox.warning(self, "Error", "Could not find file to save")
            else:
                QMessageBox.warning(
                    self, "Error", f"Could not find file {self.filename}"
                )

        except Exception as e:
            QMessageBox.warning(self, "Error", f"Baseline calculation failed: {str(e)}")

    def close_tab(self):
        """Close this tab"""
        if self.parent_analyzer:
            tab_index = self.parent_analyzer.tab_widget.indexOf(self)
            if tab_index >= 0:
                self.parent_analyzer.tab_widget.removeTab(tab_index)


class FTIRAnalyzer(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("FTIR Data Analysis Tool")
        self.setGeometry(100, 100, 1600, 800)
        # self.showFullScreen()

        # Data storage - updated for multi-folder support
        self.folders = {}  # Dict: folder_path -> {'files': [], 'ylk_data': []}
        self.selected_data = []
        self.selected_files = []  # Store as (folder_path, filename) tuples
        self.visible_files = []  # Track which selected files are visible in plot
        self.reverse_x_axis = False
        self.recent_folders = []
        self.show_baseline_corrected = (
            False  # Toggle for raw vs baseline-corrected data
        )
        self.show_normalized = True  # Toggle for normalized vs absolute values
        self.show_legend = True  # Toggle for showing/hiding legend
        self.show_coordinates = (
            False  # Toggle for showing wavelength coordinates on hover
        )
        self.auto_highlight_ranges = True  # Auto-highlight similar wavenumber ranges

        # Settings for recent folders
        self.settings = QSettings("FTIRTools", "FTIRAnalyzer")
        self.load_recent_folders()

        self.init_ui()

    def init_ui(self):
        # Create menu bar
        self.create_menu_bar()

        # Create main tab widget
        self.tab_widget = QTabWidget()
        self.setCentralWidget(self.tab_widget)

        # Create main analysis tab
        self.create_main_tab()

    def create_menu_bar(self):
        menubar = QMenuBar(self)
        self.setMenuBar(menubar)

        # File menu
        file_menu = menubar.addMenu("File")
        assert file_menu is not None

        add_folder_action = QAction("Add Folder", self)
        add_folder_action.triggered.connect(self.select_folder)
        file_menu.addAction(add_folder_action)

        file_menu.addSeparator()

        # Recent folders submenu
        self.recent_menu = file_menu.addMenu("Open Recent")
        self.update_recent_menu()

        file_menu.addSeparator()

        # Export CSV action
        export_csv_action = QAction("Export Current Graph as CSV", self)
        export_csv_action.triggered.connect(self.export_current_graph_csv)
        file_menu.addAction(export_csv_action)

        # Export PNG action
        export_png_action = QAction("Export Current Graph as PNG", self)
        export_png_action.triggered.connect(self.export_current_graph_png)
        file_menu.addAction(export_png_action)

        # Tools menu
        tools_menu = menubar.addMenu("Tools")
        assert tools_menu is not None

        self.reverse_action = QAction("Reverse X-axis", self)
        self.reverse_action.setCheckable(True)
        self.reverse_action.triggered.connect(self.on_reverse_changed)
        tools_menu.addAction(self.reverse_action)

        self.legend_action = QAction("Hide Legend", self)
        self.legend_action.setCheckable(True)
        self.legend_action.triggered.connect(self.on_legend_toggle)
        tools_menu.addAction(self.legend_action)

        self.coordinates_action = QAction("Show Coordinates on Hover", self)
        self.coordinates_action.setCheckable(True)
        self.coordinates_action.triggered.connect(self.on_coordinates_toggle)
        tools_menu.addAction(self.coordinates_action)

    def create_main_tab(self):
        main_widget = QWidget()
        main_layout = QHBoxLayout(main_widget)

        # Left panel - file management
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_panel.setMaximumWidth(400)

        # File list section
        file_group = QGroupBox("Files")
        file_layout = QVBoxLayout(file_group)

        # All files tree (hierarchical folder structure)
        file_layout.addWidget(QLabel("All Files:"))
        self.file_tree = QTreeWidget()
        self.file_tree.setHeaderHidden(True)  # Hide column header
        self.file_tree.itemDoubleClicked.connect(self.on_file_double_click)
        file_layout.addWidget(self.file_tree)

        # Control buttons
        button_layout = QHBoxLayout()
        add_btn = QPushButton("Add →")
        add_btn.clicked.connect(self.add_to_selected)
        remove_btn = QPushButton("← Remove")
        remove_btn.clicked.connect(self.remove_from_selected)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_selected)

        button_layout.addWidget(add_btn)
        button_layout.addWidget(remove_btn)
        button_layout.addWidget(clear_btn)
        file_layout.addLayout(button_layout)

        # Selected files list with right-click menu and checkboxes
        file_layout.addWidget(QLabel("Selected for Analysis:"))
        self.selected_listbox = QListWidget()
        self.selected_listbox.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.selected_listbox.customContextMenuRequested.connect(
            self.show_selected_context_menu
        )
        self.selected_listbox.itemDoubleClicked.connect(self.on_selected_double_click)
        self.selected_listbox.itemChanged.connect(self.on_selected_item_changed)

        # Track mouse press position to distinguish checkbox clicks from double-clicks
        self.selected_listbox.mousePressEvent = self.selected_list_mouse_press
        file_layout.addWidget(self.selected_listbox)

        left_layout.addWidget(file_group)

        # Analysis buttons
        analysis_group = QGroupBox("Analysis")
        analysis_layout = QVBoxLayout(analysis_group)

        # Normalization toggle button
        self.normalize_btn = QPushButton("Show Normalized Data")
        self.normalize_btn.clicked.connect(self.toggle_normalization)
        self.normalize_btn.setCheckable(True)
        analysis_layout.addWidget(self.normalize_btn)

        # Toggle button for raw/baseline-corrected data
        self.data_toggle_btn = QPushButton("Show Baseline-Corrected Data")
        self.data_toggle_btn.clicked.connect(self.toggle_data_display)
        self.data_toggle_btn.setCheckable(True)
        analysis_layout.addWidget(self.data_toggle_btn)

        # self.corr_btn = QPushButton("Calculate Correlation")
        # self.corr_btn.clicked.connect(self.calculate_correlation)
        # analysis_layout.addWidget(self.corr_btn)

        left_layout.addWidget(analysis_group)
        left_layout.addStretch()

        main_layout.addWidget(left_panel)

        # Right panel - embedded plotting
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        plot_group = QGroupBox("Spectral View")
        plot_layout = QVBoxLayout(plot_group)

        # Create matplotlib figure and canvas
        self.figure = Figure(figsize=(10, 6))
        self.canvas = FigureCanvas(self.figure)
        self.ax = self.figure.add_subplot(111)

        plot_layout.addWidget(self.canvas)
        right_layout.addWidget(plot_group)

        main_layout.addWidget(right_panel)

        # Add main tab
        self.tab_widget.addTab(main_widget, "Main Analysis")

    def on_reverse_changed(self, state):
        if isinstance(state, bool):
            self.reverse_x_axis = state
        else:
            self.reverse_x_axis = state == Qt.CheckState.Checked.value
        # Update the current plot if there are selected files
        if self.selected_data:
            self.plot_spectra()

    def on_legend_toggle(self, state):
        if isinstance(state, bool):
            self.show_legend = not state  # Action is "Hide Legend", so invert
        else:
            self.show_legend = not (state == Qt.CheckState.Checked.value)
        # Update the current plot if there are selected files
        if self.selected_data:
            self.plot_spectra()

    def on_coordinates_toggle(self, state):
        if isinstance(state, bool):
            self.show_coordinates = state
        else:
            self.show_coordinates = state == Qt.CheckState.Checked.value

        # Enable/disable hover functionality
        if self.show_coordinates:
            if not hasattr(self, "_hover_connection_id"):
                self._hover_connection_id = self.canvas.mpl_connect(
                    "motion_notify_event", self.on_main_plot_hover
                )
        else:
            # Disconnect hover event and clear crosshairs
            if hasattr(self, "_hover_connection_id"):
                self.canvas.mpl_disconnect(self._hover_connection_id)
                delattr(self, "_hover_connection_id")

            # Clean up crosshairs and text
            if hasattr(self, "crosshair_h"):
                try:
                    self.crosshair_h.set_visible(False)
                    # Don't try to remove axhline/axvline, just hide them
                except:
                    pass
                del self.crosshair_h
            if hasattr(self, "crosshair_v"):
                try:
                    self.crosshair_v.set_visible(False)
                except:
                    pass
                del self.crosshair_v
            if hasattr(self, "coord_text"):
                self.coord_text.remove()
                del self.coord_text
            if hasattr(
                self, "hover_annotation"
            ):  # Clean up old annotation if it exists
                self.hover_annotation.remove()
                del self.hover_annotation
            self.canvas.draw()

    def on_main_plot_hover(self, event):
        """Handle mouse hover over main plot to show coordinates with crosshairs"""
        if event.inaxes != self.ax or not self.selected_data:
            # Hide crosshairs if mouse is outside plot area
            self._hide_crosshairs()
            return

        if event.xdata is None or event.ydata is None:
            return

        # Store current axis limits to prevent shifting
        xlim = self.ax.get_xlim()
        ylim = self.ax.get_ylim()

        # Create crosshairs if they don't exist
        if not hasattr(self, "crosshair_h"):
            self.crosshair_h = self.ax.axhline(
                y=event.ydata, color="gray", linestyle="--", alpha=0.7, linewidth=0.8
            )
            self.crosshair_v = self.ax.axvline(
                x=event.xdata, color="gray", linestyle="--", alpha=0.7, linewidth=0.8
            )
        else:
            # Update positions
            self.crosshair_h.set_ydata([event.ydata, event.ydata])
            self.crosshair_v.set_xdata([event.xdata, event.xdata])

        # Make visible
        self.crosshair_h.set_visible(True)
        self.crosshair_v.set_visible(True)

        # Update coordinate text at bottom of plot
        wavenumber = event.xdata
        absorbance = event.ydata

        # Create or update text display
        if not hasattr(self, "coord_text"):
            # Position text at bottom left of the plot
            self.coord_text = self.ax.text(
                0.02,
                0.02,
                "",
                transform=self.ax.transAxes,
                fontsize=9,
                bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.8),
            )

        # Format text
        text = f"Wavenumber: {wavenumber:.1f} cm⁻¹  |  Absorbance: {absorbance:.4f}"
        self.coord_text.set_text(text)
        self.coord_text.set_visible(True)

        # Restore axis limits to prevent shifting
        self.ax.set_xlim(xlim)
        self.ax.set_ylim(ylim)

        self.canvas.draw()

    def _hide_crosshairs(self):
        """Hide crosshairs and coordinate text"""
        if hasattr(self, "crosshair_h"):
            self.crosshair_h.set_visible(False)
            self.crosshair_v.set_visible(False)
        if hasattr(self, "coord_text"):
            self.coord_text.set_visible(False)
        self.canvas.draw()

    def on_selected_item_changed(self, item):
        """Handle checkbox changes for selected files visibility"""
        file_key = item.data(Qt.ItemDataRole.UserRole)  # Get the stored file key
        is_visible = item.checkState() == Qt.CheckState.Checked

        if file_key in self.selected_files:
            file_index = self.selected_files.index(file_key)
            # Ensure visible_files list is the same size
            while len(self.visible_files) <= file_index:
                self.visible_files.append(True)
            self.visible_files[file_index] = is_visible

            # Update plot
            if self.selected_data:
                self.plot_spectra()

    def selected_list_mouse_press(self, event):
        """Custom mouse press handler for selected files list to track click position"""
        # Store the original mousePressEvent behavior
        original_mouse_press = QListWidget.mousePressEvent
        original_mouse_press(self.selected_listbox, event)

        # Store click position for double-click detection
        item = self.selected_listbox.itemAt(event.pos())
        if item:
            # Check if click is on the checkbox area (left part of item)
            item_rect = self.selected_listbox.visualItemRect(item)
            checkbox_area_width = 20  # Approximate checkbox width
            self._last_click_on_checkbox = (
                event.pos().x() < item_rect.left() + checkbox_area_width
            )
        else:
            self._last_click_on_checkbox = False

    def load_recent_folders(self):
        """Load recent folders from settings"""
        self.recent_folders = self.settings.value("recent_folders", [])
        if not isinstance(self.recent_folders, list):
            self.recent_folders = []

    def save_recent_folders(self):
        """Save recent folders to settings"""
        self.settings.setValue("recent_folders", self.recent_folders)

    def add_recent_folder(self, folder_path):
        """Add folder to recent list"""
        if folder_path in self.recent_folders:
            self.recent_folders.remove(folder_path)
        self.recent_folders.insert(0, folder_path)
        # Keep only last 10 folders
        self.recent_folders = self.recent_folders[:10]
        self.save_recent_folders()
        self.update_recent_menu()

    def update_recent_menu(self):
        """Update the recent folders menu"""
        if self.recent_menu is not None:
            self.recent_menu.clear()
            for folder in self.recent_folders:
                action = QAction(folder, self)
                action.triggered.connect(
                    lambda checked, path=folder: self.process_folder(path)
                )
                self.recent_menu.addAction(action)

            if not self.recent_folders:
                action = QAction("No recent folders", self)
                action.setEnabled(False)
                self.recent_menu.addAction(action)

    def show_selected_context_menu(self, position):
        """Show right-click context menu for selected files list"""
        item = self.selected_listbox.itemAt(position)
        if item is None:
            return

        menu = QMenu(self)

        create_baseline_action = QAction("Create Baseline", self)
        create_baseline_action.triggered.connect(
            lambda: self.create_baseline_for_file(item.text())
        )
        menu.addAction(create_baseline_action)

        menu.exec(self.selected_listbox.mapToGlobal(position))

    def toggle_data_display(self):
        """Toggle between raw data and baseline-corrected data display"""
        self.show_baseline_corrected = self.data_toggle_btn.isChecked()

        # Update button text
        if self.show_baseline_corrected:
            self.data_toggle_btn.setText("Show Raw Data")
        else:
            self.data_toggle_btn.setText("Show Baseline-Corrected Data")

        # Update plot if there are selected files
        if self.selected_data:
            self.plot_spectra()

    def toggle_normalization(self):
        """Toggle between normalized and absolute values display"""
        self.show_normalized = self.normalize_btn.isChecked()

        # Update button text
        if self.show_normalized:
            self.normalize_btn.setText("Show Absolute Values")
        else:
            self.normalize_btn.setText("Show Normalized Data")

        # Update plot if there are selected files
        if self.selected_data:
            self.plot_spectra()

    def create_baseline_for_file(self, filename):
        """Create a new baseline tab for the specified file"""
        # Find the YLK data for this file using the file key
        ylk_data = None
        file_key = None

        # Extract just the filename part from the display name if needed
        if " (" in filename:
            # Handle display format "filename (folder)"
            actual_filename = filename.split(" (")[0]
        else:
            actual_filename = filename

        # Search through selected files to find matching data
        for i, selected_file_key in enumerate(self.selected_files):
            folder_path, file_name = selected_file_key
            if file_name == actual_filename or filename.startswith(file_name):
                # Found the file, now get the YLK data from folders
                if folder_path in self.folders:
                    folder_data = self.folders[folder_path]
                    for j, file_path in enumerate(folder_data["files"]):
                        if os.path.basename(file_path).replace(".ylk", "") == file_name:
                            ylk_data = folder_data["ylk_data"][j]
                            file_key = selected_file_key
                            break
                if ylk_data:
                    break

        if ylk_data is None:
            QMessageBox.warning(
                self, "Error", f"Could not find data for file {filename}"
            )
            return

        # Create baseline creation tab
        baseline_tab = BaselineCreationTab(ylk_data, actual_filename, self)
        tab_name = f"Baseline: {actual_filename}"
        tab_index = self.tab_widget.addTab(baseline_tab, tab_name)
        self.tab_widget.setCurrentIndex(tab_index)

    def add_to_selected(self):
        """Add selected files to analysis list"""
        selected_items = self.file_tree.selectedItems()
        items_to_remove = []

        for item in selected_items:
            # Skip folder items (only process file items)
            if item.parent() is None:  # This is a folder item
                continue

            folder_path = item.parent().text(0)  # Get folder name
            filename = item.text(0)  # Get filename
            file_key = (folder_path, filename)

            if file_key not in self.selected_files:
                self.selected_files.append(file_key)
                self.visible_files.append(True)  # New files are visible by default

                # Find YLK data for this file
                if folder_path in self.folders:
                    folder_data = self.folders[folder_path]
                    for i, file_path in enumerate(folder_data["files"]):
                        if os.path.basename(file_path).replace(".ylk", "") == filename:
                            ylk_data = folder_data["ylk_data"][i]
                            df = ylk_to_dataframe(ylk_data)
                            if df is not None:
                                self.selected_data.append(df)
                                # Create checkable item with folder info
                                display_name = (
                                    f"{filename} ({os.path.basename(folder_path)})"
                                )
                                list_item = QListWidgetItem(display_name)
                                list_item.setFlags(
                                    list_item.flags() | Qt.ItemFlag.ItemIsUserCheckable
                                )
                                list_item.setCheckState(
                                    Qt.CheckState.Checked
                                )  # Checked by default
                                # Store the file key as item data
                                list_item.setData(Qt.ItemDataRole.UserRole, file_key)
                                self.selected_listbox.addItem(list_item)
                                items_to_remove.append(item)
                            break

        # Remove selected files from tree
        for item in items_to_remove:
            parent = item.parent()
            if parent:
                parent.removeChild(item)
                # If parent folder is now empty, we could optionally remove it
                # but let's keep empty folders visible

        # Update main plot
        if self.selected_data:
            self.plot_spectra()

    def remove_from_selected(self):
        """Remove selected files from analysis list"""
        selected_items = self.selected_listbox.selectedItems()
        for item in selected_items:
            filename = item.text()
            row = self.selected_listbox.row(item)
            self.selected_listbox.takeItem(row)
            if filename in self.selected_files:
                file_index = self.selected_files.index(filename)
                self.selected_files.pop(file_index)
                self.selected_data.pop(file_index)
                # Also remove from visible_files if it exists
                if file_index < len(self.visible_files):
                    self.visible_files.pop(file_index)

        # Rebuild file tree to restore files
        self._rebuild_file_tree()

        # Update main plot
        self.plot_spectra()

    def clear_selected(self):
        """Clear analysis list"""
        self.selected_listbox.clear()
        self.selected_files.clear()
        self.selected_data.clear()
        self.visible_files.clear()  # Also clear visibility tracking

        # Rebuild file tree to restore all files
        self._rebuild_file_tree()

        # Clear main plot
        self.ax.clear()
        self.ax.set_xlabel("Wavenumber (cm⁻¹)")
        self.ax.set_ylabel("Absorbance")
        self.ax.set_title("Select files to display spectra")
        self.canvas.draw()

    def on_file_double_click(self, item):
        """Handle double-click on file tree to move file to selected"""
        # Skip folder items (only process file items)
        if item.parent() is None:  # This is a folder item
            return

        folder_path = item.parent().data(0, Qt.ItemDataRole.UserRole)  # Get folder path
        filename = item.text(0)  # Get filename
        file_key = (folder_path, filename)

        if file_key not in self.selected_files:
            self.selected_files.append(file_key)
            self.visible_files.append(True)  # New files are visible by default

            # Find YLK data for this file
            if folder_path in self.folders:
                folder_data = self.folders[folder_path]
                for i, file_path in enumerate(folder_data["files"]):
                    if os.path.basename(file_path).replace(".ylk", "") == filename:
                        ylk_data = folder_data["ylk_data"][i]
                        df = ylk_to_dataframe(ylk_data)
                        if df is not None:
                            self.selected_data.append(df)
                            # Create checkable item with folder info
                            display_name = (
                                f"{filename} ({os.path.basename(folder_path)})"
                            )
                            list_item = QListWidgetItem(display_name)
                            list_item.setFlags(
                                list_item.flags() | Qt.ItemFlag.ItemIsUserCheckable
                            )
                            list_item.setCheckState(
                                Qt.CheckState.Checked
                            )  # Checked by default
                            # Store the file key as item data
                            list_item.setData(Qt.ItemDataRole.UserRole, file_key)
                            self.selected_listbox.addItem(list_item)
                            # Rebuild tree to hide selected files
                            self._rebuild_file_tree()
                            self.plot_spectra()
                        break

    def on_selected_double_click(self, item):
        """Handle double-click on selected listbox to remove file from selected"""
        # Don't remove file if the double-click was on the checkbox area
        if hasattr(self, "_last_click_on_checkbox") and self._last_click_on_checkbox:
            return

        file_key = item.data(Qt.ItemDataRole.UserRole)  # Get the stored file key
        row = self.selected_listbox.row(item)
        self.selected_listbox.takeItem(row)
        if file_key in self.selected_files:
            file_index = self.selected_files.index(file_key)
            self.selected_files.pop(file_index)
            self.selected_data.pop(file_index)
            # Also remove from visible_files if it exists
            if file_index < len(self.visible_files):
                self.visible_files.pop(file_index)
            # Rebuild file tree to restore files
            self._rebuild_file_tree()
            self.plot_spectra()

    # Note: _find_data_index_by_filename and _rebuild_all_files_list methods removed
    # as they are no longer needed with the new multi-folder tree structure

    def export_current_graph_csv(self):
        """Export current graph data (raw and corrected) to CSV file"""
        if not self.selected_data or not self.selected_files:
            QMessageBox.information(
                self, "No Data", "No files are currently selected for analysis."
            )
            return

        # Get save file path
        default_filename = "ftir_export.csv"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export CSV", default_filename, "CSV files (*.csv);;All files (*.*)"
        )

        if not file_path:
            return

        try:
            import csv
            import pandas as pd

            # Collect all data first
            all_data = {}

            for file_key, df in zip(self.selected_files, self.selected_data):
                # Get raw data
                wavenumber = df["wavenumber"].values
                raw_absorbance = df["absorbance"].values

                # Calculate corrected data using baseline if available
                folder_path, file_name = file_key
                corrected_absorbance = raw_absorbance.copy()  # Default to raw
                ylk_data = None

                # Find YLK data for this file
                if folder_path in self.folders:
                    folder_data = self.folders[folder_path]
                    for j, file_path in enumerate(folder_data["files"]):
                        if os.path.basename(file_path).replace(".ylk", "") == file_name:
                            ylk_data = folder_data["ylk_data"][j]
                            break

                if ylk_data is not None:
                    baseline_data = ylk_data.get("baseline", {})

                    if baseline_data.get("x") and baseline_data.get("y"):
                        try:
                            # Use saved baseline
                            baseline_x = np.array(baseline_data["x"])
                            baseline_y = np.array(baseline_data["y"])

                            # Interpolate baseline to match raw data x-values if needed
                            if len(baseline_x) != len(wavenumber) or not np.allclose(
                                baseline_x, wavenumber, rtol=1e-6
                            ):
                                from scipy.interpolate import interp1d

                                baseline_func = interp1d(
                                    baseline_x,
                                    baseline_y,
                                    kind="linear",
                                    bounds_error=False,
                                    fill_value="extrapolate",
                                )
                                baseline_interpolated = baseline_func(wavenumber)
                            else:
                                baseline_interpolated = baseline_y

                            corrected_absorbance = (
                                raw_absorbance - baseline_interpolated
                            )
                        except Exception as e:
                            print(
                                f"Error calculating baseline correction for {file_name}: {e}"
                            )
                            # Keep raw data as corrected if baseline calculation fails

                # Clean filename for column names
                clean_filename = (
                    file_name.replace(" ", "_").replace(".", "_").replace("-", "_")
                )

                # Store data
                all_data[f"wavenumber_{clean_filename}"] = wavenumber
                all_data[f"{clean_filename}_raw"] = raw_absorbance
                all_data[f"{clean_filename}_corrected"] = corrected_absorbance

            # Create DataFrame and save
            if all_data:
                df_export = pd.DataFrame(all_data)
                df_export.to_csv(file_path, index=False)

                QMessageBox.information(
                    self,
                    "Export Complete",
                    f"Data exported successfully to:\n{file_path}\n\n"
                    f"Exported {len(self.selected_files)} file(s).\n"
                    f"Each file has wavenumber, raw, and corrected columns.",
                )
            else:
                QMessageBox.warning(self, "Export Error", "No data to export")

        except Exception as e:
            QMessageBox.critical(
                self, "Export Error", f"Failed to export CSV file:\n{str(e)}"
            )

    def export_current_graph_png(self):
        """Export current graph as PNG file"""
        if not self.selected_data or not self.selected_files:
            QMessageBox.information(
                self, "No Data", "No files are currently selected for analysis."
            )
            return

        # Get save file path
        default_filename = "ftir_graph.png"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export PNG", default_filename, "PNG files (*.png);;All files (*.*)"
        )

        if not file_path:
            return

        try:
            # Save the current figure
            self.figure.savefig(
                file_path,
                dpi=300,
                bbox_inches="tight",
                facecolor="white",
                edgecolor="none",
                format="png",
            )

            QMessageBox.information(
                self, "Export Complete", f"Graph exported successfully to:\n{file_path}"
            )

        except Exception as e:
            QMessageBox.critical(
                self, "Export Error", f"Failed to export PNG file:\n{str(e)}"
            )

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self, "Select folder with .jws files"
        )
        if folder_path:
            self.add_recent_folder(folder_path)
            self.process_folder(folder_path)

    def process_folder(self, folder_path):
        """Add folder: convert JWS files to YLK format and load existing YLK files"""
        try:
            # Create converted_ylk subfolder
            ylk_folder = os.path.join(folder_path, "converted_ylk")
            os.makedirs(ylk_folder, exist_ok=True)

            # Check if folder is already loaded
            if folder_path in self.folders:
                QMessageBox.information(
                    self,
                    "Folder Already Loaded",
                    f"Folder {folder_path} is already loaded.",
                )
                return

            # Scan folder for .jws and .ylk files
            processed_files = []

            # Process JWS files first
            for filename in os.listdir(folder_path):
                file_path = os.path.join(folder_path, filename)
                if os.path.isfile(file_path) and filename.endswith(".jws"):
                    # Convert .jws file to .ylk
                    ylk_file = convert_jws_with_fallback(file_path, ylk_folder)
                    if ylk_file:
                        processed_files.append(ylk_file)

            # Also check for existing YLK files in the ylk folder
            if os.path.exists(ylk_folder):
                for filename in os.listdir(ylk_folder):
                    if filename.endswith(".ylk"):
                        ylk_file_path = os.path.join(ylk_folder, filename)
                        if ylk_file_path not in processed_files:
                            processed_files.append(ylk_file_path)

            # Load all YLK files - sort by filename
            processed_files.sort()
            ylk_data_list = []
            files = []

            for ylk_file in processed_files:
                try:
                    ylk_data = load_ylk_file(ylk_file)
                    if ylk_data:
                        ylk_data_list.append(ylk_data)
                        files.append(ylk_file)
                except Exception as e:
                    QMessageBox.warning(
                        self, "Warning", f"Unable to load file {ylk_file}: {str(e)}"
                    )

            # Store folder data
            self.folders[folder_path] = {"files": files, "ylk_data": ylk_data_list}

            # Update tree widget
            self._rebuild_file_tree()

            # QMessageBox.information(
            #     self, "Complete", f"Added folder with {len(processed_files)} files"
            # )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error processing folder: {str(e)}")

    def _rebuild_file_tree(self):
        """Rebuild the file tree with hierarchical folder structure"""
        # Store current expansion states before clearing
        expanded_folders = set()
        for i in range(self.file_tree.topLevelItemCount()):
            item = self.file_tree.topLevelItem(i)
            if item and item.isExpanded():
                folder_path = item.data(0, Qt.ItemDataRole.UserRole)
                if folder_path:
                    expanded_folders.add(folder_path)

        self.file_tree.clear()

        # Get reference wavenumber ranges from selected files for highlighting
        reference_ranges = self._get_selected_wavenumber_ranges()

        for folder_path, folder_data in self.folders.items():
            # Create folder item
            folder_name = os.path.basename(folder_path)
            folder_item = QTreeWidgetItem([folder_name])
            folder_item.setData(
                0, Qt.ItemDataRole.UserRole, folder_path
            )  # Store full path

            # Add file items under folder
            for i, file_path in enumerate(folder_data["files"]):
                ylk_data = folder_data["ylk_data"][i]
                filename = os.path.basename(file_path).replace(".ylk", "")

                # Skip files that are already selected
                if (folder_path, filename) not in self.selected_files:
                    file_item = QTreeWidgetItem([filename])
                    file_item.setData(
                        0, Qt.ItemDataRole.UserRole, file_path
                    )  # Store full file path

                    # Apply highlighting for similar wavenumber ranges
                    if self.auto_highlight_ranges and reference_ranges:
                        file_range = self._get_file_wavenumber_range(ylk_data)
                        if self._is_similar_range(file_range, reference_ranges):
                            # Keep similar files in normal color (black) and add tooltip
                            file_item.setToolTip(
                                0,
                                f"Similar range: {file_range[0]:.0f}-{file_range[1]:.0f} cm⁻¹",
                            )
                        else:
                            # Make non-similar files light gray
                            file_item.setForeground(0, Qt.GlobalColor.darkGray)

                    folder_item.addChild(file_item)

            # Only add folder to tree if it has children (unselected files)
            if folder_item.childCount() > 0:
                self.file_tree.addTopLevelItem(folder_item)
                # Restore expansion state or expand by default for new folders
                if folder_path in expanded_folders or len(expanded_folders) == 0:
                    folder_item.setExpanded(True)

    def _get_selected_wavenumber_ranges(self):
        """Get wavenumber ranges from currently selected files"""
        ranges = []
        for file_key in self.selected_files:
            folder_path, filename = file_key
            if folder_path in self.folders:
                folder_data = self.folders[folder_path]
                for i, file_path in enumerate(folder_data["files"]):
                    if os.path.basename(file_path).replace(".ylk", "") == filename:
                        ylk_data = folder_data["ylk_data"][i]
                        file_range = self._get_file_wavenumber_range(ylk_data)
                        if file_range:
                            ranges.append(file_range)
                        break
        return ranges

    def _get_file_wavenumber_range(self, ylk_data):
        """Get wavenumber range (min, max) from YLK data"""
        try:
            raw_data = ylk_data.get("raw_data", {})
            x_data = raw_data.get("x", [])
            if x_data:
                return (min(x_data), max(x_data))
        except Exception:
            pass
        return None

    def _is_similar_range(self, file_range, reference_ranges, tolerance=50):
        """Check if a file's wavenumber range is similar to any reference range"""
        if not file_range or not reference_ranges:
            return False

        file_min, file_max = file_range

        for ref_min, ref_max in reference_ranges:
            # Check if ranges overlap or are within tolerance
            min_diff = abs(file_min - ref_min)
            max_diff = abs(file_max - ref_max)

            # Consider similar if both endpoints are within tolerance
            # or if there's significant overlap
            if (min_diff <= tolerance and max_diff <= tolerance) or (
                file_min <= ref_max and file_max >= ref_min
            ):  # Ranges overlap
                return True

        return False

    def plot_spectra(self):
        """Plot selected spectra in the main window canvas"""
        # Clear the plot
        self.ax.clear()

        if not self.selected_data:
            self.ax.set_xlabel("Wavenumber (cm⁻¹)")
            self.ax.set_ylabel("Absorbance")
            self.ax.set_title("Select files to display spectra")
            self.ax.grid(True, alpha=0.3)
            self.canvas.draw()
            return

        # Plot each selected spectrum (only visible ones)
        visible_count = 0
        for i, df in enumerate(self.selected_data):
            # Skip if file is not visible
            if i < len(self.visible_files) and not self.visible_files[i]:
                continue

            file_key = self.selected_files[i]
            folder_path, filename = file_key  # Extract filename from file key
            visible_count += 1

            if self.show_baseline_corrected:
                # Show baseline-corrected data
                # Find the corresponding YLK data to get baseline-corrected values
                ylk_data = None
                # Find the YLK data by matching filename and folder
                if folder_path in self.folders:
                    folder_data = self.folders[folder_path]
                    for j, file_path in enumerate(folder_data["files"]):
                        if os.path.basename(file_path).replace(".ylk", "") == filename:
                            ylk_data = folder_data["ylk_data"][j]
                            break

                if (
                    ylk_data
                    and ylk_data.get("baseline", {}).get("x")
                    and ylk_data.get("baseline", {}).get("y")
                ):
                    try:
                        # Use baseline-corrected data if available
                        raw_x = np.array(ylk_data["raw_data"]["x"])
                        raw_y = np.array(ylk_data["raw_data"]["y"])
                        baseline_x = np.array(ylk_data["baseline"]["x"])
                        baseline_y = np.array(ylk_data["baseline"]["y"])

                        # Ensure baseline and raw data have same x coordinates
                        if len(baseline_x) != len(raw_x) or not np.allclose(
                            baseline_x, raw_x
                        ):
                            # Interpolate baseline to match raw data x-values
                            from scipy.interpolate import interp1d

                            baseline_func = interp1d(
                                baseline_x,
                                baseline_y,
                                kind="linear",
                                bounds_error=False,
                                fill_value="extrapolate",
                            )
                            baseline_y_interp = baseline_func(raw_x)
                        else:
                            baseline_y_interp = baseline_y

                        corrected_y = raw_y - baseline_y_interp

                        # Apply normalization if requested
                        if self.show_normalized:
                            corrected_y = (
                                corrected_y / np.max(np.abs(corrected_y))
                                if np.max(np.abs(corrected_y)) > 0
                                else corrected_y
                            )

                        self.ax.plot(
                            raw_x,
                            corrected_y,
                            label=f"{filename} (Baseline-corrected)",
                            linewidth=1.2,
                        )
                    except Exception as e:
                        print(
                            f"Error processing baseline-corrected data for {filename}: {e}"
                        )
                        # Fall back to raw data
                        pre_df = preprocess_data(df, normalize=self.show_normalized)
                        self.ax.plot(
                            pre_df["wavenumber"],
                            pre_df["absorbance"],
                            label=f"{filename} (Error - using raw)",
                            linewidth=1.2,
                            linestyle="-.",
                            alpha=0.7,
                        )
                else:
                    # Fall back to raw data if no baseline available
                    pre_df = preprocess_data(df, normalize=self.show_normalized)
                    self.ax.plot(
                        pre_df["wavenumber"],
                        pre_df["absorbance"],
                        label=f"{filename} (Raw - no baseline)",
                        linewidth=1.2,
                        linestyle="--",
                    )
            else:
                # Show raw data (normalized or absolute based on toggle)
                pre_df = preprocess_data(df, normalize=self.show_normalized)
                self.ax.plot(
                    pre_df["wavenumber"],
                    pre_df["absorbance"],
                    label=str(filename),
                    linewidth=1.2,
                )

        # Set labels and title
        self.ax.set_xlabel("Wavenumber (cm⁻¹)")

        # Set y-axis label based on data type
        if self.show_baseline_corrected:
            ylabel = "Baseline-corrected Absorbance"
            title_suffix = "Baseline-corrected"
        else:
            ylabel = "Normalized Absorbance" if self.show_normalized else "Absorbance"
            title_suffix = "Normalized" if self.show_normalized else "Raw"

        self.ax.set_ylabel(ylabel)
        self.ax.set_title(f"FTIR Spectra ({title_suffix})")

        self.ax.grid(True, alpha=0.3)

        # Apply x-axis reversal if enabled
        if self.reverse_x_axis:
            self.ax.invert_xaxis()

        # Add legend if there are multiple visible spectra and legend is enabled
        if visible_count > 1 and self.show_legend:
            self.ax.legend()

        # Update canvas
        self.canvas.draw()

    def calculate_correlation(self):
        if len(self.selected_data) < 2 or len(self.selected_data) > 5:
            QMessageBox.warning(self, "警告", "請選擇2-5筆資料進行相關性分析")
            return

        # Process data for correlation analysis using current normalization setting
        processed_data = [
            preprocess_data(df, normalize=self.show_normalized)
            for df in self.selected_data
        ]
        corr_matrix = calculate_correlation_matrix(processed_data)
        num = len(self.selected_data)

        # Set up OriginLab style
        setup_originlab_style()

        # 建立相關性矩陣熱圖
        fig, ax = plt.subplots(figsize=(9, 7))

        # 繪製熱圖 with OriginLab-like colors
        im = ax.imshow(
            corr_matrix,
            cmap="RdBu_r",
            aspect="auto",
            vmin=-1,
            vmax=1,
            interpolation="nearest",
        )

        # 設定標籤
        file_labels = [f.replace(".csv", "") for f in self.selected_files]
        ax.set_xticks(range(num))
        ax.set_yticks(range(num))
        ax.set_xticklabels(file_labels, rotation=45, ha="right", fontweight="normal")
        ax.set_yticklabels(file_labels, fontweight="normal")

        # 在每個格子中顯示數值
        for i in range(num):
            for j in range(num):
                text = ax.text(
                    j,
                    i,
                    f"{corr_matrix[i, j]:.3f}",
                    ha="center",
                    va="center",
                    color="white" if abs(corr_matrix[i, j]) > 0.5 else "black",
                    fontsize=10,
                    fontweight="bold",
                )

        # Format with OriginLab style
        format_originlab_plot(ax, "Correlation Matrix", "", "", show_minor_ticks=False)

        # Hide legend for correlation matrix if legend toggle is disabled
        if not self.show_legend:
            legend = ax.get_legend()
            if legend:
                legend.set_visible(False)

        # 添加色條 with OriginLab styling
        cbar = plt.colorbar(im, ax=ax, shrink=0.8)
        cbar.set_label(
            "Correlation Coefficient", rotation=270, labelpad=20, fontweight="bold"
        )
        cbar.ax.tick_params(labelsize=10)

        plt.tight_layout()
        plt.show()  # Open in separate window

        # 也顯示數值結果
        result_text = "Correlation Matrix:\n\n"
        result_text += "Files: " + ", ".join(file_labels) + "\n\n"
        for i, row in enumerate(corr_matrix):
            result_text += (
                f"{file_labels[i]}: " + " ".join([f"{val:.3f}" for val in row]) + "\n"
            )

        QMessageBox.information(self, "Correlation Results", result_text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    analyzer = FTIRAnalyzer()
    analyzer.show()
    sys.exit(app.exec())
