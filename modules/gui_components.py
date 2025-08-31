"""
GUI Components Module for FTIR Spectroscopy

Contains complex GUI components that are loaded on-demand to improve startup performance.
This includes the BaselineCreationTab which is only loaded when baseline correction is needed.
"""

import os
import numpy as np
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QCheckBox,
    QFormLayout,
    QLineEdit,
    QGroupBox,
    QMessageBox,
    QMenu,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction

from modules.file_converter import save_ylk_file, ylk_to_dataframe


class BaselineCreationTab(QWidget):
    """
    Tab widget for creating and editing baseline corrections for FTIR spectra.
    This widget is loaded on-demand to improve application startup performance.
    """
    
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
            _, als_baseline, _ = get_baseline_with_raw(
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
        except Exception:
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