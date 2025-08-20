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
    QDialog,
    QFormLayout,
    QLineEdit,
    QRadioButton,
    QButtonGroup,
    QDialogButtonBox,
    QScrollArea,
    QFrame,
    QAbstractItemView,
    QSplitter,
    QGroupBox,
    QGridLayout,
    QSpinBox,
    QDoubleSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QMenuBar,
    QMenu,
    QTabWidget,
    QListWidgetItem,
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

        self.lambda_edit = QLineEdit("1e5")
        params_layout.addRow("Lambda (smoothness):", self.lambda_edit)

        self.p_edit = QLineEdit("0.01")
        params_layout.addRow("P (asymmetry):", self.p_edit)

        self.smooth_checkbox = QCheckBox("Apply smoothing")
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

        # Store anchors for manual baseline adjustment
        self.anchors = []
        self.selected_anchor = None
        self.dragging = False

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

        # Check if we clicked near an anchor point
        if self.anchors:
            # Convert mouse coordinates to display coordinates
            mouse_display_x, mouse_display_y = self.ax.transData.transform([event.xdata, event.ydata])

            # Find the closest anchor within a certain pixel radius
            min_pixel_distance = float("inf")
            closest_anchor_idx = None

            for i, (anchor_x, anchor_y) in enumerate(self.anchors):
                # Convert anchor coordinates to display coordinates
                anchor_display_x, anchor_display_y = self.ax.transData.transform([anchor_x, anchor_y])
                
                # Calculate pixel distance
                pixel_distance = ((mouse_display_x - anchor_display_x) ** 2 + (mouse_display_y - anchor_display_y) ** 2) ** 0.5
                if pixel_distance < min_pixel_distance and pixel_distance < 10:  # 10 pixel threshold
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

    def on_mouse_move(self, event):
        """Handle mouse move events for dragging anchors"""
        if event.inaxes != self.ax or not self.dragging or self.selected_anchor is None:
            return

        # Update the position of the selected anchor
        if event.xdata is not None and event.ydata is not None:
            self.anchors[self.selected_anchor] = (event.xdata, event.ydata)
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

        # Clear anchors action
        clear_anchors_action = QAction("Clear All Anchors", self)
        clear_anchors_action.triggered.connect(self.clear_anchors)
        menu.addAction(clear_anchors_action)

        menu.exec(self.canvas.mapToGlobal(position))

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
                    from modules.baseline import baseline_als

                    als_baseline = baseline_als(y_data, lam=lambda_val, p=p_val)

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
            return

        self.ax.clear()

        try:
            from modules.baseline import baseline_als

            # Calculate ALS baseline
            als_baseline = baseline_als(y_data, lam=lambda_val, p=p_val)
            
            # Apply anchor adjustments to ALS baseline
            adjusted_baseline = self._apply_anchor_adjustments(x_data, als_baseline)

            if self.view_toggle_btn.isChecked():
                # Show corrected data
                corrected_values = y_data - adjusted_baseline
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
                    x_data, adjusted_baseline, "r--", label="Baseline", linewidth=1.5
                )
                self.ax.set_title(
                    f'Baseline View: {self.ylk_data.get("name", "Unknown")}'
                )

        except Exception:
            # If ALS calculation fails, show raw data
            self.ax.plot(x_data, y_data, "b-", label="Raw Data", linewidth=1.2)
            self.ax.set_title(
                f'Raw Data: {self.ylk_data.get("name", "Unknown")} (Baseline calc failed)'
            )

        # Draw anchor points if any (only in baseline view)
        if not self.view_toggle_btn.isChecked():
            self.draw_anchors()

        self.ax.set_xlabel("Wavenumber (cm⁻¹)")
        self.ax.set_ylabel("Absorbance")
        self.ax.legend()
        self.ax.grid(True, alpha=0.3)

        if self.parent_analyzer and self.parent_analyzer.reverse_x_axis:
            self.ax.invert_xaxis()

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
                    x_data, als_baseline, "r--", label="ALS Baseline", linewidth=1.5
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
            return

        try:
            # Calculate ALS baseline
            from modules.baseline import baseline_als
            als_baseline = baseline_als(y_data, lam=lambda_val, p=p_val)
            
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

            # Find the original YLK file path and save using the helper method
            data_index = self.parent_analyzer._find_data_index_by_filename(self.filename)
            if data_index is not None:
                ylk_file_path = self.parent_analyzer.files[data_index]
                
                if save_ylk_file(ylk_file_path, self.ylk_data):
                    full_filename = os.path.basename(ylk_file_path)
                    QMessageBox.information(
                        self, "Success", f"Baseline saved to {full_filename}"
                    )
                    # Update the parent's data
                    self.parent_analyzer.ylk_data_list[data_index] = self.ylk_data
                    
                    # Update any selected data that corresponds to this file
                    if self.filename in self.parent_analyzer.selected_files:
                        selected_index = self.parent_analyzer.selected_files.index(self.filename)
                        df = ylk_to_dataframe(self.ylk_data)
                        if df is not None:
                            self.parent_analyzer.selected_data[selected_index] = df
                    
                    # Automatically close the tab after saving
                    self.close_tab()
                else:
                    QMessageBox.warning(self, "Error", "Failed to save baseline")
            else:
                QMessageBox.warning(self, "Error", f"Could not find file {self.filename}")

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
        self.setGeometry(100, 100, 1200, 800)

        # Data storage
        self.files = []
        self.ylk_data_list = []  # Store YLK data instead of DataFrame
        self.selected_data = []
        self.selected_files = []
        self.reverse_x_axis = False
        self.recent_folders = []
        self.show_baseline_corrected = (
            False  # Toggle for raw vs baseline-corrected data
        )

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

        # File menuß
        file_menu = menubar.addMenu("File")

        open_folder_action = QAction("Open Folder", self)
        open_folder_action.triggered.connect(self.select_folder)
        file_menu.addAction(open_folder_action)

        file_menu.addSeparator()

        # Recent folders submenu
        self.recent_menu = file_menu.addMenu("Open Recent")
        self.update_recent_menu()
        
        file_menu.addSeparator()
        
        # Export CSV action
        export_csv_action = QAction("Export Current Graph as CSV", self)
        export_csv_action.triggered.connect(self.export_current_graph_csv)
        file_menu.addAction(export_csv_action)

        # Tools menu
        tools_menu = menubar.addMenu("Tools")

        self.reverse_action = QAction("Reverse X-axis", self)
        self.reverse_action.setCheckable(True)
        self.reverse_action.triggered.connect(self.on_reverse_changed)
        tools_menu.addAction(self.reverse_action)

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

        # All files list (no context menu)
        file_layout.addWidget(QLabel("All Files:"))
        self.file_listbox = QListWidget()
        self.file_listbox.itemDoubleClicked.connect(self.on_file_double_click)
        file_layout.addWidget(self.file_listbox)

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

        # Selected files list with right-click menu
        file_layout.addWidget(QLabel("Selected for Analysis:"))
        self.selected_listbox = QListWidget()
        self.selected_listbox.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.selected_listbox.customContextMenuRequested.connect(
            self.show_selected_context_menu
        )
        self.selected_listbox.itemDoubleClicked.connect(self.on_selected_double_click)
        file_layout.addWidget(self.selected_listbox)

        left_layout.addWidget(file_group)

        # Analysis buttons
        analysis_group = QGroupBox("Analysis")
        analysis_layout = QVBoxLayout(analysis_group)

        self.plot_btn = QPushButton("Plot Spectra")
        self.plot_btn.clicked.connect(self.plot_spectra)
        analysis_layout.addWidget(self.plot_btn)

        # Toggle button for raw/baseline-corrected data
        self.data_toggle_btn = QPushButton("Show Baseline-Corrected Data")
        self.data_toggle_btn.clicked.connect(self.toggle_data_display)
        self.data_toggle_btn.setCheckable(True)
        analysis_layout.addWidget(self.data_toggle_btn)

        self.corr_btn = QPushButton("Calculate Correlation")
        self.corr_btn.clicked.connect(self.calculate_correlation)
        analysis_layout.addWidget(self.corr_btn)

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

    def create_baseline_for_file(self, filename):
        """Create a new baseline tab for the specified file"""
        # Find the YLK data for this file using the helper method
        data_index = self._find_data_index_by_filename(filename)
        if data_index is None:
            QMessageBox.warning(
                self, "Error", f"Could not find data for file {filename}"
            )
            return
        
        ylk_data = self.ylk_data_list[data_index]

        # Create baseline creation tab
        baseline_tab = BaselineCreationTab(ylk_data, filename, self)
        tab_name = f"Baseline: {filename}"
        tab_index = self.tab_widget.addTab(baseline_tab, tab_name)
        self.tab_widget.setCurrentIndex(tab_index)

    def add_to_selected(self):
        """Add selected files to analysis list"""
        selected_items = self.file_listbox.selectedItems()
        items_to_remove = []

        for item in selected_items:
            filename = item.text()
            if filename not in self.selected_files:
                index = self.file_listbox.row(item)
                self.selected_files.append(filename)
                # Convert YLK data to DataFrame for analysis
                ylk_data = self.ylk_data_list[index]
                df = ylk_to_dataframe(ylk_data)
                if df is not None:
                    self.selected_data.append(df)
                    self.selected_listbox.addItem(filename)
                    items_to_remove.append(item)

        # Remove selected files from All Files list
        for item in items_to_remove:
            row = self.file_listbox.row(item)
            self.file_listbox.takeItem(row)

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

        # Rebuild All Files list to maintain proper order
        self._rebuild_all_files_list()
        
        # Update main plot
        self.plot_spectra()

    def clear_selected(self):
        """Clear analysis list"""
        self.selected_listbox.clear()
        self.selected_files.clear()
        self.selected_data.clear()

        # Rebuild All Files list to maintain proper order
        self._rebuild_all_files_list()

        # Clear main plot
        self.ax.clear()
        self.ax.set_xlabel("Wavenumber (cm⁻¹)")
        self.ax.set_ylabel("Absorbance")
        self.ax.set_title("Select files to display spectra")
        self.canvas.draw()

    def on_file_double_click(self, item):
        """Handle double-click on file listbox to move file to selected"""
        filename = item.text()
        if filename not in self.selected_files:
            # Find the correct index in the master data list
            data_index = self._find_data_index_by_filename(filename)
            if data_index is not None:
                self.selected_files.append(filename)
                # Convert YLK data to DataFrame for analysis
                ylk_data = self.ylk_data_list[data_index]
                df = ylk_to_dataframe(ylk_data)
                if df is not None:
                    self.selected_data.append(df)
                    self.selected_listbox.addItem(filename)
                    # Rebuild All Files list to hide selected files
                    self._rebuild_all_files_list()
                    self.plot_spectra()

    def on_selected_double_click(self, item):
        """Handle double-click on selected listbox to remove file from selected"""
        filename = item.text()
        row = self.selected_listbox.row(item)
        self.selected_listbox.takeItem(row)
        if filename in self.selected_files:
            file_index = self.selected_files.index(filename)
            self.selected_files.pop(file_index)
            self.selected_data.pop(file_index)
            # Rebuild All Files list to maintain proper order
            self._rebuild_all_files_list()
            self.plot_spectra()

    def _find_data_index_by_filename(self, filename):
        """Find the index in ylk_data_list for a given display filename"""
        for i, ylk_data in enumerate(self.ylk_data_list):
            # Compare with the display name (without .ylk extension)
            if ylk_data.get("name", "") == filename:
                return i
        return None

    def _rebuild_all_files_list(self):
        """Rebuild the All Files list in proper order, excluding selected files"""
        self.file_listbox.clear()
        for ylk_data in self.ylk_data_list:
            display_name = ylk_data.get("name", "")
            if display_name and display_name not in self.selected_files:
                self.file_listbox.addItem(display_name)

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
            self, 
            "Export CSV", 
            default_filename, 
            "CSV files (*.csv);;All files (*.*)"
        )
        
        if not file_path:
            return

        try:
            import csv
            import pandas as pd
            
            # Collect all data first
            all_data = {}
            
            for filename, df in zip(self.selected_files, self.selected_data):
                # Get raw data
                wavenumber = df['wavenumber'].values
                raw_absorbance = df['absorbance'].values
                
                # Calculate corrected data using baseline if available
                data_index = self._find_data_index_by_filename(filename)
                corrected_absorbance = raw_absorbance.copy()  # Default to raw
                
                if data_index is not None:
                    ylk_data = self.ylk_data_list[data_index]
                    baseline_data = ylk_data.get("baseline", {})
                    
                    if baseline_data.get("x") and baseline_data.get("y"):
                        # Use saved baseline
                        baseline_x = np.array(baseline_data["x"])
                        baseline_y = np.array(baseline_data["y"])
                        
                        # Interpolate baseline to match raw data x-values if needed
                        if len(baseline_x) != len(wavenumber) or not np.allclose(baseline_x, wavenumber):
                            from scipy.interpolate import interp1d
                            baseline_func = interp1d(baseline_x, baseline_y, kind='linear', fill_value='extrapolate')
                            baseline_interpolated = baseline_func(wavenumber)
                        else:
                            baseline_interpolated = baseline_y
                        
                        corrected_absorbance = raw_absorbance - baseline_interpolated
                
                # Clean filename for column names
                clean_filename = filename.replace(" ", "_").replace(".", "_").replace("-", "_")
                
                # Store data
                all_data[f'wavenumber_{clean_filename}'] = wavenumber
                all_data[f'{clean_filename}_raw'] = raw_absorbance
                all_data[f'{clean_filename}_corrected'] = corrected_absorbance

            # Create DataFrame and save
            if all_data:
                df_export = pd.DataFrame(all_data)
                df_export.to_csv(file_path, index=False)
                
                QMessageBox.information(
                    self, 
                    "Export Complete", 
                    f"Data exported successfully to:\n{file_path}\n\n"
                    f"Exported {len(self.selected_files)} file(s).\n"
                    f"Each file has wavenumber, raw, and corrected columns."
                )
            else:
                QMessageBox.warning(self, "Export Error", "No data to export")

        except Exception as e:
            QMessageBox.critical(
                self, "Export Error", f"Failed to export CSV file:\n{str(e)}"
            )

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(
            self, "Select folder with .jws files"
        )
        if folder_path:
            self.add_recent_folder(folder_path)
            self.process_folder(folder_path)

    def process_folder(self, folder_path):
        """Process folder: convert JWS files to YLK format and load existing YLK files"""
        try:
            # Create converted_ylk subfolder
            ylk_folder = os.path.join(folder_path, "converted_ylk")
            os.makedirs(ylk_folder, exist_ok=True)

            # Clear lists
            self.file_listbox.clear()
            self.ylk_data_list = []
            self.files = []
            self.clear_selected()

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
            for ylk_file in processed_files:
                try:
                    ylk_data = load_ylk_file(ylk_file)
                    if ylk_data:
                        self.ylk_data_list.append(ylk_data)
                        self.files.append(ylk_file)
                        # Display filename without .ylk extension
                        display_name = os.path.basename(ylk_file).replace(".ylk", "")
                        self.file_listbox.addItem(display_name)
                except Exception as e:
                    QMessageBox.warning(
                        self, "Warning", f"Unable to load file {ylk_file}: {str(e)}"
                    )

            QMessageBox.information(
                self, "Complete", f"Processed {len(processed_files)} files"
            )

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error processing folder: {str(e)}")

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

        # Plot each selected spectrum
        for i, df in enumerate(self.selected_data):
            filename = self.selected_files[i]  # Already without .ylk extension

            if self.show_baseline_corrected:
                # Show baseline-corrected data
                # Find the corresponding YLK data to get baseline-corrected values
                ylk_data = None
                # Find the YLK data by matching filename
                for file_path_idx, file_path in enumerate(self.files):
                    if os.path.basename(file_path).replace(".ylk", "") == filename:
                        ylk_data = self.ylk_data_list[file_path_idx]
                        break

                if (
                    ylk_data
                    and ylk_data.get("baseline", {}).get("x")
                    and ylk_data.get("baseline", {}).get("y")
                ):
                    # Use baseline-corrected data if available
                    raw_x = np.array(ylk_data["raw_data"]["x"])
                    raw_y = np.array(ylk_data["raw_data"]["y"])
                    baseline_y = np.array(ylk_data["baseline"]["y"])
                    corrected_y = raw_y - baseline_y

                    self.ax.plot(
                        raw_x,
                        corrected_y,
                        label=f"{filename} (Baseline-corrected)",
                        linewidth=1.2,
                    )
                else:
                    # Fall back to raw data if no baseline available
                    pre_df = preprocess_data(df, normalize=True)
                    self.ax.plot(
                        pre_df["wavenumber"],
                        pre_df["absorbance"],
                        label=f"{filename} (Raw - no baseline)",
                        linewidth=1.2,
                        linestyle="--",
                    )
            else:
                # Show raw normalized data
                pre_df = preprocess_data(df, normalize=True)
                self.ax.plot(
                    pre_df["wavenumber"],
                    pre_df["absorbance"],
                    label=filename,
                    linewidth=1.2,
                )

        # Set labels and title
        self.ax.set_xlabel("Wavenumber (cm⁻¹)")
        if self.show_baseline_corrected:
            self.ax.set_ylabel("Baseline-corrected Absorbance")
            self.ax.set_title("FTIR Spectra (Baseline-corrected)")
        else:
            self.ax.set_ylabel("Normalized Absorbance")
            self.ax.set_title("FTIR Spectra (Raw)")

        self.ax.grid(True, alpha=0.3)

        # Apply x-axis reversal if enabled
        if self.reverse_x_axis:
            self.ax.invert_xaxis()

        # Add legend if there are multiple spectra
        if len(self.selected_data) > 1:
            self.ax.legend()

        # Update canvas
        self.canvas.draw()

    def calculate_correlation(self):
        if len(self.selected_data) < 2 or len(self.selected_data) > 5:
            QMessageBox.warning(self, "警告", "請選擇2-5筆資料進行相關性分析")
            return

        # Normalize data for correlation analysis
        normalized_data = [
            preprocess_data(df, normalize=True) for df in self.selected_data
        ]
        corr_matrix = calculate_correlation_matrix(normalized_data)
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
