import sys
import os
import pandas as pd
import numpy as np
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
    QMenuBar,
    QMenu,
    QTabWidget,
    QListWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QGroupBox,
)
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QAction

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
from modules.version import get_app_info


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

        # Initialize attributes for crosshairs and coordinate display
        self.crosshair_h = None
        self.crosshair_v = None
        self.coord_text = None
        self.hover_annotation = None

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

        # Help menu
        help_menu = menubar.addMenu("Help")
        assert help_menu is not None

        version_action = QAction("Version", self)
        version_action.triggered.connect(self.show_version_dialog)
        help_menu.addAction(version_action)

        absorption_table_action = QAction("Absorption Table", self)
        absorption_table_action.triggered.connect(self.open_absorption_table)
        help_menu.addAction(absorption_table_action)

        help_menu.addSeparator()

        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

        manual_action = QAction("User Manual", self)
        manual_action.triggered.connect(self.open_manual)
        help_menu.addAction(manual_action)

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

        # Selected files list with right-click menu and checkboxes
        # Create horizontal layout for label and clear button
        selected_header_layout = QHBoxLayout()
        selected_header_layout.addWidget(QLabel("Selected for Analysis:"))

        # Add clear button on the right side
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_selected)
        selected_header_layout.addWidget(clear_btn)

        file_layout.addLayout(selected_header_layout)
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
            if hasattr(self, "crosshair_h") and self.crosshair_h is not None:
                try:
                    self.crosshair_h.set_visible(False)
                    # Don't try to remove axhline/axvline, just hide them
                except:
                    pass
                self.crosshair_h = None
            if hasattr(self, "crosshair_v") and self.crosshair_v is not None:
                try:
                    self.crosshair_v.set_visible(False)
                except:
                    pass
                self.crosshair_v = None
            if hasattr(self, "coord_text") and self.coord_text is not None:
                try:
                    self.coord_text.remove()
                except:
                    pass
                self.coord_text = None
            if hasattr(self, "hover_annotation") and self.hover_annotation is not None:
                try:
                    self.hover_annotation.remove()
                except:
                    pass
                self.hover_annotation = None
            self.canvas.draw()

    def on_main_plot_hover(self, event):
        """Handle mouse hover over main plot to show coordinates with crosshairs"""
        from modules.ui_helpers import on_main_plot_hover

        on_main_plot_hover(self, event)

    def _hide_crosshairs(self):
        """Hide crosshairs and coordinate text"""
        from modules.ui_helpers import hide_crosshairs

        hide_crosshairs(self)

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

    def selected_list_mouse_press(self, e):
        """Custom mouse press handler for selected files list to track click position"""
        from modules.ui_helpers import selected_list_mouse_press

        selected_list_mouse_press(self.selected_listbox, e)
        # Store the click result for this instance
        self._last_click_on_checkbox = getattr(
            self.selected_listbox, "_last_click_on_checkbox", False
        )

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
                    lambda _checked, path=folder: self.process_folder(path)
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

        # Extract just the filename part from the display name if needed
        if " (" in filename:
            # Handle display format "filename (folder)"
            actual_filename = filename.split(" (")[0]
        else:
            actual_filename = filename

        # Search through selected files to find matching data
        for selected_file_key in self.selected_files:
            folder_path, file_name = selected_file_key
            if file_name == actual_filename or filename.startswith(file_name):
                # Found the file, now get the YLK data from folders
                if folder_path in self.folders:
                    folder_data = self.folders[folder_path]
                    for j, file_path in enumerate(folder_data["files"]):
                        if os.path.basename(file_path).replace(".ylk", "") == file_name:
                            ylk_data = folder_data["ylk_data"][j]
                            break
                if ylk_data:
                    break

        if ylk_data is None:
            QMessageBox.warning(
                self, "Error", f"Could not find data for file {filename}"
            )
            return

        # Create baseline creation tab (lazy import)
        from modules.gui_components import BaselineCreationTab

        baseline_tab = BaselineCreationTab(ylk_data, actual_filename, self)
        tab_name = f"Baseline: {actual_filename}"
        tab_index = self.tab_widget.addTab(baseline_tab, tab_name)
        self.tab_widget.setCurrentIndex(tab_index)

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
        from modules.ui_helpers import export_current_graph_csv

        export_current_graph_csv(self)

    def export_current_graph_png(self):
        """Export current graph as PNG file"""
        from modules.ui_helpers import export_current_graph_png

        export_current_graph_png(self)

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
        from modules.ui_helpers import get_selected_wavenumber_ranges

        reference_ranges = get_selected_wavenumber_ranges(
            self.selected_files, self.folders
        )

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
                        from modules.ui_helpers import (
                            get_file_wavenumber_range,
                            is_similar_range,
                        )

                        file_range = get_file_wavenumber_range(ylk_data)
                        if file_range and is_similar_range(
                            file_range, reference_ranges
                        ):
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
                            # Interpolate baseline to match raw data x-values using numpy
                            baseline_y_interp = np.interp(raw_x, baseline_x, baseline_y)
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

    def show_version_dialog(self):
        """Show Version dialog with detailed version information"""
        from modules.dialogs import show_version_dialog

        show_version_dialog(self)

    def show_about_dialog(self):
        """Show About dialog with application information"""
        from modules.dialogs import show_about_dialog

        show_about_dialog(self)

    def open_manual(self):
        """Open user manual in the default browser"""
        from modules.dialogs import open_manual

        open_manual(self)

    def open_absorption_table(self):
        """Open infrared spectroscopy absorption table in a new tab"""
        from modules.dialogs import open_absorption_table

        open_absorption_table(self)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    analyzer = FTIRAnalyzer()
    analyzer.show()
    sys.exit(app.exec())
