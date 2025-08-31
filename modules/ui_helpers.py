"""
UI Helpers Module for FTIR Spectroscopy

Contains utility functions for UI operations that can be loaded on-demand
to improve application startup performance.
"""

import os
import numpy as np
import pandas as pd
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from PyQt6.QtCore import Qt


def get_selected_wavenumber_ranges(selected_files, folders):
    """
    Get wavenumber ranges from currently selected files
    
    Parameters:
    selected_files: list of (folder_path, filename) tuples
    folders: dict containing folder data
    
    Returns:
    list: List of wavenumber range tuples
    """
    ranges = []
    for file_key in selected_files:
        folder_path, filename = file_key
        if folder_path in folders:
            folder_data = folders[folder_path]
            for i, file_path in enumerate(folder_data["files"]):
                if os.path.basename(file_path).replace(".ylk", "") == filename:
                    ylk_data = folder_data["ylk_data"][i]
                    file_range = get_file_wavenumber_range(ylk_data)
                    if file_range:
                        ranges.append(file_range)
                    break
    return ranges


def get_file_wavenumber_range(ylk_data):
    """
    Get wavenumber range (min, max) from YLK data
    
    Parameters:
    ylk_data: YLK data dictionary
    
    Returns:
    tuple: (min_wavenumber, max_wavenumber) or None if error
    """
    try:
        raw_data = ylk_data.get("raw_data", {})
        x_data = raw_data.get("x", [])
        if x_data:
            return (min(x_data), max(x_data))
    except Exception:
        pass
    return None


def is_similar_range(file_range, reference_ranges, tolerance=50):
    """
    Check if a file's wavenumber range is similar to any reference range
    
    Parameters:
    file_range: tuple of (min, max) wavenumbers for the file
    reference_ranges: list of reference (min, max) tuples
    tolerance: tolerance in wavenumber units for similarity
    
    Returns:
    bool: True if similar to any reference range
    """
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


def hide_crosshairs(analyzer):
    """
    Hide crosshairs and coordinate text
    
    Parameters:
    analyzer: FTIRAnalyzer instance
    """
    if hasattr(analyzer, "crosshair_h"):
        analyzer.crosshair_h.set_visible(False)
        analyzer.crosshair_v.set_visible(False)
    if hasattr(analyzer, "coord_text"):
        analyzer.coord_text.set_visible(False)
    analyzer.canvas.draw()


def export_current_graph_csv(analyzer):
    """
    Export current graph data (raw and corrected) to CSV file
    
    Parameters:
    analyzer: FTIRAnalyzer instance
    """
    if not analyzer.selected_data or not analyzer.selected_files:
        QMessageBox.information(
            analyzer, "No Data", "No files are currently selected for analysis."
        )
        return

    # Get save file path
    default_filename = "ftir_export.csv"
    file_path, _ = QFileDialog.getSaveFileName(
        analyzer, "Export CSV", default_filename, "CSV files (*.csv);;All files (*.*)"
    )

    if not file_path:
        return

    try:
        import csv

        # Collect all data first
        all_data = {}

        for file_key, df in zip(analyzer.selected_files, analyzer.selected_data):
            # Get raw data
            wavenumber = df["wavenumber"].values
            raw_absorbance = df["absorbance"].values

            # Calculate corrected data using baseline if available
            folder_path, file_name = file_key
            corrected_absorbance = raw_absorbance.copy()  # Default to raw
            ylk_data = None

            # Find YLK data for this file
            if folder_path in analyzer.folders:
                folder_data = analyzer.folders[folder_path]
                for j, file_path_full in enumerate(folder_data["files"]):
                    if os.path.basename(file_path_full).replace(".ylk", "") == file_name:
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
                                fill_value=np.nan,
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
                analyzer,
                "Export Complete",
                f"Data exported successfully to:\\n{file_path}\\n\\n"
                f"Exported {len(analyzer.selected_files)} file(s).\\n"
                f"Each file has wavenumber, raw, and corrected columns.",
            )
        else:
            QMessageBox.warning(analyzer, "Export Error", "No data to export")

    except Exception as e:
        QMessageBox.critical(
            analyzer, "Export Error", f"Failed to export CSV file:\\n{str(e)}"
        )


def export_current_graph_png(analyzer):
    """
    Export current graph as PNG file
    
    Parameters:
    analyzer: FTIRAnalyzer instance
    """
    if not analyzer.selected_data or not analyzer.selected_files:
        QMessageBox.information(
            analyzer, "No Data", "No files are currently selected for analysis."
        )
        return

    # Get save file path
    default_filename = "ftir_graph.png"
    file_path, _ = QFileDialog.getSaveFileName(
        analyzer, "Export PNG", default_filename, "PNG files (*.png);;All files (*.*)"
    )

    if not file_path:
        return

    try:
        # Save the current figure
        analyzer.figure.savefig(
            file_path,
            dpi=300,
            bbox_inches="tight",
            facecolor="white",
            edgecolor="none",
            format="png",
        )

        QMessageBox.information(
            analyzer, "Export Complete", f"Graph exported successfully to:\\n{file_path}"
        )

    except Exception as e:
        QMessageBox.critical(
            analyzer, "Export Error", f"Failed to export PNG file:\\n{str(e)}"
        )


def on_main_plot_hover(analyzer, event):
    """
    Handle mouse hover over main plot to show coordinates with crosshairs
    
    Parameters:
    analyzer: FTIRAnalyzer instance
    event: matplotlib mouse event
    """
    if event.inaxes != analyzer.ax or not analyzer.selected_data:
        # Hide crosshairs if mouse is outside plot area
        hide_crosshairs(analyzer)
        return

    if event.xdata is None or event.ydata is None:
        return

    # Store current axis limits to prevent shifting
    xlim = analyzer.ax.get_xlim()
    ylim = analyzer.ax.get_ylim()

    # Create crosshairs if they don't exist
    if not hasattr(analyzer, "crosshair_h"):
        analyzer.crosshair_h = analyzer.ax.axhline(
            y=event.ydata, color="gray", linestyle="--", alpha=0.7, linewidth=0.8
        )
        analyzer.crosshair_v = analyzer.ax.axvline(
            x=event.xdata, color="gray", linestyle="--", alpha=0.7, linewidth=0.8
        )
    else:
        # Update positions
        analyzer.crosshair_h.set_ydata([event.ydata, event.ydata])
        analyzer.crosshair_v.set_xdata([event.xdata, event.xdata])

    # Make visible
    analyzer.crosshair_h.set_visible(True)
    analyzer.crosshair_v.set_visible(True)

    # Update coordinate text at bottom of plot
    wavenumber = event.xdata
    absorbance = event.ydata

    # Create or update text display
    if not hasattr(analyzer, "coord_text"):
        # Position text at bottom left of the plot
        analyzer.coord_text = analyzer.ax.text(
            0.02,
            0.02,
            "",
            transform=analyzer.ax.transAxes,
            fontsize=9,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.8),
        )

    # Format text
    text = f"Wavenumber: {wavenumber:.1f} cm⁻¹  |  Absorbance: {absorbance:.4f}"
    analyzer.coord_text.set_text(text)
    analyzer.coord_text.set_visible(True)

    # Restore axis limits to prevent shifting
    analyzer.ax.set_xlim(xlim)
    analyzer.ax.set_ylim(ylim)

    analyzer.canvas.draw()


def selected_list_mouse_press(listbox, event):
    """
    Custom mouse press handler for selected files list to track click position
    
    Parameters:
    listbox: QListWidget instance
    event: mouse event
    """
    # Store the original mousePressEvent behavior
    from PyQt6.QtWidgets import QListWidget
    original_mouse_press = QListWidget.mousePressEvent
    original_mouse_press(listbox, event)

    # Store click position for double-click detection
    item = listbox.itemAt(event.pos())
    if item:
        # Check if click is on the checkbox area (left part of item)
        item_rect = listbox.visualItemRect(item)
        checkbox_area_width = 20  # Approximate checkbox width
        listbox._last_click_on_checkbox = (
            event.pos().x() < item_rect.left() + checkbox_area_width
        )
    else:
        listbox._last_click_on_checkbox = False