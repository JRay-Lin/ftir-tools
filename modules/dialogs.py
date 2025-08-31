"""
Dialogs Module for FTIR Spectroscopy

Contains dialog functions that are loaded on-demand to improve application startup performance.
These dialogs are only loaded when the user requests them from the Help menu.
"""

from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices

from modules.version import get_app_info


def show_version_dialog(parent):
    """
    Show Version dialog with detailed version information
    
    Parameters:
    parent: Parent widget for the dialog
    """
    try:
        app_info = get_app_info()
        version_text = f"""
<h2>Version Information</h2>
<table style="margin: 10px 0;">
<tr><td style="padding-right: 20px; font-weight: bold;">Application:</td><td>{app_info['name']}</td></tr>
<tr><td style="padding-right: 20px; font-weight: bold;">Version:</td><td>{app_info['version']}</td></tr>
<tr><td style="padding-right: 20px; font-weight: bold;">Python Requirement:</td><td>{app_info['python_requirement']}</td></tr>
</table>
<br>
<h3>Key Features:</h3>
<ul>
<li>FTIR spectral data analysis from .jws files</li>
<li>Interactive baseline correction with ALS algorithm</li>
<li>Real-time visualization with matplotlib</li>
<li>Data export to CSV and PNG formats</li>
<li>Correlation analysis between spectra</li>
</ul>
<br>
<p><b>GitHub:</b> <a href="https://github.com/JRay-Lin/ftir-tools">https://github.com/JRay-Lin/ftir-tools</a></p>
"""

        msg_box = QMessageBox(parent)
        msg_box.setWindowTitle("Version - FTIR Tools")
        msg_box.setTextFormat(Qt.TextFormat.RichText)
        msg_box.setText(version_text)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        # Make the dialog larger for better readability
        msg_box.resize(450, 300)
        msg_box.exec()

    except Exception as e:
        QMessageBox.warning(
            parent, "Error", f"Could not display version information: {str(e)}"
        )


def show_about_dialog(parent):
    """
    Show About dialog with application information
    
    Parameters:
    parent: Parent widget for the dialog
    """
    try:
        app_info = get_app_info()
        about_text = f"""
<h2>{app_info['name']}</h2>
<p><b>Version:</b> {app_info['version']}</p>
<p><b>Description:</b> {app_info['description']}</p>
<p><b>Python Requirement:</b> {app_info['python_requirement']}</p>
<br>
<p>A tool for analyzing FTIR (Fourier Transform Infrared Spectroscopy) data from .jws files.</p>
<br>
<p><b>GitHub Repository:</b> <a href="https://github.com/JRay-Lin/ftir-tools">https://github.com/JRay-Lin/ftir-tools</a></p>
<br>
<p>© 2024 FTIR Tools Project</p>
"""

        msg_box = QMessageBox(parent)
        msg_box.setWindowTitle("About FTIR Tools")
        msg_box.setTextFormat(Qt.TextFormat.RichText)
        msg_box.setText(about_text)
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()

    except Exception as e:
        QMessageBox.warning(
            parent, "Error", f"Could not display application information: {str(e)}"
        )


def open_manual(parent):
    """
    Open user manual in the default browser
    
    Parameters:
    parent: Parent widget for error dialogs
    """
    try:
        manual_url = "https://github.com/JRay-Lin/ftir-tools/blob/master/docs/manual_zh_tw.md"
        QDesktopServices.openUrl(QUrl(manual_url))
    except Exception as e:
        QMessageBox.warning(parent, "Error", f"Could not open manual: {str(e)}")


def open_absorption_table(parent):
    """
    Open infrared spectroscopy absorption table in a new tab
    
    Parameters:
    parent: Parent widget (should be FTIRAnalyzer instance)
    
    Returns:
    bool: True if successful, False if tab already exists
    """
    try:
        # Check if absorption table tab already exists
        for i in range(parent.tab_widget.count()):
            if parent.tab_widget.tabText(i) == "Absorption Table":
                parent.tab_widget.setCurrentIndex(i)
                return False

        # Lazy import the AbsorptionTableTab class
        from modules.reference_data import AbsorptionTableTab

        # Create new absorption table tab
        absorption_tab = AbsorptionTableTab(parent)
        tab_index = parent.tab_widget.addTab(absorption_tab, "Absorption Table")
        parent.tab_widget.setCurrentIndex(tab_index)
        return True

    except Exception as e:
        QMessageBox.warning(
            parent, "Error", f"Could not create absorption table: {str(e)}"
        )
        return False