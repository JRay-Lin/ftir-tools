"""
Reference Data Module for FTIR Spectroscopy

Contains infrared spectroscopy reference data and related UI components.
This module is loaded only when the absorption table is accessed to improve startup performance.
"""

from PyQt6.QtWidgets import (
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont


def get_ir_absorption_data():
    """
    Get infrared spectroscopy absorption reference data.
    
    Returns:
    list: List of tuples containing (peak_position, group, class, peak_details)
    """
    return [
        (
            "3584 - 3700",
            "O-H stretching (free)",
            "alcohol, phenol",
            "strong, sharp",
        ),
        (
            "3500 - 3700",
            "O-H stretching (free)",
            "phenol, alcohol",
            "strong, sharp",
        ),
        (
            "3200 - 3550",
            "O-H stretching (hydrogen-bonded)",
            "alcohol",
            "strong, broad",
        ),
        ("3500", "N-H stretching", "primary amine", "medium, sharp"),
        (
            "3300 - 3400",
            "N-H stretching",
            "aliphatic primary amine",
            "medium, sharp",
        ),
        ("3310 - 3350", "N-H stretching", "secondary amine", "medium, sharp"),
        (
            " 3050 - 3100",
            "C-H stretching (aromatic)",
            "aromatic hydrocarbon",
            "medium to weak",
        ),
        (
            " 2500 - 3300",
            "O-H stretching",
            "carboxylic acid",
            "very strong, very broad",
        ),
        ("2700 - 3200", "O-H stretching", "alcohol", "weak, broad"),
        ("2800 - 3000", "N-H stretching", "amine salt", "strong, broad"),
        ("3267 - 3333", "C-H stretching", "alkyne", "strong, sharp"),
        ("3300", "≡C–H stretching", "terminal alkyne", "strong, sharp"),
        ("3000 - 3100", "C-H stretching", "alkene", "medium"),
        ("2840 - 3000", "C-H stretching", "alkane", "medium"),
        (
            "2695 - 2830",
            "C-H stretching",
            "aldehyde",
            "medium (plus weak Fermi doublet at 2720, 2820)",
        ),
        ("2550 - 2600", "S-H stretching", "thiol", "weak"),
        ("2349", "O=C=O stretching", "carbon dioxide", "strong"),
        ("2250 - 2275", "N=C=O stretching", "isocyanate", "strong, broad"),
        ("2222 - 2260", "C≡N stretching", "nitrile", "weak to medium, sharp"),
        ("2190 - 2260", "C≡C stretching", "alkyne", "weak"),
        (
            "2200 - 2250",
            "C≡C stretching",
            "terminal alkyne",
            "weak to medium, sharp",
        ),
        ("2140 - 2175", "S-C≡N stretching", "thiocyanate", "strong"),
        ("2120 - 2160", "N=N=N stretching", "azide", "strong"),
        ("2150", "C=C=O stretching", "ketene", "typically sharp, medium"),
        ("2120 - 2145", "N=C=N stretching", "carbodiimide", "strong, sharp"),
        ("2100 - 2140", "C≡C stretching", "alkyne", "weak"),
        ("1990 - 2140", "N=C=S stretching", "isothiocyanate", "strong, sharp"),
        ("1900 - 2000", "C=C=C stretching", "allene", "medium, sharp"),
        ("2000", "C=C=N stretching", "ketenimine", "weak to medium, sharp"),
        ("1650 - 2000", "C-H bending", "aromatic compound", "weak"),
        ("1818", "C=O stretching", "anhydride", "strong, sharp"),
        ("1785 - 1815", "C=O stretching", "acid halide", "strong, sharp"),
        (
            "1770 - 1800",
            "C=O stretching",
            "conjugated acid halide",
            "strong, sharp",
        ),
        ("1775", "C=O stretching", "conjugated anhydride", "strong, sharp"),
        ("1770 - 1780", "C=O stretching", "vinyl / phenyl ester", "strong, sharp"),
        ("1760", "C=O stretching", "carboxylic acid", "strong, sharp"),
        ("1735 - 1750", "C=O stretching", "ester", "strong, sharp"),
        ("1735 - 1750", "C=O stretching", "δ-lactone", "strong, sharp"),
        ("1745", "C=O stretching", "cyclopentanone", "strong, sharp"),
        ("1720 - 1740", "C=O stretching", "aldehyde", "strong, sharp"),
        ("1715 - 1730", "C=O stretching", "α,β-unsaturated ester", "strong, sharp"),
        ("1705 - 1725", "C=O stretching", "aliphatic ketone", "strong, sharp"),
        ("1706 - 1720", "C=O stretching", "carboxylic acid", "strong, broad"),
        ("1680 - 1710", "C=O stretching", "conjugated acid", "strong"),
        ("1685 - 1710", "C=O stretching", "conjugated aldehyde", "strong, sharp"),
        ("1690", "C=O stretching", "primary amide", "strong, sharp"),
        ("1640 - 1690", "C=N stretching", "imine / oxime", "strong"),
        ("1666 - 1685", "C=O stretching", "conjugated ketone", "strong, sharp"),
        ("1680", "C=O stretching", "secondary amide", "strong, sharp"),
        ("1680", "C=O stretching", "tertiary amide", "strong, sharp"),
        ("1650", "C=O stretching", "δ-lactam", "strong, sharp"),
        ("1668 - 1678", "C=C stretching", "alkene", "weak"),
        ("1665 - 1675", "C=C stretching", "alkene", "weak"),
        ("1626 - 1662", "C=C stretching", "alkene", "medium"),
        ("1648 - 1658", "C=C stretching", "alkene", "medium"),
        ("1600 - 1650", "C=C stretching", "conjugated alkene", "medium"),
        ("1580 - 1650", "N-H bending", "amine", "medium"),
        ("1566 - 1650", "C=C stretching", "cyclic alkene", "medium"),
        ("1638 - 1648", "C=C stretching", "alkene", "strong"),
        ("1610 - 1620", "C=C stretching", "α,β-unsaturated ketone", "strong"),
        ("1500 - 1550", "N-O stretching", "nitro compound", "strong"),
        ("1465", "C-H bending", "alkane", "medium"),
        ("1450", "C-H bending", "alkane", "medium"),
        ("1380 - 1390", "C-H bending", "aldehyde", "medium"),
        ("1380 - 1385", "C-H bending", "alkane", "medium"),
        ("1395 - 1440", "O-H bending", "carboxylic acid", "medium"),
        ("1330 - 1420", "O-H bending", "alcohol", "medium"),
        ("1380 - 1415", "S=O stretching", "sulfate", "strong, sharp"),
        ("1380 - 1410", "S=O stretching", "sulfonyl chloride", "strong, sharp"),
        ("1000 - 1400", "C-F stretching", "fluoro compound", "strong"),
        ("1310 - 1390", "O-H bending", "phenol", "medium"),
        ("1335 - 1372", "S=O stretching", "sulfonate", "strong, sharp"),
        ("1335 - 1370", "S=O stretching", "sulfonamide", "strong, sharp"),
        ("1342 - 1350", "S=O stretching", "sulfonic acid", "strong, sharp"),
        ("1300 - 1350", "S=O stretching", "sulfone", "strong, sharp"),
        ("1266 - 1342", "C-N stretching", "aromatic amine", "strong"),
        ("1250 - 1310", "C-O stretching", "aromatic ester", "strong, sharp"),
        ("1200 - 1275", "C-O stretching", "alkyl aryl ether", "strong, sharp"),
        ("1020 - 1250", "C-N stretching", "amine", "medium"),
        ("1200 - 1225", "C-O stretching", "vinyl ether", "strong, sharp"),
        ("1163 - 1210", "C-O stretching", "ester", "strong, sharp"),
        ("1124 - 1205", "C-O stretching", "tertiary alcohol", "strong, sharp"),
        ("1085 - 1150", "C-O stretching", "aliphatic ether", "strong, sharp"),
        ("1087 - 1124", "C-O stretching", "secondary alcohol", "strong, sharp"),
        ("1050 - 1085", "C-O stretching", "primary alcohol", "strong, sharp"),
        ("1030 - 1070", "S=O stretching", "sulfoxide", "strong, sharp"),
        ("1040 - 1050", "CO-O-CO stretching", "anhydride", "strong, broad"),
        ("985 - 995", "C=C bending", "allene", "strong, sharp"),
        ("970 - 990", "=C–H out-of-plane bending", "trans-alkene", "strong, sharp"),
        ("960 - 980", "C=C bending", "alkene", "strong, sharp"),
        ("905 - 920", "=C–H out-of-plane bending", "vinyl group", "medium"),
        ("885 - 895", "C=C bending", "alkene", "strong"),
        ("790 - 840", "C=C bending", "alkene", "medium"),
        ("665 - 730", "C=C bending", "alkene", "strong"),
        (
            "650 - 900",
            "C-H out-of-plane bending",
            "aromatic substitution",
            "medium to strong",
        ),
        ("515 - 690", "C-Br stretching", "halo compound", "strong"),
        ("500 - 600", "C-I stretching", "halo compound", "strong"),
        ("700 - 800", "C-Cl stretching", "alkyl halide", "strong, sharp"),
        ("550 - 850", "C-Cl stretching", "halo compound", "strong"),
        ("860 - 900", "C-H bending", "1,2,4-trisubstituted benzene", "strong"),
        ("860 - 900", "C-H bending", "1,3-disubstituted benzene", "strong"),
        ("790 - 830", "C-H bending", "1,4-disubstituted benzene", "strong"),
        ("790 - 830", "C-H bending", "1,2,3,4-tetrasubstituted benzene", "strong"),
        ("760 - 800", "C-H bending", "1,2,3-trisubstituted benzene", "strong"),
        ("735 - 775", "C-H bending", "1,2-disubstituted benzene", "strong"),
        ("730 - 770", "C-H bending", "monosubstituted benzene", "strong"),
        (
            "400 - 600",
            "M–O stretching",
            "metal–oxygen (inorganic)",
            "strong, broad",
        ),
        ("680 - 720", "C-H bending", "benzene derivative", "aromatic out-of-plane"),
    ]


class AbsorptionTableTab(QWidget):
    """
    Tab widget displaying infrared spectroscopy absorption reference data.
    This widget is created on-demand to improve application startup performance.
    """
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # Title
        title = QLabel("Infrared Spectroscopy Absorption Table")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Subtitle
        subtitle = QLabel(
            'Data from <a href="https://instanano.com/all/characterization/ftir/ftir-functional-group-search/">instanano</a>'
        )
        subtitle.setFont(QFont("Arial", 10))
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setTextFormat(Qt.TextFormat.RichText)
        subtitle.setOpenExternalLinks(True)
        layout.addWidget(subtitle)

        # Create table
        self.table = QTableWidget()
        self.populate_table()
        layout.addWidget(self.table)

        # Close button
        close_btn = QPushButton("Close Tab")
        close_btn.clicked.connect(self.close_tab)
        layout.addWidget(close_btn)

    def populate_table(self):
        """Populate the table with IR absorption reference data"""
        # Get IR absorption data
        ir_data = get_ir_absorption_data()

        # Set up table
        self.table.setRowCount(len(ir_data))
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(
            ["Peak Position", "Group", "Class", "Peak Details"]
        )

        # Populate table
        for row, (group, freq, intensity, notes) in enumerate(ir_data):
            self.table.setItem(row, 0, QTableWidgetItem(group))
            self.table.setItem(row, 1, QTableWidgetItem(freq))
            self.table.setItem(row, 2, QTableWidgetItem(intensity))
            self.table.setItem(row, 3, QTableWidgetItem(notes))

        # Set column widths
        self.table.setColumnWidth(0, 200)
        self.table.setColumnWidth(1, 120)
        self.table.setColumnWidth(2, 80)
        self.table.setColumnWidth(3, 300)

        # Make table read-only and enable sorting
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table.setSortingEnabled(True)

        # Stretch last column
        header = self.table.horizontalHeader()
        if header:
            header.setStretchLastSection(True)

        # Alternate row colors
        self.table.setAlternatingRowColors(True)

    def close_tab(self):
        """Close this tab"""
        parent = self.parent()
        while parent and not isinstance(parent, QTabWidget):
            parent = parent.parent()

        if parent:
            tab_index = parent.indexOf(self)
            if tab_index >= 0:
                parent.removeTab(tab_index)