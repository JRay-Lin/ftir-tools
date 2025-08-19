# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an FTIR (Fourier Transform Infrared Spectroscopy) data analysis tool built in Python with a PyQt6 GUI. The application processes FTIR spectral data from .jws files (Jasco format) and CSV files, providing visualization and correlation analysis capabilities.

## Dependencies and Environment

- **Python Version**: >=3.12
- **Package Manager**: uv (uses uv.lock for dependency management)
- **Key Dependencies**: 
  - matplotlib (>=3.10.5) - plotting and visualization
  - pandas (>=2.3.1) - data manipulation
  - scikit-learn (>=1.7.1) - machine learning utilities
  - scipy - signal processing and statistics
  - PyQt6 (>=6.5.0) - GUI framework

## Core Architecture

The application is structured around a main class `FTIRAnalyzer` in `main.py:227` with a modern tabbed interface:

- **File Processing**: Handles only .jws files, converting them to custom .ylk format
- **Data Format**: Uses YLK (JSON-based) format for storing raw data, baseline, and metadata
- **Data Conversion**: Direct JWS file parsing and fallback to `jws2txt` command
- **Data Preprocessing**: Applies Savitzky-Golay smoothing filter and normalization
- **Visualization**: Embedded matplotlib canvas in main window with PyQt6 integration
- **Analysis**: Pearson correlation analysis between multiple spectra
- **Baseline Creation**: Tabbed interface for creating and saving baselines per file
- **User Interface**: Modern PyQt6 interface with menu bar, context menus, and recent files

## Setup and Common Commands

### Initial Setup

The application now uses PyQt6 for the GUI instead of tkinter. This provides better cross-platform compatibility and more modern UI components.

```bash
# Install dependencies (PyQt6 will be installed automatically)
uv sync
```

### Common Commands

```bash
# Install dependencies
uv sync

# Run the application
uv run python main.py

# Or run directly if environment is activated
python main.py
```

## External Dependencies

The application uses direct JWS file parsing with optional fallback to the `jws2txt` command-line tool for converting Jasco .jws files. The conversion happens in the `convert_jws_to_ylk_direct` method in `modules/file_converter.py`.

## Data Structure - YLK Format

- **Input Files**: .jws files (Jasco format) in selected folders
- **Output Format**: .ylk files (JSON format) in `converted_ylk/` subfolder
- **YLK Structure**:
  ```json
  {
    "name": "spectrum_name",
    "range": [min_wavenumber, max_wavenumber],
    "raw_data": {
      "x": [wavenumber_array],
      "y": [absorbance_array]
    },
    "baseline": {
      "x": [wavenumber_array],
      "y": [baseline_array]
    },
    "metadata": {
      "created": "timestamp",
      "source_file": "original_jws_filename",
      "baseline_params": {...}
    }
  }
  ```
- **Processing Flow**: .jws → .ylk → analysis/plotting with baseline creation

## GUI Components

- **Main Window**: Tabbed interface with embedded matplotlib canvas
- **Menu System**: File menu (Open Folder, Recent), Tools menu (Reverse X-axis)
- **File Management**: 
  - "All Files" list shows available files for selection
  - "Selected for Analysis" list shows files chosen for analysis
  - Files automatically hide/show between lists when selected/deselected
  - Right-click context menu on selected files for baseline creation
- **Data Display Toggle**: Button to switch between raw and baseline-corrected data views
- **Plotting**: Real-time embedded spectral display with automatic updates
- **Baseline Creation**: Individual tabs per file for creating and saving baselines
- **Analysis Tools**: Correlation analysis with matrix visualization
- **Recent Files**: Persistent recent folder history

## Key Functions

- `select_folder()` - Folder selection and file processing
- `process_folder()` - Convert JWS files to YLK format and load existing YLK files
- `convert_jws_to_ylk_direct()` - Direct JWS to YLK conversion
- `load_ylk_file()` / `save_ylk_file()` - YLK file I/O operations
- `plot_spectra()` - Real-time embedded spectrum visualization
- `create_baseline_for_file()` - Launch baseline creation tab for specific file
- `calculate_correlation()` - Pearson correlation matrix computation

## New Workflow

1. **File Selection**: Use File → Open Folder or recent folders menu
2. **File Processing**: JWS files automatically converted to YLK format in `converted_ylk/` subfolder
3. **Spectrum Selection**: Use file lists or double-click to add/remove spectra
   - Selected files are automatically hidden from "All Files" list
   - Files reappear in "All Files" when removed from selection
4. **Plotting**: Spectra displayed immediately in embedded canvas
5. **Data Display Toggle**: Switch between raw data and baseline-corrected views using the toggle button
6. **Baseline Creation**: Right-click on selected files → "Create Baseline" opens dedicated tab
7. **Baseline Editing**: Interactive parameter adjustment with live preview
8. **Saving**: Baseline parameters and data saved back to YLK file
9. **Analysis**: Correlation analysis available for selected spectra

## YLK Format Benefits

- **Self-contained**: Raw data, baseline, and metadata in single file
- **Human-readable**: JSON format allows manual inspection/editing
- **Extensible**: Easy to add new metadata fields or analysis results
- **Version control friendly**: Text-based format works well with Git
- **Cross-platform**: No binary format dependencies