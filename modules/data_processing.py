"""
Data Processing Module for FTIR Spectroscopy

Contains functions for data preprocessing, normalization, and analysis.
"""

import numpy as np
import pandas as pd
from scipy.stats import pearsonr


def preprocess_data(df, normalize=False):
    """
    Preprocess FTIR data

    Parameters:
    df: pandas DataFrame with 'wavenumber' and 'absorbance' columns
    normalize: whether to apply normalization

    Returns:
    pandas DataFrame: processed data
    """
    absorbance = df["absorbance"].values

    if normalize:
        # Optional normalization (after baseline correction)
        norm_abs = (absorbance - np.min(absorbance)) / (
            np.max(absorbance) - np.min(absorbance)
        )
        return pd.DataFrame({"wavenumber": df["wavenumber"], "absorbance": norm_abs})
    else:
        # Return raw data for baseline correction
        return df.copy()


def calculate_correlation_matrix(data_list):
    """
    Calculate Pearson correlation matrix for multiple spectra

    Parameters:
    data_list: list of pandas DataFrames containing spectral data

    Returns:
    numpy array: correlation matrix
    """
    pre_data = [df["absorbance"].values for df in data_list]
    num = len(pre_data)
    corr_matrix = np.zeros((num, num))

    for i in range(num):
        for j in range(num):
            corr_coeff, _ = pearsonr(pre_data[i], pre_data[j])
            corr_matrix[i, j] = float(corr_coeff)

    return corr_matrix


def validate_spectral_data(df):
    """
    Validate spectral data format and content

    Parameters:
    df: pandas DataFrame to validate

    Returns:
    tuple: (is_valid, error_message)
    """
    # Check required columns
    if "wavenumber" not in df.columns or "absorbance" not in df.columns:
        return False, "Missing required columns: 'wavenumber' and/or 'absorbance'"

    # Check for empty data
    if len(df) == 0:
        return False, "Empty dataset"

    # Check for NaN values
    if df["wavenumber"].isna().any() or df["absorbance"].isna().any():
        return False, "Contains NaN values"

    # Check data types
    try:
        pd.to_numeric(df["wavenumber"], errors="raise")
        pd.to_numeric(df["absorbance"], errors="raise")
    except (ValueError, TypeError):
        return False, "Non-numeric data found"

    # Check for minimum data points
    if len(df) < 10:
        return False, "Insufficient data points (minimum 10 required)"

    return True, "Valid"


def get_spectral_info(df):
    """
    Get basic information about spectral data

    Parameters:
    df: pandas DataFrame containing spectral data

    Returns:
    dict: information about the spectrum
    """
    return {
        "num_points": len(df),
        "wavenumber_range": (df["wavenumber"].min(), df["wavenumber"].max()),
        "absorbance_range": (df["absorbance"].min(), df["absorbance"].max()),
        "wavenumber_resolution": np.mean(np.diff(df["wavenumber"].sort_values())),
    }


def interpolate_to_common_grid(data_list, method="linear"):
    """
    Interpolate multiple spectra to a common wavenumber grid

    Parameters:
    data_list: list of pandas DataFrames
    method: interpolation method ('linear', 'cubic', etc.)

    Returns:
    list: interpolated DataFrames on common grid
    """
    from scipy import interpolate

    # Find common wavenumber range
    min_wn = max(df["wavenumber"].min() for df in data_list)
    max_wn = min(df["wavenumber"].max() for df in data_list)

    # Create common grid (use highest resolution)
    resolutions = [
        np.mean(np.diff(np.sort(df["wavenumber"].values))) for df in data_list
    ]
    common_resolution = min(resolutions)

    common_wn = np.arange(min_wn, max_wn + common_resolution, common_resolution)

    # Interpolate each spectrum
    interpolated_data = []
    for df in data_list:
        sorted_idx = df["wavenumber"].argsort()
        wn_sorted = df["wavenumber"].iloc[sorted_idx]
        abs_sorted = df["absorbance"].iloc[sorted_idx]

        f = interpolate.interp1d(
            wn_sorted, abs_sorted, kind=method, bounds_error=False, fill_value=np.nan
        )
        interpolated_abs = f(common_wn)

        # Remove NaN values
        valid_mask = ~np.isnan(interpolated_abs)
        interpolated_df = pd.DataFrame(
            {
                "wavenumber": common_wn[valid_mask],
                "absorbance": interpolated_abs[valid_mask],
            }
        )
        interpolated_data.append(interpolated_df)

    return interpolated_data
