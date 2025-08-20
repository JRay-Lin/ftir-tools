"""
Baseline Correction Module for FTIR Spectroscopy

Contains functions for various baseline correction methods including
ALS (Asymmetric Least Squares), polynomial, linear, and rolling minimum.
"""

import numpy as np
from scipy import sparse
from scipy.sparse.linalg import spsolve
from scipy.signal import savgol_filter
from scipy.ndimage import minimum_filter1d


def baseline_als(y, lam=1e5, p=0.01, niter=15):
    """
    Asymmetric Least Squares baseline correction

    Parameters:
    y: signal data
    lam: smoothness parameter (larger = smoother, try 1e4 to 1e8)
    p: asymmetry parameter (0 < p < 1, smaller = more asymmetric)
    niter: number of iterations

    Returns:
    baseline: calculated baseline
    """
    y = np.array(y, dtype=float)
    L = len(y)
    
    # Input validation
    if L < 3:
        raise ValueError("Input data must have at least 3 points for ALS baseline correction")
    
    if not np.isfinite(y).all():
        raise ValueError("Input data contains non-finite values (NaN or infinity)")
    
    if lam <= 0:
        raise ValueError("Lambda parameter must be positive")
        
    if not (0 < p < 1):
        raise ValueError("Asymmetry parameter p must be between 0 and 1")

    # Create difference matrix
    D = sparse.diags([1, -2, 1], [0, -1, -2], shape=(L, L - 2))
    D = D.tocsc()  # Convert to CSC format for efficiency

    w = np.ones(L)
    z = np.copy(y)

    for i in range(niter):
        W = sparse.diags(w, 0, shape=(L, L), format="csc")
        Z = W + lam * D.dot(D.T)
        z = spsolve(Z, w * y)
        w = p * (y > z) + (1 - p) * (y < z)

    return z


def baseline_correction(
    wavenumber, absorbance, method="als", lam=1e5, p=0.01, smooth=True
):
    """
    Apply baseline correction to FTIR spectrum (improved for horizontal baselines)
    Processing order: Raw data → Smooth → Baseline correct

    Parameters:
    wavenumber: wavenumber array
    absorbance: absorbance data array
    method: baseline correction method ('als', 'linear', 'polynomial', 'rolling_min')
    lam: lambda parameter for ALS
    p: asymmetry parameter for ALS
    smooth: whether to apply smoothing before baseline correction

    Returns:
    corrected: baseline-corrected spectrum
    """
    absorbance = np.array(absorbance, dtype=float)

    # Step 1: Apply smoothing if requested (before baseline correction)
    if smooth:
        # Use appropriate window length for FTIR data
        window_length = min(51, len(absorbance) // 10 if len(absorbance) > 50 else 5)
        if window_length % 2 == 0:  # Must be odd
            window_length += 1
        if window_length >= 5:
            absorbance = savgol_filter(
                absorbance, window_length=window_length, polyorder=3
            )

    # Step 2: Apply baseline correction
    if method == "als":
        baseline = baseline_als(absorbance, lam=lam, p=p, niter=50)
        corrected = absorbance - baseline
    elif method == "linear":
        # Simple linear baseline from endpoints
        x = np.array([0, len(absorbance) - 1])
        y = np.array([absorbance[0], absorbance[-1]])
        baseline = np.interp(range(len(absorbance)), x, y)
        corrected = absorbance - baseline
    elif method == "polynomial":
        # Polynomial baseline (degree 2)
        x = np.arange(len(absorbance))
        coeffs = np.polyfit(x, absorbance, 2)
        baseline = np.polyval(coeffs, x)
        corrected = absorbance - baseline
    elif method == "rolling_min":
        # Rolling minimum baseline (simpler approach)
        window_size = max(10, len(absorbance) // 20)
        baseline = minimum_filter1d(absorbance, size=window_size, mode="constant")
        corrected = absorbance - baseline
    else:
        corrected = absorbance

    return corrected


def get_baseline_with_raw(wavenumber, absorbance, method="als", lam=1e5, p=0.01, smooth=True):
    """
    Calculate baseline and return both raw data and baseline for visualization
    
    Parameters:
    wavenumber: wavenumber array
    absorbance: absorbance data array
    method: baseline correction method ('als', 'linear', 'polynomial', 'rolling_min')
    lam: lambda parameter for ALS
    p: asymmetry parameter for ALS
    smooth: whether to apply smoothing before baseline correction
    
    Returns:
    tuple: (processed_absorbance, baseline, corrected_spectrum)
    """
    absorbance = np.array(absorbance, dtype=float)
    original_abs = absorbance.copy()
    
    # Step 1: Apply smoothing if requested (before baseline correction)
    if smooth:
        window_length = min(51, len(absorbance) // 10 if len(absorbance) > 50 else 5)
        if window_length % 2 == 0:
            window_length += 1
        if window_length >= 5:
            absorbance = savgol_filter(
                absorbance, window_length=window_length, polyorder=3
            )
    
    # Step 2: Calculate baseline
    if method == "als":
        baseline = baseline_als(absorbance, lam=lam, p=p, niter=50)
        corrected = absorbance - baseline
    elif method == "linear":
        x = np.array([0, len(absorbance) - 1])
        y = np.array([absorbance[0], absorbance[-1]])
        baseline = np.interp(range(len(absorbance)), x, y)
        corrected = absorbance - baseline
    elif method == "polynomial":
        x = np.arange(len(absorbance))
        coeffs = np.polyfit(x, absorbance, 2)
        baseline = np.polyval(coeffs, x)
        corrected = absorbance - baseline
    elif method == "rolling_min":
        window_size = max(10, len(absorbance) // 20)
        baseline = minimum_filter1d(absorbance, size=window_size, mode="constant")
        corrected = absorbance - baseline
    else:
        baseline = np.zeros_like(absorbance)
        corrected = absorbance
    
    return absorbance, baseline, corrected


def get_baseline_methods():
    """
    Get available baseline correction methods with descriptions

    Returns:
    list: tuples of (description, method_key)
    """
    return [
        ("ALS (Asymmetric Least Squares)", "als"),
    ]
