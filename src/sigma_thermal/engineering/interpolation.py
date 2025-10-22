"""
Interpolation utilities for engineering calculations.

This module provides interpolation functions that replicate
the VBA Interpolate() function from EngineeringFunctions.bas.
"""

from typing import Union, List
import numpy as np


def linear_interpolate(
    x1: float,
    x: float,
    x2: float,
    y1: float,
    y2: float
) -> float:
    """
    Perform linear interpolation.

    This replicates the VBA function:
    Function Interpolate(x1 As Single, x As Single, x2 As Single,
                         y1 As Single, y2 As Single)

    Parameters
    ----------
    x1 : float
        Lower bound x value
    x : float
        Interpolation point
    x2 : float
        Upper bound x value
    y1 : float
        Function value at x1
    y2 : float
        Function value at x2

    Returns
    -------
    float
        Interpolated value at x

    Examples
    --------
    >>> linear_interpolate(0, 5, 10, 0, 100)
    50.0

    >>> linear_interpolate(100, 150, 200, 20, 30)
    25.0

    Notes
    -----
    If x1 == x2 (division by zero), returns y1.
    Extrapolation is permitted (x outside [x1, x2]).
    """
    if abs(x2 - x1) < 1e-10:
        # Avoid division by zero
        return y1

    # Linear interpolation formula: y = y1 + (x - x1) * (y2 - y1) / (x2 - x1)
    slope = (y2 - y1) / (x2 - x1)
    y = y1 + (x - x1) * slope

    return y


def interpolate_from_table(
    x: float,
    x_values: Union[List[float], np.ndarray],
    y_values: Union[List[float], np.ndarray],
    extrapolate: bool = True
) -> float:
    """
    Interpolate from a table of x and y values.

    Parameters
    ----------
    x : float
        Value to interpolate at
    x_values : array-like
        Array of x values (must be sorted)
    y_values : array-like
        Array of corresponding y values
    extrapolate : bool, optional
        Whether to extrapolate outside the range (default True)
        If False, returns boundary values for out-of-range x

    Returns
    -------
    float
        Interpolated y value

    Examples
    --------
    >>> x_vals = [0, 10, 20, 30]
    >>> y_vals = [0, 100, 150, 180]
    >>> interpolate_from_table(15, x_vals, y_vals)
    125.0

    Raises
    ------
    ValueError
        If x_values and y_values have different lengths
        If x_values has fewer than 2 elements
    """
    x_values = np.asarray(x_values)
    y_values = np.asarray(y_values)

    if len(x_values) != len(y_values):
        raise ValueError("x_values and y_values must have the same length")

    if len(x_values) < 2:
        raise ValueError("Need at least 2 points for interpolation")

    # Use numpy's interp function
    # numpy.interp automatically handles sorted arrays and extrapolation
    if extrapolate:
        y = np.interp(x, x_values, y_values)
    else:
        # Clip to boundaries
        if x < x_values[0]:
            y = y_values[0]
        elif x > x_values[-1]:
            y = y_values[-1]
        else:
            y = np.interp(x, x_values, y_values)

    return float(y)


def bilinear_interpolate(
    x: float,
    y: float,
    x_values: Union[List[float], np.ndarray],
    y_values: Union[List[float], np.ndarray],
    z_grid: np.ndarray
) -> float:
    """
    Perform bilinear interpolation on a 2D grid.

    Useful for property tables with two independent variables
    (e.g., temperature and pressure).

    Parameters
    ----------
    x : float
        First coordinate value
    y : float
        Second coordinate value
    x_values : array-like
        Array of x coordinate values (rows)
    y_values : array-like
        Array of y coordinate values (columns)
    z_grid : ndarray
        2D array of z values with shape (len(x_values), len(y_values))

    Returns
    -------
    float
        Interpolated z value at (x, y)

    Examples
    --------
    >>> x_vals = [0, 10, 20]
    >>> y_vals = [0, 5, 10]
    >>> z_grid = np.array([[0, 5, 10],
    ...                    [10, 15, 20],
    ...                    [20, 25, 30]])
    >>> bilinear_interpolate(5, 2.5, x_vals, y_vals, z_grid)
    7.5
    """
    from scipy.interpolate import RegularGridInterpolator

    x_values = np.asarray(x_values)
    y_values = np.asarray(y_values)
    z_grid = np.asarray(z_grid)

    # Create interpolator
    interpolator = RegularGridInterpolator(
        (x_values, y_values),
        z_grid,
        method='linear',
        bounds_error=False,
        fill_value=None  # Extrapolate
    )

    # Interpolate at point
    z = interpolator([x, y])

    return float(z[0])


# Alias for compatibility with VBA function name
Interpolate = linear_interpolate
