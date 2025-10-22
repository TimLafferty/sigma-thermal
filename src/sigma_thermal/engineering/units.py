"""
Unit conversion utilities using pint.

This module provides a wrapper around pint for consistent unit handling
throughout the sigma_thermal library.
"""

import pint
from typing import Union, Optional
import numpy as np


# Create a unit registry for the application
ureg = pint.UnitRegistry()

# Enable wrapping of numpy arrays
ureg.setup_matplotlib()

# Define common aliases
ureg.define('scfh = ft**3/hour = standard_cubic_feet_per_hour')
ureg.define('scfm = ft**3/minute = standard_cubic_feet_per_minute')
ureg.define('mmBtu = 1e6 * Btu = million_Btu')

# Quantity type alias
Quantity = ureg.Quantity


def convert(
    value: Union[float, np.ndarray],
    from_units: str,
    to_units: str
) -> Union[float, np.ndarray]:
    """
    Convert a value from one unit to another.

    Parameters
    ----------
    value : float or ndarray
        Value to convert
    from_units : str
        Source units
    to_units : str
        Target units

    Returns
    -------
    float or ndarray
        Converted value

    Examples
    --------
    >>> convert(1000, 'Btu/hr', 'kW')
    0.293...

    >>> convert(100, 'degF', 'degC')
    37.7...

    >>> convert(1, 'ft', 'meter')
    0.3048
    """
    # Use Quantity constructor to handle offset units (degF, degC)
    quantity = ureg.Quantity(value, from_units)
    converted = quantity.to(to_units)
    return converted.magnitude


def Q_(value: Union[float, np.ndarray], units: str) -> Quantity:
    """
    Create a Quantity with units.

    This is a shorthand for creating pint Quantity objects.
    Handles offset temperature units (degF, degC) properly.

    Parameters
    ----------
    value : float or ndarray
        Numerical value
    units : str
        Units string

    Returns
    -------
    Quantity
        Pint quantity object

    Examples
    --------
    >>> duty = Q_(1000000, 'Btu/hr')
    >>> duty.to('kW')
    <Quantity(293.07107, 'kilowatt')>

    >>> temp = Q_(500, 'degF')
    >>> temp.to('degC')
    <Quantity(260.0, 'degree_Celsius')>
    """
    # For offset units (degF, degC), use Quantity constructor directly
    # to avoid multiplication which pint doesn't allow
    return ureg.Quantity(value, units)


def strip_units(quantity: Quantity) -> Union[float, np.ndarray]:
    """
    Extract magnitude from a Quantity, discarding units.

    Parameters
    ----------
    quantity : Quantity
        Pint quantity object

    Returns
    -------
    float or ndarray
        Magnitude without units
    """
    return quantity.magnitude


def ensure_units(
    value: Union[float, Quantity],
    expected_units: str
) -> Quantity:
    """
    Ensure a value has the specified units.

    If value is a float, attach units.
    If value is a Quantity, convert to expected units.

    Parameters
    ----------
    value : float or Quantity
        Value to process
    expected_units : str
        Expected units

    Returns
    -------
    Quantity
        Value with expected units

    Examples
    --------
    >>> ensure_units(100, 'degF')
    <Quantity(100, 'degree_Fahrenheit')>

    >>> temp_c = Q_(37.8, 'degC')
    >>> ensure_units(temp_c, 'degF')
    <Quantity(100.04, 'degree_Fahrenheit')>
    """
    if isinstance(value, Quantity):
        return value.to(expected_units)
    else:
        return Q_(value, expected_units)


# Common unit conversion functions for thermal engineering

def btu_hr_to_kw(btu_hr: float) -> float:
    """Convert BTU/hr to kW"""
    return convert(btu_hr, 'Btu/hr', 'kW')


def kw_to_btu_hr(kw: float) -> float:
    """Convert kW to BTU/hr"""
    return convert(kw, 'kW', 'Btu/hr')


def degf_to_degc(degf: float) -> float:
    """Convert Fahrenheit to Celsius"""
    return (degf - 32) * 5 / 9


def degc_to_degf(degc: float) -> float:
    """Convert Celsius to Fahrenheit"""
    return degc * 9 / 5 + 32


def psi_to_pa(psi: float) -> float:
    """Convert PSI to Pascals"""
    return convert(psi, 'psi', 'Pa')


def pa_to_psi(pa: float) -> float:
    """Convert Pascals to PSI"""
    return convert(pa, 'Pa', 'psi')


def scfh_to_kg_hr(
    scfh: float,
    molecular_weight: float,
    temperature: float = 60,
    pressure: float = 14.7
) -> float:
    """
    Convert standard cubic feet per hour to kg/hr.

    Parameters
    ----------
    scfh : float
        Volumetric flow rate in SCFH
    molecular_weight : float
        Molecular weight in lb/lbmol
    temperature : float, optional
        Standard temperature in degF (default 60)
    pressure : float, optional
        Standard pressure in psia (default 14.7)

    Returns
    -------
    float
        Mass flow rate in kg/hr

    Notes
    -----
    Uses ideal gas law: PV = nRT
    Standard conditions: 60°F, 14.7 psia
    """
    # Convert to SI
    temp_k = (temperature + 459.67) * 5 / 9  # Rankine to Kelvin
    pressure_pa = pressure * 6894.76  # psia to Pa

    # Ideal gas constant
    R = 8.314  # J/(mol·K)

    # Molar volume at standard conditions (m³/mol)
    molar_volume = R * temp_k / pressure_pa

    # Convert SCFH to m³/hr
    volume_m3_hr = scfh * 0.0283168

    # Calculate molar flow (mol/hr)
    molar_flow = volume_m3_hr / molar_volume

    # Convert to mass flow (kg/hr)
    mass_flow_kg_hr = molar_flow * molecular_weight / 2.20462  # lb to kg

    return mass_flow_kg_hr


# Export commonly used items
__all__ = [
    'ureg',
    'Quantity',
    'Q_',
    'convert',
    'strip_units',
    'ensure_units',
    'btu_hr_to_kw',
    'kw_to_btu_hr',
    'degf_to_degc',
    'degc_to_degf',
    'psi_to_pa',
    'pa_to_psi',
    'scfh_to_kg_hr',
]
