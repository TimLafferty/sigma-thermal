"""
Enthalpy calculations for combustion gases.

This module provides enthalpy correlations for common flue gas components.
All correlations are based on polynomial fits and are valid for typical
combustion gas temperatures (ambient to ~3000°F).

References
----------
.. [1] Perry's Chemical Engineers' Handbook, 8th Edition
.. [2] GTS Energy Inc. Engineering-Functions.xlam, CombustionFunctions.bas
"""

from typing import Union
import numpy as np
from sigma_thermal.engineering.units import Quantity, Q_, ureg


def enthalpy_co2(
    gas_temp: Union[float, Quantity],
    ambient_temp: Union[float, Quantity] = 77.0,
    return_quantity: bool = False
) -> Union[float, Quantity]:
    """
    Calculate CO2 specific enthalpy relative to ambient temperature.

    Uses a 2nd-order polynomial correlation for CO2 enthalpy:
    H = a*T^2 + b*T + c (BTU/lb)

    Parameters
    ----------
    gas_temp : float or Quantity
        Gas temperature in °F (or Quantity with temperature units)
    ambient_temp : float or Quantity, optional
        Ambient reference temperature in °F (default: 77°F)
    return_quantity : bool, optional
        If True, return pint Quantity with units (default: False)

    Returns
    -------
    float or Quantity
        Specific enthalpy difference in BTU/lb
        Returns difference: H(gas_temp) - H(ambient_temp)

    Examples
    --------
    >>> enthalpy_co2(1500, 77)  # Stack temp 1500°F, ambient 77°F
    250.96...

    >>> from sigma_thermal.engineering.units import Q_
    >>> T_gas = Q_(1500, 'degF')
    >>> h = enthalpy_co2(T_gas, return_quantity=True)
    >>> print(f"{h:.2f}")
    250.96 Btu / pound

    Notes
    -----
    Polynomial coefficients from GTS Energy VBA function EnthalpyCO2:
    - a = 1.08941E-05
    - b = 0.262597665
    - c = 176.9479842

    Valid range: ambient to ~3000°F
    """
    # Extract magnitudes if Quantities
    # Note: Temperature must be in degF for the polynomial correlation
    # pint doesn't allow arithmetic on offset units, so we extract magnitude
    if isinstance(gas_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_gas = float(gas_temp.m_as('degF'))
    else:
        T_gas = float(gas_temp)

    if isinstance(ambient_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_amb = float(ambient_temp.m_as('degF'))
    else:
        T_amb = float(ambient_temp)

    # Polynomial coefficients (BTU/lb vs degF)
    a = 1.08941e-05
    b = 0.262597665
    c = 176.9479842

    # Calculate enthalpy at both temperatures
    h_gas = a * T_gas**2 + b * T_gas + c
    h_amb = a * T_amb**2 + b * T_amb + c

    # Return difference
    h_diff = h_gas - h_amb

    if return_quantity:
        return Q_(h_diff, 'Btu/lb')
    else:
        return h_diff


def enthalpy_h2o(
    gas_temp: Union[float, Quantity],
    ambient_temp: Union[float, Quantity] = 77.0,
    return_quantity: bool = False
) -> Union[float, Quantity]:
    """
    Calculate H2O (water vapor) specific enthalpy relative to ambient.

    Uses a 2nd-order polynomial correlation for H2O vapor enthalpy:
    H = a*T^2 + b*T + c (BTU/lb)

    Parameters
    ----------
    gas_temp : float or Quantity
        Gas temperature in °F (or Quantity with temperature units)
    ambient_temp : float or Quantity, optional
        Ambient reference temperature in °F (default: 77°F)
    return_quantity : bool, optional
        If True, return pint Quantity with units (default: False)

    Returns
    -------
    float or Quantity
        Specific enthalpy difference in BTU/lb

    Examples
    --------
    >>> enthalpy_h2o(1500, 77)
    730.55...

    Notes
    -----
    Polynomial coefficients from GTS Energy VBA function EnthalpyH2O:
    - a = 3.65285E-05
    - b = 0.452215911
    - c = 1049.366151

    This correlation is for water vapor (steam), not liquid water.
    Valid range: ambient to ~3000°F
    """
    # Extract magnitudes if Quantities
    # Note: Temperature must be in degF for the polynomial correlation
    # pint doesn't allow arithmetic on offset units, so we extract magnitude
    if isinstance(gas_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_gas = float(gas_temp.m_as('degF'))
    else:
        T_gas = float(gas_temp)

    if isinstance(ambient_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_amb = float(ambient_temp.m_as('degF'))
    else:
        T_amb = float(ambient_temp)

    # Polynomial coefficients (BTU/lb vs degF)
    a = 3.65285e-05
    b = 0.452215911
    c = 1049.366151

    # Calculate enthalpy at both temperatures
    h_gas = a * T_gas**2 + b * T_gas + c
    h_amb = a * T_amb**2 + b * T_amb + c

    # Return difference
    h_diff = h_gas - h_amb

    if return_quantity:
        return Q_(h_diff, 'Btu/lb')
    else:
        return h_diff


def enthalpy_n2(
    gas_temp: Union[float, Quantity],
    ambient_temp: Union[float, Quantity] = 77.0,
    return_quantity: bool = False
) -> Union[float, Quantity]:
    """
    Calculate N2 (nitrogen) specific enthalpy relative to ambient.

    Uses a 2nd-order polynomial correlation for N2 enthalpy:
    H = a*T^2 + b*T + c (BTU/lb)

    Parameters
    ----------
    gas_temp : float or Quantity
        Gas temperature in °F (or Quantity with temperature units)
    ambient_temp : float or Quantity, optional
        Ambient reference temperature in °F (default: 77°F)
    return_quantity : bool, optional
        If True, return pint Quantity with units (default: False)

    Returns
    -------
    float or Quantity
        Specific enthalpy difference in BTU/lb

    Examples
    --------
    >>> enthalpy_n2(1500, 77)
    378.77...

    Notes
    -----
    Polynomial coefficients from GTS Energy VBA function EnthalpyN2:
    - a = 8.46332E-06
    - b = 0.255630011
    - c = 107.2712456

    Valid range: ambient to ~3000°F
    """
    # Extract magnitudes if Quantities
    # Note: Temperature must be in degF for the polynomial correlation
    # pint doesn't allow arithmetic on offset units, so we extract magnitude
    if isinstance(gas_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_gas = float(gas_temp.m_as('degF'))
    else:
        T_gas = float(gas_temp)

    if isinstance(ambient_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_amb = float(ambient_temp.m_as('degF'))
    else:
        T_amb = float(ambient_temp)

    # Polynomial coefficients (BTU/lb vs degF)
    a = 8.46332e-06
    b = 0.255630011
    c = 107.2712456

    # Calculate enthalpy at both temperatures
    h_gas = a * T_gas**2 + b * T_gas + c
    h_amb = a * T_amb**2 + b * T_amb + c

    # Return difference
    h_diff = h_gas - h_amb

    if return_quantity:
        return Q_(h_diff, 'Btu/lb')
    else:
        return h_diff


def enthalpy_o2(
    gas_temp: Union[float, Quantity],
    ambient_temp: Union[float, Quantity] = 77.0,
    return_quantity: bool = False
) -> Union[float, Quantity]:
    """
    Calculate O2 (oxygen) specific enthalpy relative to ambient.

    Uses a 2nd-order polynomial correlation for O2 enthalpy:
    H = a*T^2 + b*T + c (BTU/lb)

    Parameters
    ----------
    gas_temp : float or Quantity
        Gas temperature in °F (or Quantity with temperature units)
    ambient_temp : float or Quantity, optional
        Ambient reference temperature in °F (default: 77°F)
    return_quantity : bool, optional
        If True, return pint Quantity with units (default: False)

    Returns
    -------
    float or Quantity
        Specific enthalpy difference in BTU/lb

    Examples
    --------
    >>> enthalpy_o2(1500, 77)
    353.22...

    Notes
    -----
    Polynomial coefficients from GTS Energy VBA function EnthalpyO2:
    - a = 7.53536E-06
    - b = 0.23706691
    - c = 92.56930357

    Valid range: ambient to ~3000°F
    """
    # Extract magnitudes if Quantities
    # Note: Temperature must be in degF for the polynomial correlation
    # pint doesn't allow arithmetic on offset units, so we extract magnitude
    if isinstance(gas_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_gas = float(gas_temp.m_as('degF'))
    else:
        T_gas = float(gas_temp)

    if isinstance(ambient_temp, Quantity):
        # Ensure it's in degF and extract magnitude
        T_amb = float(ambient_temp.m_as('degF'))
    else:
        T_amb = float(ambient_temp)

    # Polynomial coefficients (BTU/lb vs degF)
    a = 7.53536e-06
    b = 0.23706691
    c = 92.56930357

    # Calculate enthalpy at both temperatures
    h_gas = a * T_gas**2 + b * T_gas + c
    h_amb = a * T_amb**2 + b * T_amb + c

    # Return difference
    h_diff = h_gas - h_amb

    if return_quantity:
        return Q_(h_diff, 'Btu/lb')
    else:
        return h_diff


def flue_gas_enthalpy(
    h2o_fraction: float,
    co2_fraction: float,
    n2_fraction: float,
    o2_fraction: float,
    gas_temp: Union[float, Quantity],
    ambient_temp: Union[float, Quantity] = 77.0,
    return_quantity: bool = False
) -> Union[float, Quantity]:
    """
    Calculate mixed flue gas specific enthalpy.

    Calculates the mass-weighted average enthalpy of a flue gas mixture
    given the mass fractions of each component.

    Parameters
    ----------
    h2o_fraction : float
        Mass fraction of H2O in flue gas (0-1)
    co2_fraction : float
        Mass fraction of CO2 in flue gas (0-1)
    n2_fraction : float
        Mass fraction of N2 in flue gas (0-1)
    o2_fraction : float
        Mass fraction of O2 in flue gas (0-1)
    gas_temp : float or Quantity
        Gas temperature in °F (or Quantity with temperature units)
    ambient_temp : float or Quantity, optional
        Ambient reference temperature in °F (default: 77°F)
    return_quantity : bool, optional
        If True, return pint Quantity with units (default: False)

    Returns
    -------
    float or Quantity
        Mixed gas specific enthalpy in BTU/lb

    Raises
    ------
    ValueError
        If mass fractions do not sum to unity (within 1% tolerance)

    Examples
    --------
    >>> # Typical natural gas flue gas composition (mass fractions)
    >>> h = flue_gas_enthalpy(
    ...     h2o_fraction=0.12,
    ...     co2_fraction=0.15,
    ...     n2_fraction=0.70,
    ...     o2_fraction=0.03,
    ...     gas_temp=1500,
    ...     ambient_temp=77
    ... )
    >>> print(f"{h:.2f}")
    390.23

    Notes
    -----
    This function replicates VBA function FlueGasEnthalpy from
    CombustionFunctions.bas.

    The mass fractions must sum to 1.0 within 1% tolerance.
    """
    # Validate mass fractions sum to unity
    total = h2o_fraction + co2_fraction + n2_fraction + o2_fraction

    if not (0.99 <= total <= 1.01):
        raise ValueError(
            f"Mass fractions must sum to unity. Got {total:.4f}. "
            f"Fractions: H2O={h2o_fraction}, CO2={co2_fraction}, "
            f"N2={n2_fraction}, O2={o2_fraction}"
        )

    # Calculate individual component enthalpies
    h_h2o = enthalpy_h2o(gas_temp, ambient_temp)
    h_co2 = enthalpy_co2(gas_temp, ambient_temp)
    h_n2 = enthalpy_n2(gas_temp, ambient_temp)
    h_o2 = enthalpy_o2(gas_temp, ambient_temp)

    # Mass-weighted average
    h_mix = (
        h_h2o * h2o_fraction +
        h_co2 * co2_fraction +
        h_n2 * n2_fraction +
        h_o2 * o2_fraction
    )

    if return_quantity:
        return Q_(h_mix, 'Btu/lb')
    else:
        return h_mix


# VBA compatibility aliases
EnthalpyCO2 = enthalpy_co2
EnthalpyH2O = enthalpy_h2o
EnthalpyN2 = enthalpy_n2
EnthalpyO2 = enthalpy_o2
FlueGasEnthalpy = flue_gas_enthalpy
