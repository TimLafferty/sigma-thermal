"""
Heating value calculations for fuels.

This module provides functions to calculate higher heating value (HHV) and
lower heating value (LHV) for various fuels including natural gas mixtures
and liquid fuels.

References
----------
.. [1] GPSA Engineering Data Book, 13th Edition
.. [2] Perry's Chemical Engineers' Handbook, 8th Edition
.. [3] GTS Energy Inc. Engineering-Functions.xlam, CombustionFunctions.bas
"""

from typing import Optional, Dict
from dataclasses import dataclass


# Heating value data for gaseous fuel components (BTU/lb, mass basis)
GAS_COMPONENT_HHV = {
    'Air': 0,
    'Argon': 0,
    'Methane': 23875,
    'Ethane': 22323,
    'Propane': 21669,
    'Butane': 21321,
    'Pentane': 21095,
    'Hexane': 20966,
    'CO2': 0,
    'CO': 4347,
    'C': 14093,  # Carbon
    'N2': 0,
    'H2': 61095,
    'O2': 0,
    'H2S': 7097,
    'H2O': 0,
}

GAS_COMPONENT_LHV = {
    'Air': 0,
    'Argon': 0,
    'Methane': 21495,
    'Ethane': 20418,
    'Propane': 19937,
    'Butane': 19678,
    'Pentane': 20485,
    'Hexane': 19403,
    'CO2': 0,
    'CO': 4347,
    'C': 14093,
    'N2': 0,
    'H2': 51623,
    'O2': 0,
    'H2S': 6537,
    'H2O': 0,
}

# Heating values for liquid fuels (BTU/lb)
LIQUID_FUEL_HHV = {
    'methanol': 9797,
    'gasoline': 20190,
    '#1 oil': 19423,
    '#2 oil': 18993,
    '#4 oil': 18844,
    '#5 oil': 18909,
    '#6 oil': 18126,
}

LIQUID_FUEL_LHV = {
    'methanol': 8706,
    'gasoline': 18790,
    '#1 oil': 18211,
    '#2 oil': 17855,
    '#4 oil': 17790,
    '#5 oil': 17929,
    '#6 oil': 17277,
}


@dataclass
class GasComposition:
    """
    Fuel gas composition in mass percentages.

    All values should be mass percentages (0-100).
    The sum should equal 100%.

    Attributes
    ----------
    air_mass : float, optional
        Air mass percentage (default: 0)
    argon_mass : float, optional
        Argon mass percentage (default: 0)
    methane_mass : float, optional
        Methane (CH4) mass percentage (default: 0)
    ethane_mass : float, optional
        Ethane (C2H6) mass percentage (default: 0)
    propane_mass : float, optional
        Propane (C3H8) mass percentage (default: 0)
    butane_mass : float, optional
        Butane (C4H10) mass percentage (default: 0)
    pentane_mass : float, optional
        Pentane (C5H12) mass percentage (default: 0)
    hexane_mass : float, optional
        Hexane (C6H14) mass percentage (default: 0)
    co2_mass : float, optional
        Carbon dioxide mass percentage (default: 0)
    co_mass : float, optional
        Carbon monoxide mass percentage (default: 0)
    c_mass : float, optional
        Carbon mass percentage (default: 0)
    n2_mass : float, optional
        Nitrogen mass percentage (default: 0)
    h2_mass : float, optional
        Hydrogen mass percentage (default: 0)
    o2_mass : float, optional
        Oxygen mass percentage (default: 0)
    h2s_mass : float, optional
        Hydrogen sulfide mass percentage (default: 0)
    h2o_mass : float, optional
        Water vapor mass percentage (default: 0)
    """
    air_mass: float = 0.0
    argon_mass: float = 0.0
    methane_mass: float = 0.0
    ethane_mass: float = 0.0
    propane_mass: float = 0.0
    butane_mass: float = 0.0
    pentane_mass: float = 0.0
    hexane_mass: float = 0.0
    co2_mass: float = 0.0
    co_mass: float = 0.0
    c_mass: float = 0.0
    n2_mass: float = 0.0
    h2_mass: float = 0.0
    o2_mass: float = 0.0
    h2s_mass: float = 0.0
    h2o_mass: float = 0.0

    def to_dict(self) -> Dict[str, float]:
        """Convert composition to dictionary for calculations"""
        return {
            'Air': self.air_mass,
            'Argon': self.argon_mass,
            'Methane': self.methane_mass,
            'Ethane': self.ethane_mass,
            'Propane': self.propane_mass,
            'Butane': self.butane_mass,
            'Pentane': self.pentane_mass,
            'Hexane': self.hexane_mass,
            'CO2': self.co2_mass,
            'CO': self.co_mass,
            'C': self.c_mass,
            'N2': self.n2_mass,
            'H2': self.h2_mass,
            'O2': self.o2_mass,
            'H2S': self.h2s_mass,
            'H2O': self.h2o_mass,
        }


def hhv_mass_gas(composition: GasComposition) -> float:
    """
    Calculate higher heating value (HHV) for a gas mixture on mass basis.

    Parameters
    ----------
    composition : GasComposition
        Gas composition with mass percentages (should sum to 100%)

    Returns
    -------
    float
        Higher heating value in BTU/lb

    Examples
    --------
    >>> # Pure methane
    >>> comp = GasComposition(methane_mass=100.0)
    >>> hhv_mass_gas(comp)
    23875.0

    >>> # Natural gas mixture (90% CH4, 5% C2H6, 3% C3H8, 2% N2)
    >>> comp = GasComposition(
    ...     methane_mass=90.0,
    ...     ethane_mass=5.0,
    ...     propane_mass=3.0,
    ...     n2_mass=2.0
    ... )
    >>> hhv_mass_gas(comp)
    22841.85

    Notes
    -----
    Replicates VBA function HHVMass for FuelType="Gas"
    Mass-weighted sum: HHV = Σ(HHV_i * mass_i / 100)
    """
    comp_dict = composition.to_dict()
    hhv = 0.0

    for component, mass_percent in comp_dict.items():
        component_hhv = GAS_COMPONENT_HHV.get(component, 0.0)
        hhv += component_hhv * mass_percent / 100.0

    return hhv


def lhv_mass_gas(composition: GasComposition) -> float:
    """
    Calculate lower heating value (LHV) for a gas mixture on mass basis.

    The LHV excludes the latent heat of vaporization of water in the
    combustion products.

    Parameters
    ----------
    composition : GasComposition
        Gas composition with mass percentages (should sum to 100%)

    Returns
    -------
    float
        Lower heating value in BTU/lb

    Examples
    --------
    >>> # Pure methane
    >>> comp = GasComposition(methane_mass=100.0)
    >>> lhv_mass_gas(comp)
    21495.0

    >>> # Pure hydrogen
    >>> comp = GasComposition(h2_mass=100.0)
    >>> lhv_mass_gas(comp)
    51623.0

    Notes
    -----
    Replicates VBA function LHVMass for FuelType="Gas"
    Mass-weighted sum: LHV = Σ(LHV_i * mass_i / 100)

    The difference between HHV and LHV accounts for the latent heat
    of water vapor formation from hydrogen in the fuel.
    """
    comp_dict = composition.to_dict()
    lhv = 0.0

    for component, mass_percent in comp_dict.items():
        component_lhv = GAS_COMPONENT_LHV.get(component, 0.0)
        lhv += component_lhv * mass_percent / 100.0

    return lhv


def hhv_mass_liquid(fuel_type: str) -> float:
    """
    Get higher heating value (HHV) for liquid fuels.

    Parameters
    ----------
    fuel_type : str
        Liquid fuel type. Valid options:
        - 'methanol'
        - 'gasoline'
        - '#1 oil'
        - '#2 oil'
        - '#4 oil'
        - '#5 oil'
        - '#6 oil'

    Returns
    -------
    float
        Higher heating value in BTU/lb

    Raises
    ------
    ValueError
        If fuel_type is not recognized

    Examples
    --------
    >>> hhv_mass_liquid('#2 oil')
    18993

    >>> hhv_mass_liquid('gasoline')
    20190

    Notes
    -----
    Replicates VBA function HHVMass for liquid fuel types.
    Values from GPSA Engineering Data Book.
    """
    fuel_lower = fuel_type.lower()
    if fuel_lower not in LIQUID_FUEL_HHV:
        valid_fuels = ', '.join(LIQUID_FUEL_HHV.keys())
        raise ValueError(
            f"Unknown liquid fuel type: '{fuel_type}'. "
            f"Valid options: {valid_fuels}"
        )

    return LIQUID_FUEL_HHV[fuel_lower]


def lhv_mass_liquid(fuel_type: str) -> float:
    """
    Get lower heating value (LHV) for liquid fuels.

    Parameters
    ----------
    fuel_type : str
        Liquid fuel type. Valid options:
        - 'methanol'
        - 'gasoline'
        - '#1 oil'
        - '#2 oil'
        - '#4 oil'
        - '#5 oil'
        - '#6 oil'

    Returns
    -------
    float
        Lower heating value in BTU/lb

    Raises
    ------
    ValueError
        If fuel_type is not recognized

    Examples
    --------
    >>> lhv_mass_liquid('#2 oil')
    17855

    >>> lhv_mass_liquid('methanol')
    8706

    Notes
    -----
    Replicates VBA function LHVMass for liquid fuel types.
    """
    fuel_lower = fuel_type.lower()
    if fuel_lower not in LIQUID_FUEL_LHV:
        valid_fuels = ', '.join(LIQUID_FUEL_LHV.keys())
        raise ValueError(
            f"Unknown liquid fuel type: '{fuel_type}'. "
            f"Valid options: {valid_fuels}"
        )

    return LIQUID_FUEL_LHV[fuel_lower]


# VBA compatibility wrappers with original function signature
def HHVMass(
    fuel_type: str,
    air_mass: float = 0,
    argon_mass: float = 0,
    methane_mass: float = 0,
    ethane_mass: float = 0,
    propane_mass: float = 0,
    butane_mass: float = 0,
    pentane_mass: float = 0,
    hexane_mass: float = 0,
    co2_mass: float = 0,
    co_mass: float = 0,
    c_mass: float = 0,
    n2_mass: float = 0,
    h2_mass: float = 0,
    o2_mass: float = 0,
    h2s_mass: float = 0,
    h2o_mass: float = 0,
) -> float:
    """
    VBA-compatible wrapper for HHV calculation.

    Parameters
    ----------
    fuel_type : str
        "Gas" for gas mixture, or liquid fuel name
    air_mass through h2o_mass : float
        Component mass percentages (0-100)

    Returns
    -------
    float
        Higher heating value in BTU/lb

    Notes
    -----
    This function replicates the exact VBA signature of HHVMass.
    For cleaner code, prefer using hhv_mass_gas() with GasComposition
    or hhv_mass_liquid() directly.
    """
    if fuel_type.lower() == "gas":
        composition = GasComposition(
            air_mass=air_mass,
            argon_mass=argon_mass,
            methane_mass=methane_mass,
            ethane_mass=ethane_mass,
            propane_mass=propane_mass,
            butane_mass=butane_mass,
            pentane_mass=pentane_mass,
            hexane_mass=hexane_mass,
            co2_mass=co2_mass,
            co_mass=co_mass,
            c_mass=c_mass,
            n2_mass=n2_mass,
            h2_mass=h2_mass,
            o2_mass=o2_mass,
            h2s_mass=h2s_mass,
            h2o_mass=h2o_mass,
        )
        return hhv_mass_gas(composition)
    else:
        return hhv_mass_liquid(fuel_type)


def LHVMass(
    fuel_type: str,
    air_mass: float = 0,
    argon_mass: float = 0,
    methane_mass: float = 0,
    ethane_mass: float = 0,
    propane_mass: float = 0,
    butane_mass: float = 0,
    pentane_mass: float = 0,
    hexane_mass: float = 0,
    co2_mass: float = 0,
    co_mass: float = 0,
    c_mass: float = 0,
    n2_mass: float = 0,
    h2_mass: float = 0,
    o2_mass: float = 0,
    h2s_mass: float = 0,
    h2o_mass: float = 0,
) -> float:
    """
    VBA-compatible wrapper for LHV calculation.

    Parameters
    ----------
    fuel_type : str
        "Gas" for gas mixture, or liquid fuel name
    air_mass through h2o_mass : float
        Component mass percentages (0-100)

    Returns
    -------
    float
        Lower heating value in BTU/lb

    Notes
    -----
    This function replicates the exact VBA signature of LHVMass.
    For cleaner code, prefer using lhv_mass_gas() with GasComposition
    or lhv_mass_liquid() directly.
    """
    if fuel_type.lower() == "gas":
        composition = GasComposition(
            air_mass=air_mass,
            argon_mass=argon_mass,
            methane_mass=methane_mass,
            ethane_mass=ethane_mass,
            propane_mass=propane_mass,
            butane_mass=butane_mass,
            pentane_mass=pentane_mass,
            hexane_mass=hexane_mass,
            co2_mass=co2_mass,
            co_mass=co_mass,
            c_mass=c_mass,
            n2_mass=n2_mass,
            h2_mass=h2_mass,
            o2_mass=o2_mass,
            h2s_mass=h2s_mass,
            h2o_mass=h2o_mass,
        )
        return lhv_mass_gas(composition)
    else:
        return lhv_mass_liquid(fuel_type)
