"""
Products of combustion calculations.

This module provides functions to calculate the mass and volume of combustion
products (CO2, H2O, N2, O2, SO2) based on fuel composition and flow rates.

The calculations are based on stoichiometric combustion equations and use
coefficients derived from balanced chemical reactions.

References
----------
.. [1] GPSA Engineering Data Book, 13th Edition
.. [2] Perry's Chemical Engineers' Handbook, 8th Edition
.. [3] GTS Energy Inc. Engineering-Functions.xlam, CombustionFunctions.bas
"""

from typing import Optional
from dataclasses import dataclass


# Stoichiometric coefficients for gas combustion (mass basis)
# These represent lb product per lb fuel component

# Water production coefficients (lb H2O / lb fuel)
GAS_H2O_MASS_COEFF = {
    'Air': 0.0,
    'Argon': 0.0,
    'Methane': 2.246,
    'Ethane': 1.797,
    'Propane': 1.634,
    'Butane': 1.55,
    'Pentane': 1.5,
    'Hexane': 1.46,
    'CO2': 0.0,
    'CO': 0.0,
    'C': 0.0,
    'N2': 0.0,
    'H2': 8.937,
    'O2': 0.0,
    'H2S': 0.529,
    'H2O': 0.0,
}

# CO2 production coefficients (lb CO2 / lb fuel)
GAS_CO2_MASS_COEFF = {
    'Air': 0.0,
    'Argon': 0.0,
    'Methane': 2.743,
    'Ethane': 2.927,
    'Propane': 2.994,
    'Butane': 3.029,
    'Pentane': 3.05,
    'Hexane': 3.06,
    'CO2': 0.0,
    'CO': 1.571,
    'C': 3.664,
    'N2': 0.0,
    'H2': 0.0,
    'O2': 0.0,
    'H2S': 0.0,
}

# N2 production coefficients (lb N2 / lb fuel)
GAS_N2_MASS_COEFF = {
    'Air': 0.0,
    'Argon': 0.0,
    'Methane': 13.246,
    'Ethane': 12.367,
    'Propane': 12.047,
    'Butane': 11.882,
    'Pentane': 11.81,
    'Hexane': 11.74,
    'CO2': 0.0,
    'CO': 1.897,
    'C': 8.846,
    'N2': 0.0,
    'H2': 26.353,
    'O2': 0.0,
    'H2S': 4.682,
}

# Stoichiometric coefficients for gas combustion (volume basis)
# These represent volumes of product per volume of fuel component

# CO2 production coefficients (vol CO2 / vol fuel)
GAS_CO2_VOL_COEFF = {
    'Air': 0.0,
    'Ammonia': 0.0,
    'Argon': 0.0,
    'Methane': 1.0,
    'Ethane': 2.0,
    'Propane': 3.0,
    'Butane': 4.0,
    'IButene': 4.0,
    'Pentane': 5.0,
    'Hexane': 6.0,
    'CO2': 0.0,
    'CO': 1.0,
    'C': 1.0,
    'N2': 0.0,
    'H2': 0.0,
    'O2': 0.0,
    'H2S': 0.0,
    'S': 0.0,
    'SO2': 0.0,
    'H2O': 0.0,
}

# H2O production coefficients (vol H2O / vol fuel)
GAS_H2O_VOL_COEFF = {
    'Air': 0.0,
    'Ammonia': 1.5,
    'Argon': 0.0,
    'Methane': 2.0,
    'Ethane': 3.0,
    'Propane': 4.0,
    'Butane': 5.0,
    'IButene': 4.0,
    'Pentane': 6.0,
    'Hexane': 7.0,
    'CO2': 0.0,
    'CO': 0.0,
    'C': 0.0,
    'N2': 0.0,
    'H2': 1.0,
    'O2': 0.0,
    'H2S': 1.0,
    'S': 0.0,
    'SO2': 0.0,
    'H2O': 1.0,
}

# N2 production coefficients (vol N2 / vol fuel)
# Note: These include N2 from combustion air (78% of air)
GAS_N2_VOL_COEFF = {
    'Air': 0.0,
    'Ammonia': 3.32,
    'Argon': 0.0,
    'Methane': 7.53,
    'Ethane': 13.18,
    'Propane': 18.82,
    'Butane': 24.47,
    'IButene': 22.59,
    'Pentane': 30.11,
    'Hexane': 35.76,
    'CO2': 0.0,
    'CO': 1.88,
    'C': 3.76,
    'N2': 0.0,
    'H2': 1.88,
    'O2': 3.76,
    'H2S': 0.0,
    'S': 1.88,
    'SO2': 0.0,
    'H2O': 5.65,
}

# SO2 production coefficients (vol SO2 / vol fuel)
GAS_SO2_VOL_COEFF = {
    'Air': 0.0,
    'Ammonia': 0.0,
    'Argon': 0.0,
    'Methane': 0.0,
    'Ethane': 0.0,
    'Propane': 0.0,
    'Butane': 0.0,
    'IButene': 0.0,
    'Pentane': 0.0,
    'Hexane': 0.0,
    'CO2': 0.0,
    'CO': 0.0,
    'C': 0.0,
    'N2': 0.0,
    'H2': 0.0,
    'O2': 0.0,
    'H2S': 1.0,
    'S': 1.0,
    'SO2': 1.0,
    'H2O': 0.0,
}

# Liquid fuel coefficients
LIQUID_H2O_MASS = {
    'methanol': 1.13,
    'gasoline': 1.3,
    '#1 oil': 1.2,
    '#2 oil': 1.12,
    '#4 oil': 1.04,
    '#5 oil': 0.97,
    '#6 oil': 0.84,
}

LIQUID_CO2_MASS = {
    'methanol': 1.38,
    'gasoline': 3.14,
    '#1 oil': 3.17,
    '#2 oil': 3.2,
    '#4 oil': 3.16,
    '#5 oil': 3.24,
    '#6 oil': 3.25,
}

LIQUID_N2_MASS = {
    'methanol': 4.97,
    'gasoline': 11.36,
    '#1 oil': 11.1,
    '#2 oil': 10.95,
    '#4 oil': 10.68,
    '#5 oil': 10.59,
    '#6 oil': 10.25,
}


@dataclass
class GasCompositionMass:
    """
    Gas composition in mass percentages for POC calculations.

    Attributes
    ----------
    air_mass : float, optional
        Air mass percentage (default: 0)
    argon_mass : float, optional
        Argon mass percentage (default: 0)
    methane_mass : float, optional
        Methane mass percentage (default: 0)
    ethane_mass : float, optional
        Ethane mass percentage (default: 0)
    propane_mass : float, optional
        Propane mass percentage (default: 0)
    butane_mass : float, optional
        Butane mass percentage (default: 0)
    pentane_mass : float, optional
        Pentane mass percentage (default: 0)
    hexane_mass : float, optional
        Hexane mass percentage (default: 0)
    co2_mass : float, optional
        CO2 mass percentage (default: 0)
    co_mass : float, optional
        CO mass percentage (default: 0)
    c_mass : float, optional
        Carbon mass percentage (default: 0)
    n2_mass : float, optional
        N2 mass percentage (default: 0)
    h2_mass : float, optional
        H2 mass percentage (default: 0)
    o2_mass : float, optional
        O2 mass percentage (default: 0)
    h2s_mass : float, optional
        H2S mass percentage (default: 0)
    h2o_mass : float, optional
        H2O mass percentage (default: 0)
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


@dataclass
class GasCompositionVolume:
    """
    Gas composition in volume percentages for POC calculations.

    Attributes
    ----------
    air_vol : float, optional
        Air volume percentage (default: 0)
    ammonia_vol : float, optional
        Ammonia volume percentage (default: 0)
    argon_vol : float, optional
        Argon volume percentage (default: 0)
    methane_vol : float, optional
        Methane volume percentage (default: 0)
    ethane_vol : float, optional
        Ethane volume percentage (default: 0)
    propane_vol : float, optional
        Propane volume percentage (default: 0)
    butane_vol : float, optional
        Butane volume percentage (default: 0)
    ibutene_vol : float, optional
        Isobutene volume percentage (default: 0)
    pentane_vol : float, optional
        Pentane volume percentage (default: 0)
    hexane_vol : float, optional
        Hexane volume percentage (default: 0)
    co2_vol : float, optional
        CO2 volume percentage (default: 0)
    co_vol : float, optional
        CO volume percentage (default: 0)
    c_vol : float, optional
        Carbon volume percentage (default: 0)
    n2_vol : float, optional
        N2 volume percentage (default: 0)
    h2_vol : float, optional
        H2 volume percentage (default: 0)
    o2_vol : float, optional
        O2 volume percentage (default: 0)
    h2s_vol : float, optional
        H2S volume percentage (default: 0)
    s_vol : float, optional
        Sulfur volume percentage (default: 0)
    so2_vol : float, optional
        SO2 volume percentage (default: 0)
    h2o_vol : float, optional
        H2O volume percentage (default: 0)
    """
    air_vol: float = 0.0
    ammonia_vol: float = 0.0
    argon_vol: float = 0.0
    methane_vol: float = 0.0
    ethane_vol: float = 0.0
    propane_vol: float = 0.0
    butane_vol: float = 0.0
    ibutene_vol: float = 0.0
    pentane_vol: float = 0.0
    hexane_vol: float = 0.0
    co2_vol: float = 0.0
    co_vol: float = 0.0
    c_vol: float = 0.0
    n2_vol: float = 0.0
    h2_vol: float = 0.0
    o2_vol: float = 0.0
    h2s_vol: float = 0.0
    s_vol: float = 0.0
    so2_vol: float = 0.0
    h2o_vol: float = 0.0


def poc_h2o_mass_gas(
    composition: GasCompositionMass,
    fuel_flow_mass: float,
    humidity: float = 0.0,
    air_flow_mass: float = 0.0
) -> float:
    """
    Calculate water mass in combustion products for gas fuel.

    Parameters
    ----------
    composition : GasCompositionMass
        Gas composition with mass percentages
    fuel_flow_mass : float
        Fuel mass flow rate (lb/hr)
    humidity : float, optional
        Humidity ratio of combustion air (lb H2O / lb dry air, default: 0)
    air_flow_mass : float, optional
        Air mass flow rate (lb/hr, default: 0)

    Returns
    -------
    float
        Water mass in products (lb/hr)

    Examples
    --------
    >>> comp = GasCompositionMass(methane_mass=100.0)
    >>> poc_h2o_mass_gas(comp, fuel_flow_mass=100.0)
    224.6

    Notes
    -----
    H2O = Σ(coeff_i * fuel_flow * comp_i / 100) + humidity * air_flow + fuel_H2O
    Replicates VBA function POC_H2OMass for FuelType="Gas"
    """
    h2o_mass = 0.0

    # Stoichiometric water from fuel combustion
    h2o_mass += GAS_H2O_MASS_COEFF['Air'] * composition.air_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['Argon'] * composition.argon_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['Methane'] * composition.methane_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['Ethane'] * composition.ethane_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['Propane'] * composition.propane_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['Butane'] * composition.butane_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['Pentane'] * composition.pentane_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['Hexane'] * composition.hexane_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['CO2'] * composition.co2_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['CO'] * composition.co_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['C'] * composition.c_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['N2'] * composition.n2_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['H2'] * composition.h2_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['O2'] * composition.o2_mass * fuel_flow_mass / 100
    h2o_mass += GAS_H2O_MASS_COEFF['H2S'] * composition.h2s_mass * fuel_flow_mass / 100

    # Water from humidity in air
    h2o_mass += humidity * air_flow_mass

    # Water already in fuel
    h2o_mass += composition.h2o_mass * fuel_flow_mass / 100

    return h2o_mass


def poc_h2o_mass_liquid(
    fuel_type: str,
    fuel_flow_mass: float,
    humidity: float = 0.0,
    air_flow_mass: float = 0.0
) -> float:
    """
    Calculate water mass in combustion products for liquid fuel.

    Parameters
    ----------
    fuel_type : str
        Liquid fuel type ('methanol', 'gasoline', '#1 oil', '#2 oil', etc.)
    fuel_flow_mass : float
        Fuel mass flow rate (lb/hr)
    humidity : float, optional
        Humidity ratio of combustion air (lb H2O / lb dry air, default: 0)
    air_flow_mass : float, optional
        Air mass flow rate (lb/hr, default: 0)

    Returns
    -------
    float
        Water mass in products (lb/hr)

    Examples
    --------
    >>> poc_h2o_mass_liquid('#2 oil', fuel_flow_mass=100.0)
    112.0

    Notes
    -----
    Replicates VBA function POC_H2OMass for liquid fuels
    """
    fuel_lower = fuel_type.lower()
    if fuel_lower not in LIQUID_H2O_MASS:
        valid_fuels = ', '.join(LIQUID_H2O_MASS.keys())
        raise ValueError(
            f"Unknown liquid fuel type: '{fuel_type}'. "
            f"Valid options: {valid_fuels}"
        )

    coeff = LIQUID_H2O_MASS[fuel_lower]
    return fuel_flow_mass * coeff + humidity * air_flow_mass


def poc_co2_mass_gas(
    composition: GasCompositionMass,
    fuel_flow_mass: float,
    air_flow_mass: float = 0.0
) -> float:
    """
    Calculate CO2 mass in combustion products for gas fuel.

    Parameters
    ----------
    composition : GasCompositionMass
        Gas composition with mass percentages
    fuel_flow_mass : float
        Fuel mass flow rate (lb/hr)
    air_flow_mass : float, optional
        Air mass flow rate (lb/hr, default: 0) - not used but kept for API consistency

    Returns
    -------
    float
        CO2 mass in products (lb/hr)

    Examples
    --------
    >>> comp = GasCompositionMass(methane_mass=100.0)
    >>> poc_co2_mass_gas(comp, fuel_flow_mass=100.0)
    274.3

    Notes
    -----
    CO2 = Σ(coeff_i * fuel_flow * comp_i / 100) + fuel_CO2
    Replicates VBA function POC_CO2Mass for FuelType="Gas"
    """
    co2_mass = 0.0

    # Stoichiometric CO2 from fuel combustion
    co2_mass += GAS_CO2_MASS_COEFF['Air'] * composition.air_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['Argon'] * composition.argon_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['Methane'] * composition.methane_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['Ethane'] * composition.ethane_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['Propane'] * composition.propane_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['Butane'] * composition.butane_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['Pentane'] * composition.pentane_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['Hexane'] * composition.hexane_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['CO2'] * composition.co2_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['CO'] * composition.co_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['C'] * composition.c_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['N2'] * composition.n2_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['H2'] * composition.h2_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['O2'] * composition.o2_mass * fuel_flow_mass / 100
    co2_mass += GAS_CO2_MASS_COEFF['H2S'] * composition.h2s_mass * fuel_flow_mass / 100

    # CO2 already in fuel
    co2_mass += composition.co2_mass * fuel_flow_mass / 100

    return co2_mass


def poc_co2_mass_liquid(
    fuel_type: str,
    fuel_flow_mass: float
) -> float:
    """
    Calculate CO2 mass in combustion products for liquid fuel.

    Parameters
    ----------
    fuel_type : str
        Liquid fuel type
    fuel_flow_mass : float
        Fuel mass flow rate (lb/hr)

    Returns
    -------
    float
        CO2 mass in products (lb/hr)

    Examples
    --------
    >>> poc_co2_mass_liquid('#2 oil', fuel_flow_mass=100.0)
    320.0

    Notes
    -----
    Replicates VBA function POC_CO2Mass for liquid fuels
    """
    fuel_lower = fuel_type.lower()
    if fuel_lower not in LIQUID_CO2_MASS:
        valid_fuels = ', '.join(LIQUID_CO2_MASS.keys())
        raise ValueError(
            f"Unknown liquid fuel type: '{fuel_type}'. "
            f"Valid options: {valid_fuels}"
        )

    coeff = LIQUID_CO2_MASS[fuel_lower]
    return fuel_flow_mass * coeff


def poc_n2_mass_gas(
    composition: GasCompositionMass,
    fuel_flow_mass: float,
    excess_air_mass: float = 0.0,
    air_flow_mass: float = 0.0
) -> float:
    """
    Calculate N2 mass in combustion products for gas fuel.

    Parameters
    ----------
    composition : GasCompositionMass
        Gas composition with mass percentages
    fuel_flow_mass : float
        Fuel mass flow rate (lb/hr)
    excess_air_mass : float, optional
        Total air mass including excess (lb/hr, default: 0)
    air_flow_mass : float, optional
        Stoichiometric air mass flow rate (lb/hr, default: 0)

    Returns
    -------
    float
        N2 mass in products (lb/hr)

    Examples
    --------
    >>> comp = GasCompositionMass(methane_mass=100.0)
    >>> poc_n2_mass_gas(comp, fuel_flow_mass=100.0, excess_air_mass=1800.0, air_flow_mass=1720.0)
    1386.1

    Notes
    -----
    N2 = Σ(coeff_i * fuel_flow * comp_i / 100) + (excess_air - stoich_air) * 0.7686 + fuel_N2
    The factor 0.7686 is the mass fraction of N2 in air (78.08% by volume, 75.5% by mass)
    Replicates VBA function POC_N2Mass for FuelType="Gas"
    """
    n2_mass = 0.0

    # Stoichiometric N2 from fuel combustion
    n2_mass += GAS_N2_MASS_COEFF['Air'] * composition.air_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['Argon'] * composition.argon_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['Methane'] * composition.methane_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['Ethane'] * composition.ethane_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['Propane'] * composition.propane_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['Butane'] * composition.butane_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['Pentane'] * composition.pentane_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['Hexane'] * composition.hexane_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['CO2'] * composition.co2_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['CO'] * composition.co_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['C'] * composition.c_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['N2'] * composition.n2_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['H2'] * composition.h2_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['O2'] * composition.o2_mass * fuel_flow_mass / 100
    n2_mass += GAS_N2_MASS_COEFF['H2S'] * composition.h2s_mass * fuel_flow_mass / 100

    # N2 from excess air (air contains 76.86% N2 by mass)
    n2_mass += (excess_air_mass - air_flow_mass) * 0.7686

    # N2 already in fuel
    n2_mass += composition.n2_mass * fuel_flow_mass / 100

    return n2_mass


def poc_n2_mass_liquid(
    fuel_type: str,
    fuel_flow_mass: float,
    excess_air_mass: float = 0.0,
    air_flow_mass: float = 0.0
) -> float:
    """
    Calculate N2 mass in combustion products for liquid fuel.

    Parameters
    ----------
    fuel_type : str
        Liquid fuel type
    fuel_flow_mass : float
        Fuel mass flow rate (lb/hr)
    excess_air_mass : float, optional
        Total air mass including excess (lb/hr, default: 0)
    air_flow_mass : float, optional
        Stoichiometric air mass flow rate (lb/hr, default: 0)

    Returns
    -------
    float
        N2 mass in products (lb/hr)

    Examples
    --------
    >>> poc_n2_mass_liquid('#2 oil', fuel_flow_mass=100.0, excess_air_mass=1500.0, air_flow_mass=1450.0)
    1133.4

    Notes
    -----
    Replicates VBA function POC_N2Mass for liquid fuels
    """
    fuel_lower = fuel_type.lower()
    if fuel_lower not in LIQUID_N2_MASS:
        valid_fuels = ', '.join(LIQUID_N2_MASS.keys())
        raise ValueError(
            f"Unknown liquid fuel type: '{fuel_type}'. "
            f"Valid options: {valid_fuels}"
        )

    coeff = LIQUID_N2_MASS[fuel_lower]
    return fuel_flow_mass * coeff + (excess_air_mass - air_flow_mass) * 0.7686


def poc_o2_mass(
    fuel_flow_mass: float,
    excess_air_mass: float,
    air_flow_mass: float,
    o2_in_fuel_mass: float = 0.0
) -> float:
    """
    Calculate O2 mass in combustion products.

    Parameters
    ----------
    fuel_flow_mass : float
        Fuel mass flow rate (lb/hr)
    excess_air_mass : float
        Total air mass including excess (lb/hr)
    air_flow_mass : float
        Stoichiometric air mass flow rate (lb/hr)
    o2_in_fuel_mass : float, optional
        O2 mass percentage in fuel (default: 0)

    Returns
    -------
    float
        O2 mass in products (lb/hr)

    Examples
    --------
    >>> poc_o2_mass(fuel_flow_mass=100.0, excess_air_mass=1800.0, air_flow_mass=1720.0)
    18.51

    Notes
    -----
    O2 = (excess_air - stoich_air) * 0.2314 + fuel_O2
    The factor 0.2314 is the mass fraction of O2 in air (20.95% by volume, 23.14% by mass)
    Replicates VBA function POC_O2Mass
    """
    o2_mass = (excess_air_mass - air_flow_mass) * 0.2314
    o2_mass += fuel_flow_mass * o2_in_fuel_mass / 100
    return o2_mass


# Volume-based POC functions

def poc_co2_vol_gas(composition: GasCompositionVolume) -> float:
    """
    Calculate CO2 volume fraction in combustion products for gas fuel.

    Parameters
    ----------
    composition : GasCompositionVolume
        Gas composition with volume percentages

    Returns
    -------
    float
        CO2 volume fraction (%)

    Examples
    --------
    >>> comp = GasCompositionVolume(methane_vol=100.0)
    >>> poc_co2_vol_gas(comp)
    1.0

    Notes
    -----
    CO2_vol = Σ(coeff_i * comp_i / 100) + CO2_in_fuel / 100
    Replicates VBA function POC_CO2Vol for FuelType="Gas"
    """
    co2_vol = 0.0

    co2_vol += GAS_CO2_VOL_COEFF['Air'] * composition.air_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Ammonia'] * composition.ammonia_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Argon'] * composition.argon_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Methane'] * composition.methane_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Ethane'] * composition.ethane_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Propane'] * composition.propane_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Butane'] * composition.butane_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['IButene'] * composition.ibutene_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Pentane'] * composition.pentane_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['Hexane'] * composition.hexane_vol / 100
    co2_vol += composition.co2_vol / 100  # CO2 already in fuel
    co2_vol += GAS_CO2_VOL_COEFF['CO'] * composition.co_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['C'] * composition.c_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['N2'] * composition.n2_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['H2'] * composition.h2_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['O2'] * composition.o2_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['H2S'] * composition.h2s_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['S'] * composition.s_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['SO2'] * composition.so2_vol / 100
    co2_vol += GAS_CO2_VOL_COEFF['H2O'] * composition.h2o_vol / 100

    return co2_vol


def poc_h2o_vol_gas(composition: GasCompositionVolume) -> float:
    """
    Calculate H2O volume fraction in combustion products for gas fuel.

    Parameters
    ----------
    composition : GasCompositionVolume
        Gas composition with volume percentages

    Returns
    -------
    float
        H2O volume fraction (%)

    Examples
    --------
    >>> comp = GasCompositionVolume(methane_vol=100.0)
    >>> poc_h2o_vol_gas(comp)
    2.0

    Notes
    -----
    H2O_vol = Σ(coeff_i * comp_i / 100)
    Replicates VBA function POC_H2OVol for FuelType="Gas"
    """
    h2o_vol = 0.0

    h2o_vol += GAS_H2O_VOL_COEFF['Air'] * composition.air_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Ammonia'] * composition.ammonia_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Argon'] * composition.argon_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Methane'] * composition.methane_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Ethane'] * composition.ethane_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Propane'] * composition.propane_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Butane'] * composition.butane_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['IButene'] * composition.ibutene_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Pentane'] * composition.pentane_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['Hexane'] * composition.hexane_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['CO2'] * composition.co2_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['CO'] * composition.co_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['C'] * composition.c_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['N2'] * composition.n2_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['H2'] * composition.h2_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['O2'] * composition.o2_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['H2S'] * composition.h2s_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['S'] * composition.s_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['SO2'] * composition.so2_vol / 100
    h2o_vol += GAS_H2O_VOL_COEFF['H2O'] * composition.h2o_vol / 100

    return h2o_vol


def poc_n2_vol_gas(composition: GasCompositionVolume) -> float:
    """
    Calculate N2 volume fraction in combustion products for gas fuel.

    Parameters
    ----------
    composition : GasCompositionVolume
        Gas composition with volume percentages

    Returns
    -------
    float
        N2 volume fraction (%)

    Examples
    --------
    >>> comp = GasCompositionVolume(methane_vol=100.0)
    >>> poc_n2_vol_gas(comp)
    7.53

    Notes
    -----
    N2_vol = Σ(coeff_i * comp_i / 100) + N2_in_fuel / 100
    Replicates VBA function POC_N2Vol for FuelType="Gas"
    """
    n2_vol = 0.0

    n2_vol += GAS_N2_VOL_COEFF['Air'] * composition.air_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Ammonia'] * composition.ammonia_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Argon'] * composition.argon_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Methane'] * composition.methane_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Ethane'] * composition.ethane_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Propane'] * composition.propane_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Butane'] * composition.butane_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['IButene'] * composition.ibutene_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Pentane'] * composition.pentane_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['Hexane'] * composition.hexane_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['CO2'] * composition.co2_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['CO'] * composition.co_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['C'] * composition.c_vol / 100
    n2_vol += composition.n2_vol / 100  # N2 already in fuel
    n2_vol += GAS_N2_VOL_COEFF['H2'] * composition.h2_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['O2'] * composition.o2_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['H2S'] * composition.h2s_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['S'] * composition.s_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['SO2'] * composition.so2_vol / 100
    n2_vol += GAS_N2_VOL_COEFF['H2O'] * composition.h2o_vol / 100

    return n2_vol


def poc_so2_vol_gas(composition: GasCompositionVolume) -> float:
    """
    Calculate SO2 volume fraction in combustion products for gas fuel.

    Parameters
    ----------
    composition : GasCompositionVolume
        Gas composition with volume percentages

    Returns
    -------
    float
        SO2 volume fraction (%)

    Examples
    --------
    >>> comp = GasCompositionVolume(h2s_vol=1.0)
    >>> poc_so2_vol_gas(comp)
    0.01

    Notes
    -----
    SO2_vol = Σ(coeff_i * comp_i / 100) + SO2_in_fuel / 100
    Replicates VBA function POC_SO2Vol for FuelType="Gas"
    """
    so2_vol = 0.0

    so2_vol += GAS_SO2_VOL_COEFF['Air'] * composition.air_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Ammonia'] * composition.ammonia_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Argon'] * composition.argon_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Methane'] * composition.methane_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Ethane'] * composition.ethane_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Propane'] * composition.propane_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Butane'] * composition.butane_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['IButene'] * composition.ibutene_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Pentane'] * composition.pentane_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['Hexane'] * composition.hexane_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['CO2'] * composition.co2_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['CO'] * composition.co_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['C'] * composition.c_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['N2'] * composition.n2_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['H2'] * composition.h2_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['O2'] * composition.o2_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['H2S'] * composition.h2s_vol / 100
    so2_vol += GAS_SO2_VOL_COEFF['S'] * composition.s_vol / 100
    so2_vol += composition.so2_vol / 100  # SO2 already in fuel
    so2_vol += GAS_SO2_VOL_COEFF['H2O'] * composition.h2o_vol / 100

    return so2_vol


# VBA-compatible wrapper functions

def POC_H2OMass(
    fuel_type: str,
    fuel_flow_mass: float,
    humidity: float = 0.0,
    air_flow_mass: float = 0.0,
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
    h2o_mass: float = 0
) -> float:
    """
    VBA-compatible wrapper for H2O mass in products.

    Replicates VBA function POC_H2OMass.
    """
    if fuel_type.lower() == "gas":
        composition = GasCompositionMass(
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
        return poc_h2o_mass_gas(composition, fuel_flow_mass, humidity, air_flow_mass)
    else:
        return poc_h2o_mass_liquid(fuel_type, fuel_flow_mass, humidity, air_flow_mass)


def POC_CO2Mass(
    fuel_type: str,
    fuel_flow_mass: float,
    air_flow_mass: float = 0.0,
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
    h2o_mass: float = 0
) -> float:
    """
    VBA-compatible wrapper for CO2 mass in products.

    Replicates VBA function POC_CO2Mass.
    """
    if fuel_type.lower() == "gas":
        composition = GasCompositionMass(
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
        return poc_co2_mass_gas(composition, fuel_flow_mass, air_flow_mass)
    else:
        return poc_co2_mass_liquid(fuel_type, fuel_flow_mass)


def POC_N2Mass(
    fuel_type: str,
    fuel_flow_mass: float,
    excess_air_mass: float = 0.0,
    air_flow_mass: float = 0.0,
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
    h2o_mass: float = 0
) -> float:
    """
    VBA-compatible wrapper for N2 mass in products.

    Replicates VBA function POC_N2Mass.
    """
    if fuel_type.lower() == "gas":
        composition = GasCompositionMass(
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
        return poc_n2_mass_gas(composition, fuel_flow_mass, excess_air_mass, air_flow_mass)
    else:
        return poc_n2_mass_liquid(fuel_type, fuel_flow_mass, excess_air_mass, air_flow_mass)


def POC_O2Mass(
    fuel_flow_mass: float,
    excess_air_mass: float,
    air_flow_mass: float,
    o2_mass: float = 0.0
) -> float:
    """
    VBA-compatible wrapper for O2 mass in products.

    Replicates VBA function POC_O2Mass.
    """
    return poc_o2_mass(fuel_flow_mass, excess_air_mass, air_flow_mass, o2_mass)


def POC_CO2Vol(
    fuel_type: str,
    air_vol: float = 0,
    ammonia_vol: float = 0,
    argon_vol: float = 0,
    methane_vol: float = 0,
    ethane_vol: float = 0,
    propane_vol: float = 0,
    butane_vol: float = 0,
    ibutene_vol: float = 0,
    pentane_vol: float = 0,
    hexane_vol: float = 0,
    co2_vol: float = 0,
    co_vol: float = 0,
    c_vol: float = 0,
    n2_vol: float = 0,
    h2_vol: float = 0,
    o2_vol: float = 0,
    h2s_vol: float = 0,
    s_vol: float = 0,
    so2_vol: float = 0,
    h2o_vol: float = 0
) -> float:
    """
    VBA-compatible wrapper for CO2 volume in products.

    Replicates VBA function POC_CO2Vol.
    """
    if fuel_type.lower() == "gas":
        composition = GasCompositionVolume(
            air_vol=air_vol,
            ammonia_vol=ammonia_vol,
            argon_vol=argon_vol,
            methane_vol=methane_vol,
            ethane_vol=ethane_vol,
            propane_vol=propane_vol,
            butane_vol=butane_vol,
            ibutene_vol=ibutene_vol,
            pentane_vol=pentane_vol,
            hexane_vol=hexane_vol,
            co2_vol=co2_vol,
            co_vol=co_vol,
            c_vol=c_vol,
            n2_vol=n2_vol,
            h2_vol=h2_vol,
            o2_vol=o2_vol,
            h2s_vol=h2s_vol,
            s_vol=s_vol,
            so2_vol=so2_vol,
            h2o_vol=h2o_vol,
        )
        return poc_co2_vol_gas(composition)
    else:
        raise ValueError("Volume-based POC calculations only supported for gas fuels")


def POC_H2OVol(
    fuel_type: str,
    air_vol: float = 0,
    ammonia_vol: float = 0,
    argon_vol: float = 0,
    methane_vol: float = 0,
    ethane_vol: float = 0,
    propane_vol: float = 0,
    butane_vol: float = 0,
    ibutene_vol: float = 0,
    pentane_vol: float = 0,
    hexane_vol: float = 0,
    co2_vol: float = 0,
    co_vol: float = 0,
    c_vol: float = 0,
    n2_vol: float = 0,
    h2_vol: float = 0,
    o2_vol: float = 0,
    h2s_vol: float = 0,
    s_vol: float = 0,
    so2_vol: float = 0,
    h2o_vol: float = 0
) -> float:
    """
    VBA-compatible wrapper for H2O volume in products.

    Replicates VBA function POC_H2OVol.
    """
    if fuel_type.lower() == "gas":
        composition = GasCompositionVolume(
            air_vol=air_vol,
            ammonia_vol=ammonia_vol,
            argon_vol=argon_vol,
            methane_vol=methane_vol,
            ethane_vol=ethane_vol,
            propane_vol=propane_vol,
            butane_vol=butane_vol,
            ibutene_vol=ibutene_vol,
            pentane_vol=pentane_vol,
            hexane_vol=hexane_vol,
            co2_vol=co2_vol,
            co_vol=co_vol,
            c_vol=c_vol,
            n2_vol=n2_vol,
            h2_vol=h2_vol,
            o2_vol=o2_vol,
            h2s_vol=h2s_vol,
            s_vol=s_vol,
            so2_vol=so2_vol,
            h2o_vol=h2o_vol,
        )
        return poc_h2o_vol_gas(composition)
    else:
        raise ValueError("Volume-based POC calculations only supported for gas fuels")


def POC_N2Vol(
    fuel_type: str,
    air_vol: float = 0,
    ammonia_vol: float = 0,
    argon_vol: float = 0,
    methane_vol: float = 0,
    ethane_vol: float = 0,
    propane_vol: float = 0,
    butane_vol: float = 0,
    ibutene_vol: float = 0,
    pentane_vol: float = 0,
    hexane_vol: float = 0,
    co2_vol: float = 0,
    co_vol: float = 0,
    c_vol: float = 0,
    n2_vol: float = 0,
    h2_vol: float = 0,
    o2_vol: float = 0,
    h2s_vol: float = 0,
    s_vol: float = 0,
    so2_vol: float = 0,
    h2o_vol: float = 0
) -> float:
    """
    VBA-compatible wrapper for N2 volume in products.

    Replicates VBA function POC_N2Vol.
    """
    if fuel_type.lower() == "gas":
        composition = GasCompositionVolume(
            air_vol=air_vol,
            ammonia_vol=ammonia_vol,
            argon_vol=argon_vol,
            methane_vol=methane_vol,
            ethane_vol=ethane_vol,
            propane_vol=propane_vol,
            butane_vol=butane_vol,
            ibutene_vol=ibutene_vol,
            pentane_vol=pentane_vol,
            hexane_vol=hexane_vol,
            co2_vol=co2_vol,
            co_vol=co_vol,
            c_vol=c_vol,
            n2_vol=n2_vol,
            h2_vol=h2_vol,
            o2_vol=o2_vol,
            h2s_vol=h2s_vol,
            s_vol=s_vol,
            so2_vol=so2_vol,
            h2o_vol=h2o_vol,
        )
        return poc_n2_vol_gas(composition)
    else:
        raise ValueError("Volume-based POC calculations only supported for gas fuels")


def POC_SO2Vol(
    fuel_type: str,
    air_vol: float = 0,
    ammonia_vol: float = 0,
    argon_vol: float = 0,
    methane_vol: float = 0,
    ethane_vol: float = 0,
    propane_vol: float = 0,
    butane_vol: float = 0,
    ibutene_vol: float = 0,
    pentane_vol: float = 0,
    hexane_vol: float = 0,
    co2_vol: float = 0,
    co_vol: float = 0,
    c_vol: float = 0,
    n2_vol: float = 0,
    h2_vol: float = 0,
    o2_vol: float = 0,
    h2s_vol: float = 0,
    s_vol: float = 0,
    so2_vol: float = 0,
    h2o_vol: float = 0
) -> float:
    """
    VBA-compatible wrapper for SO2 volume in products.

    Replicates VBA function POC_SO2Vol.
    """
    if fuel_type.lower() == "gas":
        composition = GasCompositionVolume(
            air_vol=air_vol,
            ammonia_vol=ammonia_vol,
            argon_vol=argon_vol,
            methane_vol=methane_vol,
            ethane_vol=ethane_vol,
            propane_vol=propane_vol,
            butane_vol=butane_vol,
            ibutene_vol=ibutene_vol,
            pentane_vol=pentane_vol,
            hexane_vol=hexane_vol,
            co2_vol=co2_vol,
            co_vol=co_vol,
            c_vol=c_vol,
            n2_vol=n2_vol,
            h2_vol=h2_vol,
            o2_vol=o2_vol,
            h2s_vol=h2s_vol,
            s_vol=s_vol,
            so2_vol=so2_vol,
            h2o_vol=h2o_vol,
        )
        return poc_so2_vol_gas(composition)
    else:
        raise ValueError("Volume-based POC calculations only supported for gas fuels")
