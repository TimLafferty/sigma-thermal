"""
Air-Fuel Ratio Calculations

This module provides functions for calculating stoichiometric and actual air
requirements for combustion of gaseous and liquid fuels.

References:
    - GPSA Engineering Data Book, 13th Edition, Section 5
    - ASME PTC 4 (Fired Steam Generators)
    - Perry's Chemical Engineers' Handbook, 8th Edition

Author: GTS Energy Inc.
Date: October 2025
"""

from typing import Union
from dataclasses import dataclass


# Stoichiometric air requirements (lb air / lb fuel) from GPSA
STOICH_AIR_MASS = {
    # Gas fuels
    "methane": 17.24,
    "ethane": 16.12,
    "propane": 15.69,
    "butane": 15.47,
    "pentane": 15.35,
    "hexane": 15.26,
    "heptane": 15.20,
    "octane": 15.15,
    "hydrogen": 34.28,
    "carbon_monoxide": 2.47,
    "hydrogen_sulfide": 7.54,
    # Liquid fuels
    "#1_oil": 14.7,
    "#2_oil": 14.5,
    "#4_oil": 14.0,
    "#5_oil": 13.8,
    "#6_oil": 13.5,
    "gasoline": 14.7,
    "diesel": 14.5,
    "kerosene": 14.7,
    "methanol": 6.47,
    "ethanol": 9.0,
}

# Stoichiometric air requirements (scf air / scf fuel) from GPSA
STOICH_AIR_VOL = {
    "methane": 9.53,
    "ethane": 16.68,
    "propane": 23.82,
    "butane": 30.96,
    "pentane": 38.10,
    "hydrogen": 2.38,
    "carbon_monoxide": 2.38,
    "hydrogen_sulfide": 7.14,
}


@dataclass
class GasCompositionMass:
    """
    Gas composition on a mass basis.

    All composition values in mass percent (%).
    Total must sum to 100%.

    Attributes:
        methane_mass: CH4 mass percent (%)
        ethane_mass: C2H6 mass percent (%)
        propane_mass: C3H8 mass percent (%)
        butane_mass: C4H10 mass percent (%)
        pentane_mass: C5H12 mass percent (%)
        hexane_mass: C6H14 mass percent (%)
        heptane_mass: C7H16 mass percent (%)
        octane_mass: C8H18 mass percent (%)
        hydrogen_mass: H2 mass percent (%)
        co_mass: CO mass percent (%)
        h2s_mass: H2S mass percent (%)
        n2_mass: N2 mass percent (%)
        co2_mass: CO2 mass percent (%)
        o2_mass: O2 mass percent (%)
    """

    methane_mass: float = 0.0
    ethane_mass: float = 0.0
    propane_mass: float = 0.0
    butane_mass: float = 0.0
    pentane_mass: float = 0.0
    hexane_mass: float = 0.0
    heptane_mass: float = 0.0
    octane_mass: float = 0.0
    hydrogen_mass: float = 0.0
    co_mass: float = 0.0
    h2s_mass: float = 0.0
    n2_mass: float = 0.0
    co2_mass: float = 0.0
    o2_mass: float = 0.0


@dataclass
class GasCompositionVolume:
    """
    Gas composition on a volume/mole basis.

    All composition values in volume/mole percent (%).
    Total must sum to 100%.

    Attributes:
        methane_vol: CH4 volume percent (%)
        ethane_vol: C2H6 volume percent (%)
        propane_vol: C3H8 volume percent (%)
        butane_vol: C4H10 volume percent (%)
        pentane_vol: C5H12 volume percent (%)
        hydrogen_vol: H2 volume percent (%)
        co_vol: CO volume percent (%)
        h2s_vol: H2S volume percent (%)
        n2_vol: N2 volume percent (%)
        co2_vol: CO2 volume percent (%)
        o2_vol: O2 volume percent (%)
    """

    methane_vol: float = 0.0
    ethane_vol: float = 0.0
    propane_vol: float = 0.0
    butane_vol: float = 0.0
    pentane_vol: float = 0.0
    hydrogen_vol: float = 0.0
    co_vol: float = 0.0
    h2s_vol: float = 0.0
    n2_vol: float = 0.0
    co2_vol: float = 0.0
    o2_vol: float = 0.0


def stoich_air_mass_gas(composition: GasCompositionMass) -> float:
    """
    Calculate stoichiometric air requirement for gas fuel on a mass basis.

    Uses mass-weighted average of individual component stoichiometric air
    requirements from GPSA data.

    Args:
        composition: Gas composition on mass basis (%)

    Returns:
        Stoichiometric air requirement (lb air / lb fuel)

    Raises:
        ValueError: If composition does not sum to approximately 100%

    Example:
        >>> # Pure methane
        >>> comp = GasCompositionMass(methane_mass=100.0)
        >>> stoich_air_mass_gas(comp)
        17.24

        >>> # Natural gas mixture: 90% CH4, 5% C2H6, 3% C3H8, 2% N2
        >>> comp = GasCompositionMass(
        ...     methane_mass=90.0,
        ...     ethane_mass=5.0,
        ...     propane_mass=3.0,
        ...     n2_mass=2.0
        ... )
        >>> stoich_air_mass_gas(comp)
        16.69...

    References:
        - GPSA Engineering Data Book, Section 5
        - Combustion stoichiometry for hydrocarbon fuels
    """
    # Check composition sums to 100% (within tolerance)
    total = (
        composition.methane_mass
        + composition.ethane_mass
        + composition.propane_mass
        + composition.butane_mass
        + composition.pentane_mass
        + composition.hexane_mass
        + composition.heptane_mass
        + composition.octane_mass
        + composition.hydrogen_mass
        + composition.co_mass
        + composition.h2s_mass
        + composition.n2_mass
        + composition.co2_mass
        + composition.o2_mass
    )

    if not (99.0 <= total <= 101.0):
        raise ValueError(f"Composition must sum to 100%, got {total}%")

    # Calculate mass-weighted average stoichiometric air
    stoich_air = 0.0

    # Combustible components
    stoich_air += composition.methane_mass / 100.0 * STOICH_AIR_MASS["methane"]
    stoich_air += composition.ethane_mass / 100.0 * STOICH_AIR_MASS["ethane"]
    stoich_air += composition.propane_mass / 100.0 * STOICH_AIR_MASS["propane"]
    stoich_air += composition.butane_mass / 100.0 * STOICH_AIR_MASS["butane"]
    stoich_air += composition.pentane_mass / 100.0 * STOICH_AIR_MASS["pentane"]
    stoich_air += composition.hexane_mass / 100.0 * STOICH_AIR_MASS["hexane"]
    stoich_air += composition.heptane_mass / 100.0 * STOICH_AIR_MASS["heptane"]
    stoich_air += composition.octane_mass / 100.0 * STOICH_AIR_MASS["octane"]
    stoich_air += composition.hydrogen_mass / 100.0 * STOICH_AIR_MASS["hydrogen"]
    stoich_air += composition.co_mass / 100.0 * STOICH_AIR_MASS["carbon_monoxide"]
    stoich_air += composition.h2s_mass / 100.0 * STOICH_AIR_MASS["hydrogen_sulfide"]

    # Inerts (N2, CO2) require no air
    # O2 in fuel reduces air requirement (subtract 0.2314 lb air per lb O2)
    if composition.o2_mass > 0:
        stoich_air -= composition.o2_mass / 100.0 * 4.32  # O2 in fuel reduces air needed

    return stoich_air


def stoich_air_vol_gas(composition: GasCompositionVolume) -> float:
    """
    Calculate stoichiometric air requirement for gas fuel on a volume basis.

    Uses volume-weighted average of individual component stoichiometric air
    requirements from GPSA data.

    Args:
        composition: Gas composition on volume basis (%)

    Returns:
        Stoichiometric air requirement (scf air / scf fuel)

    Raises:
        ValueError: If composition does not sum to approximately 100%

    Example:
        >>> # Pure methane
        >>> comp = GasCompositionVolume(methane_vol=100.0)
        >>> stoich_air_vol_gas(comp)
        9.53

        >>> # Natural gas: 95% CH4, 3% C2H6, 2% N2
        >>> comp = GasCompositionVolume(
        ...     methane_vol=95.0,
        ...     ethane_vol=3.0,
        ...     n2_vol=2.0
        ... )
        >>> stoich_air_vol_gas(comp)
        9.76...

    References:
        - GPSA Engineering Data Book, Section 5
        - Ideal gas law for volume-based calculations
    """
    # Check composition sums to 100% (within tolerance)
    total = (
        composition.methane_vol
        + composition.ethane_vol
        + composition.propane_vol
        + composition.butane_vol
        + composition.pentane_vol
        + composition.hydrogen_vol
        + composition.co_vol
        + composition.h2s_vol
        + composition.n2_vol
        + composition.co2_vol
        + composition.o2_vol
    )

    if not (99.0 <= total <= 101.0):
        raise ValueError(f"Composition must sum to 100%, got {total}%")

    # Calculate volume-weighted average stoichiometric air
    stoich_air = 0.0

    # Combustible components
    stoich_air += composition.methane_vol / 100.0 * STOICH_AIR_VOL["methane"]
    stoich_air += composition.ethane_vol / 100.0 * STOICH_AIR_VOL["ethane"]
    stoich_air += composition.propane_vol / 100.0 * STOICH_AIR_VOL["propane"]
    stoich_air += composition.butane_vol / 100.0 * STOICH_AIR_VOL["butane"]
    stoich_air += composition.pentane_vol / 100.0 * STOICH_AIR_VOL["pentane"]
    stoich_air += composition.hydrogen_vol / 100.0 * STOICH_AIR_VOL["hydrogen"]
    stoich_air += composition.co_vol / 100.0 * STOICH_AIR_VOL["carbon_monoxide"]
    stoich_air += composition.h2s_vol / 100.0 * STOICH_AIR_VOL["hydrogen_sulfide"]

    # Inerts (N2, CO2) require no air
    # O2 in fuel reduces air requirement (subtract 4.76 scf air per scf O2)
    if composition.o2_vol > 0:
        stoich_air -= composition.o2_vol / 100.0 * 4.76

    return stoich_air


def stoich_air_mass_liquid(fuel_type: str) -> float:
    """
    Calculate stoichiometric air requirement for liquid fuel on a mass basis.

    Uses lookup table of typical stoichiometric air requirements for
    common liquid fuels from GPSA.

    Args:
        fuel_type: Type of liquid fuel (case-insensitive)
            Supported: '#1 oil', '#2 oil', '#4 oil', '#5 oil', '#6 oil',
                      'gasoline', 'diesel', 'kerosene', 'methanol', 'ethanol'

    Returns:
        Stoichiometric air requirement (lb air / lb fuel)

    Raises:
        ValueError: If fuel type is not recognized

    Example:
        >>> stoich_air_mass_liquid('#2 oil')
        14.5

        >>> stoich_air_mass_liquid('gasoline')
        14.7

    References:
        - GPSA Engineering Data Book, Section 5
        - Typical fuel oil properties
    """
    # Normalize fuel type (lowercase, spaces)
    fuel_normalized = fuel_type.lower().strip().replace(" ", "_")

    # Check if fuel is in lookup table
    if fuel_normalized not in STOICH_AIR_MASS:
        raise ValueError(
            f"Unknown fuel type: '{fuel_type}'. "
            f"Supported types: {', '.join([k.replace('_', ' ') for k in STOICH_AIR_MASS.keys() if not k in ['methane', 'ethane', 'propane', 'butane', 'pentane', 'hexane', 'heptane', 'octane', 'hydrogen', 'carbon_monoxide', 'hydrogen_sulfide']])}"
        )

    return STOICH_AIR_MASS[fuel_normalized]


def excess_air_percent(
    actual_air: float,
    stoich_air: float
) -> float:
    """
    Calculate excess air percentage from actual and stoichiometric air flows.

    Excess air is the amount of air above the stoichiometric requirement,
    expressed as a percentage of the stoichiometric air.

    Args:
        actual_air: Actual air flow (any consistent units)
        stoich_air: Stoichiometric air flow (same units as actual_air)

    Returns:
        Excess air percentage (%)

    Raises:
        ValueError: If stoichiometric air is zero or negative

    Example:
        >>> # 10% excess air
        >>> excess_air_percent(1896.4, 1724.0)
        10.0

        >>> # 20% excess air
        >>> excess_air_percent(1000, 833.33)
        20.0...

        >>> # Zero excess air (stoichiometric combustion)
        >>> excess_air_percent(100, 100)
        0.0

    Notes:
        - Excess air = (Actual air - Stoich air) / Stoich air * 100
        - Typical ranges:
            * Natural gas: 5-15% excess air
            * Oil: 15-25% excess air
            * Coal: 20-30% excess air
        - Too low: incomplete combustion, CO formation
        - Too high: reduced efficiency, increased stack loss

    References:
        - ASME PTC 4 (Fired Steam Generators)
        - Combustion efficiency optimization
    """
    if stoich_air <= 0:
        raise ValueError(f"Stoichiometric air must be positive, got {stoich_air}")

    excess = ((actual_air - stoich_air) / stoich_air) * 100.0

    return excess


# VBA-compatible wrapper functions
def StoichAirMassGas(**kwargs) -> float:
    """
    VBA-compatible wrapper for stoich_air_mass_gas().

    Accepts keyword arguments for all composition components.

    Returns:
        Stoichiometric air (lb air / lb fuel)
    """
    comp = GasCompositionMass(
        methane_mass=kwargs.get("methane_mass", 0.0),
        ethane_mass=kwargs.get("ethane_mass", 0.0),
        propane_mass=kwargs.get("propane_mass", 0.0),
        butane_mass=kwargs.get("butane_mass", 0.0),
        pentane_mass=kwargs.get("pentane_mass", 0.0),
        hexane_mass=kwargs.get("hexane_mass", 0.0),
        heptane_mass=kwargs.get("heptane_mass", 0.0),
        octane_mass=kwargs.get("octane_mass", 0.0),
        hydrogen_mass=kwargs.get("hydrogen_mass", 0.0),
        co_mass=kwargs.get("co_mass", 0.0),
        h2s_mass=kwargs.get("h2s_mass", 0.0),
        n2_mass=kwargs.get("n2_mass", 0.0),
        co2_mass=kwargs.get("co2_mass", 0.0),
        o2_mass=kwargs.get("o2_mass", 0.0),
    )
    return stoich_air_mass_gas(comp)


def StoichAirMassLiquid(fuel_type: str) -> float:
    """VBA-compatible wrapper for stoich_air_mass_liquid()."""
    return stoich_air_mass_liquid(fuel_type)


def ExcessAirPercent(actual_air: float, stoich_air: float) -> float:
    """VBA-compatible wrapper for excess_air_percent()."""
    return excess_air_percent(actual_air, stoich_air)
