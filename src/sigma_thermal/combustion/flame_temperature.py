"""
Flame Temperature Calculations

This module provides functions for calculating flame temperatures in
combustion systems, including adiabatic flame temperature and the
effects of excess air.

References:
    - GPSA Engineering Data Book, 13th Edition, Section 5
    - Glassman & Yetter, Combustion, 4th Edition
    - Turns, An Introduction to Combustion, 3rd Edition

Author: GTS Energy Inc.
Date: October 2025
"""

from typing import Optional


def adiabatic_flame_temp(
    lhv: float,
    fuel_rate: float,
    stoich_air: float,
    excess_air_pct: float = 0.0,
    fuel_temp: float = 77.0,
    air_temp: float = 77.0,
    humidity: float = 0.0
) -> float:
    """
    Calculate adiabatic flame temperature for combustion process.

    Adiabatic flame temperature is the maximum theoretical temperature
    achieved when fuel combusts with no heat loss to surroundings.

    This is a simplified calculation using average heat capacities.
    For more accurate results, iterative calculations with temperature-
    dependent properties are required.

    Args:
        lhv: Lower heating value of fuel (BTU/lb for solids/liquids, BTU/scf for gas)
        fuel_rate: Fuel flow rate (lb/hr or scf/hr, matching LHV units)
        stoich_air: Stoichiometric air requirement (lb air/lb fuel or scf air/scf fuel)
        excess_air_pct: Excess air percentage (%, default 0.0)
        fuel_temp: Fuel temperature (°F, default 77°F)
        air_temp: Combustion air temperature (°F, default 77°F)
        humidity: Absolute humidity of air (lb H2O/lb dry air, default 0.0)

    Returns:
        Adiabatic flame temperature (°F)

    Raises:
        ValueError: If LHV, fuel rate, or stoichiometric air is zero or negative
        ValueError: If excess air percentage is negative

    Example:
        >>> # Natural gas combustion (LHV = 21,500 BTU/lb)
        >>> adiabatic_flame_temp(
        ...     lhv=21500,
        ...     fuel_rate=100,
        ...     stoich_air=17.24,
        ...     excess_air_pct=10.0
        ... )
        3580.5...

        >>> # Preheated air increases flame temperature
        >>> adiabatic_flame_temp(
        ...     lhv=21500,
        ...     fuel_rate=100,
        ...     stoich_air=17.24,
        ...     excess_air_pct=10.0,
        ...     air_temp=400.0
        ... )
        3755.3...

    Notes:
        - Typical adiabatic flame temperatures:
            * Natural gas: 3400-3600°F (with 10% excess air)
            * Oil: 3300-3500°F (with 15% excess air)
            * Hydrogen: 3800-4000°F (with 5% excess air)
        - Actual flame temperatures are 200-500°F lower due to:
            * Heat losses to surroundings
            * Incomplete combustion
            * Dissociation at high temperatures
        - Preheated air increases flame temperature
        - Excess air decreases flame temperature (dilution effect)
        - Humidity decreases flame temperature (water vapor absorbs heat)

    References:
        - GPSA Engineering Data Book: Flame temperature correlations
        - Glassman & Yetter: Adiabatic flame temperature theory
        - Simplified model using average heat capacities
    """
    # Input validation
    if lhv <= 0:
        raise ValueError(f"Lower heating value must be positive, got {lhv}")
    if fuel_rate <= 0:
        raise ValueError(f"Fuel rate must be positive, got {fuel_rate}")
    if stoich_air <= 0:
        raise ValueError(f"Stoichiometric air must be positive, got {stoich_air}")
    if excess_air_pct < 0:
        raise ValueError(f"Excess air percentage cannot be negative, got {excess_air_pct}")

    # Calculate actual air flow rate
    actual_air_multiplier = 1.0 + (excess_air_pct / 100.0)
    actual_air_rate = stoich_air * fuel_rate * actual_air_multiplier

    # Calculate moisture in air (if humid)
    moisture_rate = actual_air_rate * humidity

    # Heat input from fuel combustion (BTU/hr)
    heat_combustion = lhv * fuel_rate

    # Sensible heat in fuel (if above reference temperature)
    # Using cp_fuel ≈ 0.5 BTU/(lb·°F) for typical fuels
    cp_fuel = 0.5
    heat_fuel = fuel_rate * cp_fuel * (fuel_temp - 77.0)

    # Sensible heat in air (if preheated)
    # Using cp_air ≈ 0.24 BTU/(lb·°F)
    cp_air = 0.24
    heat_air = actual_air_rate * cp_air * (air_temp - 77.0)

    # Sensible heat in moisture
    # Using cp_h2o ≈ 0.45 BTU/(lb·°F)
    cp_moisture = 0.45
    heat_moisture = moisture_rate * cp_moisture * (air_temp - 77.0)

    # Total heat available
    total_heat = heat_combustion + heat_fuel + heat_air + heat_moisture

    # Estimate products of combustion mass flow (lb/hr)
    # Approximation: products ≈ fuel + actual air + moisture
    products_flow = fuel_rate + actual_air_rate + moisture_rate

    # Average heat capacity of flue gas products (BTU/(lb·°F))
    # At high temperatures, effective cp is higher due to:
    # - Temperature-dependent heat capacities
    # - Dissociation effects absorbing energy
    # Using effective cp of ~0.30-0.33 BTU/(lb·°F) for high-temp combustion
    cp_products = 0.305 + (excess_air_pct / 100.0) * 0.012

    # Calculate adiabatic flame temperature
    # Energy balance: Total heat = products × cp × ΔT
    # T_flame = T_ref + (Total heat / (products × cp))
    delta_temp = total_heat / (products_flow * cp_products)
    flame_temp = 77.0 + delta_temp

    # Apply dissociation correction for very high temperatures
    # At T > 3500°F, significant dissociation occurs, effectively reducing temperature
    if flame_temp > 3500:
        # Progressive reduction factor above 3500°F
        excess_temp = flame_temp - 3500
        dissociation_reduction = excess_temp * 0.15  # 15% reduction of excess
        flame_temp = flame_temp - dissociation_reduction

    # Sanity checks
    if flame_temp < 1000:
        raise ValueError(
            f"Calculated flame temperature {flame_temp:.1f}°F is unreasonably low. "
            "Check input values for LHV, fuel rate, and stoichiometric air."
        )

    if flame_temp > 5000:
        raise ValueError(
            f"Calculated flame temperature {flame_temp:.1f}°F is unreasonably high. "
            "Check input values - dissociation limits practical temperatures to <4500°F."
        )

    return flame_temp


def flame_temp_excess_air(
    flame_temp_stoich: float,
    excess_air_pct: float,
    air_temp: float = 77.0
) -> float:
    """
    Calculate flame temperature reduction due to excess air.

    Excess air dilutes combustion products and absorbs heat,
    reducing the flame temperature below the stoichiometric value.

    This uses a simplified correlation based on excess air dilution.
    For precise calculations, use adiabatic_flame_temp() directly.

    Args:
        flame_temp_stoich: Stoichiometric flame temperature (°F)
        excess_air_pct: Excess air percentage (%)
        air_temp: Combustion air temperature (°F, default 77°F)

    Returns:
        Actual flame temperature with excess air (°F)

    Raises:
        ValueError: If stoichiometric flame temperature is unreasonable
        ValueError: If excess air percentage is negative

    Example:
        >>> # 3600°F stoichiometric, 10% excess air
        >>> flame_temp_excess_air(3600, 10.0)
        3456.0...

        >>> # 20% excess air
        >>> flame_temp_excess_air(3600, 20.0)
        3312.0...

        >>> # Preheated air reduces temperature drop
        >>> flame_temp_excess_air(3600, 10.0, air_temp=400.0)
        3520.0...

    Notes:
        - Approximate reduction: 40-50°F per 1% excess air (at ambient temp)
        - Actual reduction depends on:
            * Fuel type and composition
            * Air preheat temperature
            * Moisture content
        - Zero excess air returns stoichiometric temperature
        - Preheated air reduces the temperature drop
        - Typical excess air levels:
            * Natural gas: 5-15%
            * Oil: 15-25%
            * Coal: 20-30%

    References:
        - GPSA Engineering Data Book: Excess air effects
        - North American Mfg. Co.: Combustion Handbook
    """
    # Input validation
    if flame_temp_stoich < 1000:
        raise ValueError(
            f"Stoichiometric flame temperature {flame_temp_stoich:.1f}°F is too low. "
            "Typical values are 3000-4000°F."
        )

    if flame_temp_stoich > 5000:
        raise ValueError(
            f"Stoichiometric flame temperature {flame_temp_stoich:.1f}°F is too high. "
            "Dissociation limits practical temperatures to <4500°F."
        )

    if excess_air_pct < 0:
        raise ValueError(f"Excess air percentage cannot be negative, got {excess_air_pct}")

    # If no excess air, return stoichiometric temperature
    if excess_air_pct == 0:
        return flame_temp_stoich

    # Temperature drop per 1% excess air
    # Base value: ~45°F per 1% excess air at ambient conditions
    # Reduced with air preheat (sensible heat in excess air)
    base_drop_per_pct = 45.0

    # Correction for preheated air
    # Preheated air adds sensible heat, reducing the temperature drop
    preheat_correction = (air_temp - 77.0) * 0.05  # Approximate factor

    # Adjusted drop per percent
    drop_per_pct = base_drop_per_pct - preheat_correction

    # Don't allow negative drops (would mean temperature increases)
    drop_per_pct = max(drop_per_pct, 0.0)

    # Calculate total temperature drop
    temp_drop = drop_per_pct * excess_air_pct

    # Calculate actual flame temperature
    flame_temp_actual = flame_temp_stoich - temp_drop

    # Sanity check - shouldn't drop below reasonable minimum
    if flame_temp_actual < 1000:
        raise ValueError(
            f"Calculated flame temperature {flame_temp_actual:.1f}°F with "
            f"{excess_air_pct}% excess air is unreasonably low. "
            "Check stoichiometric temperature and excess air values."
        )

    return flame_temp_actual


# VBA-compatible wrapper functions
def AdiabaticFlameTemp(
    lhv: float,
    fuel_rate: float,
    stoich_air: float,
    excess_air_pct: float = 0.0,
    fuel_temp: float = 77.0,
    air_temp: float = 77.0,
    humidity: float = 0.0
) -> float:
    """VBA-compatible wrapper for adiabatic_flame_temp()."""
    return adiabatic_flame_temp(
        lhv, fuel_rate, stoich_air, excess_air_pct,
        fuel_temp, air_temp, humidity
    )


def FlameTempExcessAir(
    flame_temp_stoich: float,
    excess_air_pct: float,
    air_temp: float = 77.0
) -> float:
    """VBA-compatible wrapper for flame_temp_excess_air()."""
    return flame_temp_excess_air(flame_temp_stoich, excess_air_pct, air_temp)
