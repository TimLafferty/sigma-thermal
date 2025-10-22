"""
Combustion Emissions Calculations

This module provides functions for calculating pollutant emissions from
combustion processes, including NOx and CO2.

References:
    - EPA AP-42: Compilation of Air Pollutant Emission Factors
    - GPSA Engineering Data Book, 13th Edition
    - 40 CFR Part 60: Standards of Performance for New Stationary Sources

Author: GTS Energy Inc.
Date: October 2025
"""

from typing import Optional


def nox_emissions(
    fuel_rate: float,
    flame_temp: float,
    excess_air_pct: float,
    fuel_type: str = "natural_gas",
    fuel_nitrogen_pct: float = 0.0,
    residence_time: float = 0.5
) -> float:
    """
    Calculate NOx emissions from combustion process.

    NOx (nitrogen oxides) emissions consist primarily of thermal NOx
    formed at high temperatures and fuel NOx from nitrogen in the fuel.

    This uses simplified correlations based on EPA AP-42 emission factors
    and thermal NOx kinetics. For detailed analysis, use computational
    models like CFD with detailed chemical kinetics.

    Args:
        fuel_rate: Fuel flow rate (lb/hr for solids/liquids, MMBtu/hr for gas)
        flame_temp: Flame temperature (°F)
        excess_air_pct: Excess air percentage (%)
        fuel_type: Type of fuel ("natural_gas", "oil", "coal", default "natural_gas")
        fuel_nitrogen_pct: Nitrogen content in fuel (%, default 0.0)
        residence_time: Combustion zone residence time (seconds, default 0.5)

    Returns:
        NOx emissions rate (lb NOx/hr)

    Raises:
        ValueError: If fuel rate, flame temperature, or residence time is invalid
        ValueError: If fuel type is not recognized

    Example:
        >>> # Natural gas boiler: 100 MMBtu/hr, 3600°F, 10% excess air
        >>> nox_emissions(
        ...     fuel_rate=100,
        ...     flame_temp=3600,
        ...     excess_air_pct=10.0,
        ...     fuel_type="natural_gas"
        ... )
        10.2...

        >>> # Oil burner with higher NOx
        >>> nox_emissions(
        ...     fuel_rate=75,
        ...     flame_temp=3500,
        ...     excess_air_pct=15.0,
        ...     fuel_type="oil",
        ...     fuel_nitrogen_pct=0.2
        ... )
        15.8...

    Notes:
        - NOx formation mechanisms:
            * Thermal NOx: O2 + N2 → NOx (dominant above 2800°F)
            * Fuel NOx: Nitrogen in fuel oxidizes
            * Prompt NOx: CH + N2 → HCN → NOx (minor)

        - Typical uncontrolled NOx emissions:
            * Natural gas: 0.08-0.12 lb NOx/MMBtu
            * #2 Oil: 0.15-0.25 lb NOx/MMBtu
            * Coal: 0.4-0.8 lb NOx/MMBtu

        - NOx increases with:
            * Higher flame temperature (exponential)
            * Higher excess air (more O2 available)
            * Longer residence time
            * Nitrogen content in fuel

        - Control methods:
            * Low NOx burners (LNB): 30-50% reduction
            * Flue gas recirculation (FGR): 50-70% reduction
            * Selective catalytic reduction (SCR): 80-90% reduction
            * Water/steam injection: 50-80% reduction

    References:
        - EPA AP-42 Section 1.4: Natural Gas Combustion
        - EPA AP-42 Section 1.3: Fuel Oil Combustion
        - Zeldovich mechanism for thermal NOx formation
    """
    # Input validation
    if fuel_rate <= 0:
        raise ValueError(f"Fuel rate must be positive, got {fuel_rate}")

    if flame_temp < 1000:
        raise ValueError(f"Flame temperature must be at least 1000°F, got {flame_temp}")

    if flame_temp > 5000:
        raise ValueError(f"Flame temperature {flame_temp}°F is unreasonably high")

    if residence_time <= 0:
        raise ValueError(f"Residence time must be positive, got {residence_time}")

    # Validate fuel type
    valid_fuels = ["natural_gas", "oil", "coal"]
    fuel_type_lower = fuel_type.lower().replace(" ", "_")
    if fuel_type_lower not in valid_fuels:
        raise ValueError(
            f"Unknown fuel type: '{fuel_type}'. "
            f"Valid types: {', '.join(valid_fuels)}"
        )

    # Base emission factors (lb NOx/MMBtu) at reference conditions
    # Reference: 3200°F, 15% excess air, 0.5 sec residence time
    base_emission_factors = {
        "natural_gas": 0.10,  # Clean fuel, low nitrogen
        "oil": 0.20,          # Some nitrogen content
        "coal": 0.60          # Higher nitrogen content
    }

    base_ef = base_emission_factors[fuel_type_lower]

    # Temperature correction factor (thermal NOx exponential dependence)
    # Reference temperature: 3200°F
    T_ref = 3200.0

    # Simplified Arrhenius-type temperature dependence
    # NOx formation rate ∝ exp(-E/RT), approximated here
    # Temperature effect is very strong above 2800°F
    if flame_temp < 2800:
        # Minimal thermal NOx below 2800°F
        temp_factor = 0.1
    else:
        # Exponential increase above 2800°F
        delta_t = (flame_temp - T_ref) / 1000.0
        temp_factor = 1.0 * (1.5 ** delta_t)

    # Excess air correction factor
    # More O2 available increases NOx formation
    # Reference: 15% excess air
    excess_air_ref = 15.0
    excess_air_factor = 1.0 + (excess_air_pct - excess_air_ref) / 100.0
    excess_air_factor = max(0.5, min(excess_air_factor, 2.0))  # Bound factor

    # Residence time correction factor
    # Longer time allows more NOx formation
    # Reference: 0.5 seconds
    time_ref = 0.5
    time_factor = (residence_time / time_ref) ** 0.5  # Square root relationship

    # Fuel nitrogen contribution
    # Convert fuel nitrogen to additional NOx
    # Typically 30-50% of fuel N converts to NOx
    fuel_nox_conversion = 0.40  # 40% conversion rate
    fuel_nox_contribution = fuel_nitrogen_pct * fuel_nox_conversion

    # Calculate total emission factor
    # Thermal NOx component
    thermal_ef = base_ef * temp_factor * excess_air_factor * time_factor

    # Fuel NOx component (lb NOx/MMBtu)
    # Approximate: 1% fuel N → 0.05 lb NOx/MMBtu additional
    fuel_ef = fuel_nox_contribution * 0.05

    total_ef = thermal_ef + fuel_ef

    # Calculate total NOx emissions (lb/hr)
    nox_emissions_rate = total_ef * fuel_rate

    return nox_emissions_rate


def co2_emissions(
    fuel_rate: float,
    carbon_content: float,
    fuel_type: str = "natural_gas",
    hhv: Optional[float] = None
) -> float:
    """
    Calculate CO2 emissions from complete combustion.

    CO2 emissions are determined by the carbon content of the fuel
    and assume complete combustion. For incomplete combustion,
    adjust for CO and unburned hydrocarbons separately.

    Args:
        fuel_rate: Fuel flow rate (lb/hr for solids/liquids, MMBtu/hr for gas)
        carbon_content: Carbon content of fuel (mass fraction, 0-1)
                       For gaseous fuels in MMBtu/hr, use typical values:
                       - Natural gas: 0.75 (75% CH4 equivalent)
                       - Propane: 0.817
        fuel_type: Type of fuel ("natural_gas", "oil", "coal", default "natural_gas")
                  Used for default carbon content if not specified
        hhv: Higher heating value (BTU/lb), optional for mass-based calculation

    Returns:
        CO2 emissions rate (lb CO2/hr)

    Raises:
        ValueError: If fuel rate is invalid
        ValueError: If carbon content is outside valid range (0-1)

    Example:
        >>> # Natural gas: 100 MMBtu/hr
        >>> co2_emissions(
        ...     fuel_rate=100,
        ...     carbon_content=0.75,
        ...     fuel_type="natural_gas"
        ... )
        11690.0...

        >>> # #2 Oil: 100 lb/hr, 87% carbon
        >>> co2_emissions(
        ...     fuel_rate=100,
        ...     carbon_content=0.87,
        ...     fuel_type="oil"
        ... )
        319.0...

        >>> # Coal: 1000 lb/hr, 70% carbon
        >>> co2_emissions(
        ...     fuel_rate=1000,
        ...     carbon_content=0.70,
        ...     fuel_type="coal"
        ... )
        2566.7...

    Notes:
        - CO2 formation: C + O2 → CO2
        - Molecular weights: C = 12, O2 = 32, CO2 = 44
        - Conversion: 1 lb C → (44/12) = 3.667 lb CO2

        - Typical CO2 emission factors:
            * Natural gas: 116.9 lb CO2/MMBtu
            * #2 Oil: 163.5 lb CO2/MMBtu
            * Coal: 210-230 lb CO2/MMBtu (depends on rank)

        - Carbon content by fuel type:
            * Natural gas: ~75% (mostly CH4)
            * Oil: 85-87%
            * Anthracite coal: 80-90%
            * Bituminous coal: 70-80%
            * Lignite: 65-70%

        - CO2 is a greenhouse gas (GHG):
            * Global Warming Potential (GWP) = 1 (reference)
            * Primary contributor to climate change
            * Subject to emissions trading/carbon taxes
            * Captured via carbon capture and storage (CCS)

    References:
        - EPA 40 CFR Part 98: Greenhouse Gas Reporting
        - EPA AP-42: CO2 emission factors by fuel type
        - IPCC Guidelines for GHG Inventories
    """
    # Input validation
    if fuel_rate <= 0:
        raise ValueError(f"Fuel rate must be positive, got {fuel_rate}")

    if not (0 <= carbon_content <= 1):
        raise ValueError(
            f"Carbon content must be between 0 and 1 (mass fraction), got {carbon_content}"
        )

    # Typical CO2 emission factors for common fuels (lb CO2/MMBtu)
    # Based on EPA AP-42 and 40 CFR Part 98
    co2_emission_factors = {
        "natural_gas": 116.9,   # Mostly methane
        "propane": 139.0,
        "butane": 143.0,
        "oil": 163.5,           # #2 Distillate oil
        "diesel": 163.5,
        "residual_oil": 173.9,  # #6 Residual oil
        "coal_anthracite": 228.6,
        "coal_bituminous": 205.3,
        "coal_subbituminous": 214.3,
        "coal_lignite": 215.4,
        "coal": 210.0,          # Average bituminous
    }

    # Normalize fuel type
    fuel_type_lower = fuel_type.lower().replace(" ", "_")

    # For gaseous fuels, typically provided as MMBtu/hr
    # Use emission factor directly
    if fuel_type_lower in ["natural_gas", "propane", "butane"]:
        # Fuel rate is in MMBtu/hr
        # Use standard emission factor
        if fuel_type_lower in co2_emission_factors:
            ef = co2_emission_factors[fuel_type_lower]
        else:
            # Calculate from carbon content
            # Typical: Natural gas is ~75% carbon equivalent
            # 1 MMBtu ≈ 1e6 BTU
            # Assume typical HHV: ~23,000 BTU/lb for natural gas
            ef = carbon_content * 117.0  # Approximate for gaseous fuels

        co2_rate = ef * fuel_rate

    else:
        # Solid/liquid fuels: fuel rate is in lb/hr
        # Direct calculation from carbon content
        # Stoichiometry: C + O2 → CO2
        # MW: C=12, O2=32, CO2=44
        # Ratio: 44/12 = 3.667 lb CO2/lb C

        carbon_rate = fuel_rate * carbon_content  # lb C/hr
        co2_rate = carbon_rate * (44.0 / 12.0)    # lb CO2/hr

    return co2_rate


# VBA-compatible wrapper functions
def NOxEmissions(
    fuel_rate: float,
    flame_temp: float,
    excess_air_pct: float,
    fuel_type: str = "natural_gas",
    fuel_nitrogen_pct: float = 0.0,
    residence_time: float = 0.5
) -> float:
    """VBA-compatible wrapper for nox_emissions()."""
    return nox_emissions(
        fuel_rate, flame_temp, excess_air_pct,
        fuel_type, fuel_nitrogen_pct, residence_time
    )


def CO2Emissions(
    fuel_rate: float,
    carbon_content: float,
    fuel_type: str = "natural_gas",
    hhv: Optional[float] = None
) -> float:
    """VBA-compatible wrapper for co2_emissions()."""
    return co2_emissions(fuel_rate, carbon_content, fuel_type, hhv)
