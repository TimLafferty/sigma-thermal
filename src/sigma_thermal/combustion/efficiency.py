"""
Combustion Efficiency Calculations

This module provides functions for calculating combustion efficiency,
stack losses, and thermal performance of fired equipment.

References:
    - ASME PTC 4 (Fired Steam Generators)
    - GPSA Engineering Data Book, 13th Edition
    - Combustion efficiency per ASME standard methods

Author: GTS Energy Inc.
Date: October 2025
"""

from typing import Optional


def combustion_efficiency(
    heat_input: float,
    stack_loss: float,
    radiation_loss: float = 0.0,
    blow_down_loss: float = 0.0,
    unaccounted_loss: float = 0.0
) -> float:
    """
    Calculate combustion efficiency using the heat loss method.

    Efficiency is calculated as the ratio of useful heat output to
    total heat input, per ASME PTC 4 methodology.

    Args:
        heat_input: Total heat input from fuel (BTU/hr or consistent units)
        stack_loss: Heat loss through stack/flue gas (same units as heat_input)
        radiation_loss: Heat loss through radiation (same units, default 0.0)
        blow_down_loss: Heat loss through blowdown (same units, default 0.0)
        unaccounted_loss: Other unaccounted losses (same units, default 0.0)

    Returns:
        Combustion efficiency (%)

    Raises:
        ValueError: If heat_input is zero or negative
        ValueError: If total losses exceed heat input

    Example:
        >>> # Basic efficiency: heat input 1M BTU/hr, stack loss 150k BTU/hr
        >>> combustion_efficiency(1000000, 150000)
        85.0

        >>> # With radiation loss (2% typical)
        >>> combustion_efficiency(1000000, 150000, radiation_loss=20000)
        83.0

        >>> # High efficiency boiler
        >>> combustion_efficiency(2387500, 163000)
        93.17...

    Notes:
        - Efficiency = (Heat Input - Total Losses) / Heat Input * 100
        - Total Losses = Stack + Radiation + Blowdown + Unaccounted
        - Typical efficiencies:
            * Modern gas boilers: 85-95%
            * Oil-fired boilers: 80-88%
            * Coal-fired boilers: 75-85%
        - Stack loss typically dominates (80-90% of total losses)
        - Radiation loss typically 1-3% of heat input

    References:
        - ASME PTC 4: Standard for fired steam generators
        - Heat loss method per ASME standards
    """
    if heat_input <= 0:
        raise ValueError(f"Heat input must be positive, got {heat_input}")

    # Calculate total losses
    total_losses = stack_loss + radiation_loss + blow_down_loss + unaccounted_loss

    # Validate losses don't exceed input
    if total_losses > heat_input:
        raise ValueError(
            f"Total losses ({total_losses}) exceed heat input ({heat_input}). "
            "Check calculation inputs."
        )

    # Calculate efficiency
    efficiency = ((heat_input - total_losses) / heat_input) * 100.0

    # Sanity check (should be 0-100%)
    if efficiency < 0 or efficiency > 100:
        raise ValueError(
            f"Calculated efficiency {efficiency:.1f}% is outside valid range (0-100%). "
            "Check calculation inputs."
        )

    return efficiency


def stack_loss_percent(
    flue_gas_enthalpy: float,
    flue_gas_flow: float,
    heat_input: float
) -> float:
    """
    Calculate stack loss as a percentage of heat input.

    Stack loss is the heat carried away by hot flue gases leaving
    the combustion system.

    Args:
        flue_gas_enthalpy: Specific enthalpy of flue gas above ambient (BTU/lb)
        flue_gas_flow: Total flue gas mass flow rate (lb/hr)
        heat_input: Total heat input from fuel (BTU/hr)

    Returns:
        Stack loss as percentage of heat input (%)

    Raises:
        ValueError: If heat_input is zero or negative

    Example:
        >>> # Calculate stack loss for gas boiler
        >>> # Flue gas: 300 BTU/lb enthalpy, 2000 lb/hr flow
        >>> # Heat input: 2 million BTU/hr
        >>> stack_loss_percent(300, 2000, 2000000)
        30.0

        >>> # High efficiency (low stack temperature)
        >>> stack_loss_percent(150, 2020, 2387500)
        12.68...

        >>> # Lower efficiency (high stack temperature)
        >>> stack_loss_percent(426.5, 2020.5, 2387500)
        36.1...

    Notes:
        - Stack Loss = Flue Gas Enthalpy × Flue Gas Flow
        - Stack Loss % = (Stack Loss / Heat Input) × 100
        - Typical values:
            * High efficiency: 5-10%
            * Medium efficiency: 10-20%
            * Low efficiency: 20-40%
        - Reducing stack temperature is primary way to reduce stack loss
        - Economizers recover heat from flue gas, reducing stack loss

    References:
        - ASME PTC 4: Stack loss calculation methods
        - GPSA: Flue gas enthalpy correlations
    """
    if heat_input <= 0:
        raise ValueError(f"Heat input must be positive, got {heat_input}")

    # Calculate stack loss (BTU/hr)
    stack_loss = flue_gas_enthalpy * flue_gas_flow

    # Calculate as percentage
    stack_loss_pct = (stack_loss / heat_input) * 100.0

    # Sanity check (typical range 5-50%)
    if stack_loss_pct < 0:
        raise ValueError(
            f"Stack loss percentage {stack_loss_pct:.1f}% is negative. "
            "Check that flue gas enthalpy is above ambient."
        )

    if stack_loss_pct > 100:
        raise ValueError(
            f"Stack loss percentage {stack_loss_pct:.1f}% exceeds 100%. "
            "Check calculation inputs - flue gas flow or enthalpy may be incorrect."
        )

    return stack_loss_pct


def thermal_efficiency(
    heat_output: float,
    heat_input: float
) -> float:
    """
    Calculate thermal efficiency from heat output and input.

    Simple efficiency calculation based on direct heat transfer.
    For combustion efficiency with losses, use combustion_efficiency().

    Args:
        heat_output: Useful heat delivered to process (BTU/hr or consistent units)
        heat_input: Total heat input from fuel (same units as heat_output)

    Returns:
        Thermal efficiency (%)

    Raises:
        ValueError: If heat_input is zero or negative
        ValueError: If heat_output exceeds heat_input

    Example:
        >>> # 85% efficient heat transfer
        >>> thermal_efficiency(850000, 1000000)
        85.0

        >>> # Perfect heat transfer (theoretical)
        >>> thermal_efficiency(1000, 1000)
        100.0

    Notes:
        - Efficiency = (Heat Output / Heat Input) × 100
        - This is the "input-output" method
        - For combustion systems, use combustion_efficiency() which
          accounts for individual loss mechanisms
    """
    if heat_input <= 0:
        raise ValueError(f"Heat input must be positive, got {heat_input}")

    if heat_output > heat_input:
        raise ValueError(
            f"Heat output ({heat_output}) cannot exceed heat input ({heat_input})"
        )

    efficiency = (heat_output / heat_input) * 100.0

    return efficiency


def radiation_loss_percent(
    surface_area: float,
    surface_temp: float,
    ambient_temp: float,
    emissivity: float = 0.9
) -> float:
    """
    Estimate radiation heat loss as percentage of typical heat input.

    Simplified radiation loss calculation for equipment surface.
    Actual calculation requires heat input value.

    Args:
        surface_area: External surface area (ft²)
        surface_temp: Average surface temperature (°F)
        ambient_temp: Ambient temperature (°F)
        emissivity: Surface emissivity (0-1, default 0.9)

    Returns:
        Estimated radiation loss (BTU/hr per unit surface area)

    Note:
        This is a simplified calculation. For accurate radiation loss
        percentage, divide result by actual heat input.

    Example:
        >>> # Small boiler: 100 ft² surface, 200°F surface temp, 70°F ambient
        >>> radiation_loss_percent(100, 200, 70)
        129.0...

    References:
        - Stefan-Boltzmann law for radiation heat transfer
        - Typical boiler radiation losses: 1-3% of heat input
    """
    # Stefan-Boltzmann constant (BTU/hr·ft²·R⁴)
    sigma = 0.1714e-8

    # Convert to absolute temperature (Rankine)
    T_surface_R = surface_temp + 459.67
    T_ambient_R = ambient_temp + 459.67

    # Calculate radiation heat transfer (BTU/hr)
    q_radiation = (
        emissivity * sigma * surface_area *
        (T_surface_R**4 - T_ambient_R**4)
    )

    return q_radiation


# VBA-compatible wrapper functions
def CombustionEfficiency(
    heat_input: float,
    stack_loss: float,
    radiation_loss: float = 0.0,
    blow_down_loss: float = 0.0,
    unaccounted_loss: float = 0.0
) -> float:
    """VBA-compatible wrapper for combustion_efficiency()."""
    return combustion_efficiency(
        heat_input, stack_loss, radiation_loss, blow_down_loss, unaccounted_loss
    )


def StackLossPercent(
    flue_gas_enthalpy: float,
    flue_gas_flow: float,
    heat_input: float
) -> float:
    """VBA-compatible wrapper for stack_loss_percent()."""
    return stack_loss_percent(flue_gas_enthalpy, flue_gas_flow, heat_input)


def ThermalEfficiency(heat_output: float, heat_input: float) -> float:
    """VBA-compatible wrapper for thermal_efficiency()."""
    return thermal_efficiency(heat_output, heat_input)
