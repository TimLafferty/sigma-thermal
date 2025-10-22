"""
Fluids module for thermophysical property calculations.

This module provides functions for calculating properties of water, steam,
and other fluids used in thermal engineering applications.

Submodules:
    - water_properties: Water and steam saturation, density, enthalpy, etc.
    - psychrometric: Humid air properties (future)
    - thermal_fluids: Thermal oil and heat transfer fluid properties (future)

All calculations validated against industry-standard references including
ASME Steam Tables, IAPWS-IF97, and engineering handbooks.
"""

from sigma_thermal.fluids.water_properties import (
    saturation_pressure,
    saturation_temperature,
    water_density,
    water_viscosity,
    water_specific_heat,
    water_thermal_conductivity,
    steam_enthalpy,
    steam_quality,
    SaturationPressure,  # VBA compatibility
    SaturationTemperature,  # VBA compatibility
    WaterDensity,  # VBA compatibility
    WaterViscosity,  # VBA compatibility
    WaterSpecificHeat,  # VBA compatibility
    WaterThermalConductivity,  # VBA compatibility
    SteamEnthalpy,  # VBA compatibility
    SteamQuality,  # VBA compatibility
)

__all__ = [
    # Water saturation properties
    "saturation_pressure",
    "saturation_temperature",
    "SaturationPressure",
    "SaturationTemperature",
    # Water transport properties
    "water_density",
    "water_viscosity",
    "WaterDensity",
    "WaterViscosity",
    # Water thermal properties
    "water_specific_heat",
    "water_thermal_conductivity",
    "WaterSpecificHeat",
    "WaterThermalConductivity",
    # Steam properties
    "steam_enthalpy",
    "steam_quality",
    "SteamEnthalpy",
    "SteamQuality",
]
