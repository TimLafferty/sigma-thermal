"""
Combustion calculations module.

This module provides functions for combustion-related calculations including:
- Flue gas enthalpies
- Heating values (HHV, LHV)
- Products of combustion
- Air-fuel ratios
- Flame temperature
- Combustion efficiency
- Emissions calculations

All calculations validated against VBA functions from Engineering-Functions.xlam
"""

from sigma_thermal.combustion.enthalpy import (
    enthalpy_co2,
    enthalpy_h2o,
    enthalpy_n2,
    enthalpy_o2,
    flue_gas_enthalpy,
)
from sigma_thermal.combustion.heating_values import (
    GasComposition,
    hhv_mass_gas,
    lhv_mass_gas,
    hhv_mass_liquid,
    lhv_mass_liquid,
    HHVMass,  # VBA compatibility
    LHVMass,  # VBA compatibility
)
from sigma_thermal.combustion.products import (
    GasCompositionMass,
    GasCompositionVolume,
    poc_h2o_mass_gas,
    poc_h2o_mass_liquid,
    poc_co2_mass_gas,
    poc_co2_mass_liquid,
    poc_n2_mass_gas,
    poc_n2_mass_liquid,
    poc_o2_mass,
    poc_co2_vol_gas,
    poc_h2o_vol_gas,
    poc_n2_vol_gas,
    poc_so2_vol_gas,
    POC_H2OMass,  # VBA compatibility
    POC_CO2Mass,  # VBA compatibility
    POC_N2Mass,   # VBA compatibility
    POC_O2Mass,   # VBA compatibility
    POC_CO2Vol,   # VBA compatibility
    POC_H2OVol,   # VBA compatibility
    POC_N2Vol,    # VBA compatibility
    POC_SO2Vol,   # VBA compatibility
)
from sigma_thermal.combustion.air_fuel import (
    stoich_air_mass_gas,
    stoich_air_vol_gas,
    stoich_air_mass_liquid,
    excess_air_percent,
    StoichAirMassGas,  # VBA compatibility
    StoichAirMassLiquid,  # VBA compatibility
    ExcessAirPercent,  # VBA compatibility
)
from sigma_thermal.combustion.efficiency import (
    combustion_efficiency,
    stack_loss_percent,
    thermal_efficiency,
    radiation_loss_percent,
    CombustionEfficiency,  # VBA compatibility
    StackLossPercent,  # VBA compatibility
    ThermalEfficiency,  # VBA compatibility
)
from sigma_thermal.combustion.flame_temperature import (
    adiabatic_flame_temp,
    flame_temp_excess_air,
    AdiabaticFlameTemp,  # VBA compatibility
    FlameTempExcessAir,  # VBA compatibility
)
from sigma_thermal.combustion.emissions import (
    nox_emissions,
    co2_emissions,
    NOxEmissions,  # VBA compatibility
    CO2Emissions,  # VBA compatibility
)

__all__ = [
    # Enthalpy
    "enthalpy_co2",
    "enthalpy_h2o",
    "enthalpy_n2",
    "enthalpy_o2",
    "flue_gas_enthalpy",
    # Heating values
    "GasComposition",
    "hhv_mass_gas",
    "lhv_mass_gas",
    "hhv_mass_liquid",
    "lhv_mass_liquid",
    "HHVMass",
    "LHVMass",
    # Products of combustion
    "GasCompositionMass",
    "GasCompositionVolume",
    "poc_h2o_mass_gas",
    "poc_h2o_mass_liquid",
    "poc_co2_mass_gas",
    "poc_co2_mass_liquid",
    "poc_n2_mass_gas",
    "poc_n2_mass_liquid",
    "poc_o2_mass",
    "poc_co2_vol_gas",
    "poc_h2o_vol_gas",
    "poc_n2_vol_gas",
    "poc_so2_vol_gas",
    "POC_H2OMass",
    "POC_CO2Mass",
    "POC_N2Mass",
    "POC_O2Mass",
    "POC_CO2Vol",
    "POC_H2OVol",
    "POC_N2Vol",
    "POC_SO2Vol",
    # Air-fuel ratios
    "stoich_air_mass_gas",
    "stoich_air_vol_gas",
    "stoich_air_mass_liquid",
    "excess_air_percent",
    "StoichAirMassGas",
    "StoichAirMassLiquid",
    "ExcessAirPercent",
    # Efficiency
    "combustion_efficiency",
    "stack_loss_percent",
    "thermal_efficiency",
    "radiation_loss_percent",
    "CombustionEfficiency",
    "StackLossPercent",
    "ThermalEfficiency",
    # Flame temperature
    "adiabatic_flame_temp",
    "flame_temp_excess_air",
    "AdiabaticFlameTemp",
    "FlameTempExcessAir",
    # Emissions
    "nox_emissions",
    "co2_emissions",
    "NOxEmissions",
    "CO2Emissions",
]
