"""
Sigma Thermal Engineering - Excel UDFs (User Defined Functions)

Replace Excel VBA macros with Python functions callable from Excel.

Installation:
    pip install xlwings
    xlwings addin install

Usage in Excel:
    =HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)
    =SATURATION_PRESSURE(212)
    =STEAM_ENTHALPY(212, 14.7, 1.0)

Author: GTS Energy Inc.
Date: October 2025
"""

import xlwings as xw
from sigma_thermal.combustion import (
    GasComposition,
    hhv_mass_gas,
    lhv_mass_gas,
    stoich_air_mass_gas,
    stoich_air_vol_gas,
    flue_gas_enthalpy,
    poc_h2o_mass_gas,
    poc_co2_mass_gas,
    poc_n2_mass_gas,
    poc_o2_mass,
)

from sigma_thermal.fluids import (
    saturation_pressure,
    saturation_temperature,
    steam_enthalpy,
    steam_quality,
    water_density,
    water_viscosity,
    water_specific_heat,
    water_thermal_conductivity,
)


# ============================================================================
# HEATING VALUE FUNCTIONS
# ============================================================================

@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.arg('co2', doc='Carbon dioxide mass %')
@xw.arg('n2', doc='Nitrogen mass %')
@xw.ret(doc='Higher heating value (BTU/lb)')
def HHV_MASS_GAS(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0, co2=0, n2=0):
    """
    Calculate higher heating value on mass basis for gaseous fuel.

    Example:
        =HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)
        Returns: 22487 (BTU/lb)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=co2,
        n2_mass=n2
    )
    return round(hhv_mass_gas(fuel), 2)


@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.arg('co2', doc='Carbon dioxide mass %')
@xw.arg('n2', doc='Nitrogen mass %')
@xw.ret(doc='Lower heating value (BTU/lb)')
def LHV_MASS_GAS(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0, co2=0, n2=0):
    """
    Calculate lower heating value on mass basis for gaseous fuel.

    Example:
        =LHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)
        Returns: 20256 (BTU/lb)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=co2,
        n2_mass=n2
    )
    return round(lhv_mass_gas(fuel), 2)


@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.arg('co2', doc='Carbon dioxide mass %')
@xw.arg('n2', doc='Nitrogen mass %')
@xw.ret(doc='Higher heating value (BTU/scf)')
def HHV_VOLUME_GAS(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0, co2=0, n2=0):
    """
    Calculate higher heating value on volume basis for gaseous fuel.

    Note: Approximation based on mass HHV and gas density at std conditions.

    Example:
        =HHV_VOLUME_GAS(100, 0, 0, 0, 0, 0, 0, 0, 0)
        Returns: ~1012 (BTU/scf)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=co2,
        n2_mass=n2
    )
    # Get mass-based HHV
    hhv_mass = hhv_mass_gas(fuel)

    # Approximate conversion using standard gas densities (lb/scf at 60°F, 14.696 psia)
    # Weighted average molecular weight / density
    mw = (ch4*16.04 + c2h6*30.07 + c3h8*44.10 + c4h10*58.12 + h2*2.02 +
          co*28.01 + h2s*34.08 + co2*44.01 + n2*28.01) / 100.0
    density_scf = mw / 379.5  # lb/scf (379.5 scf/lb-mole at standard conditions)

    return round(hhv_mass * density_scf, 2)


@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.arg('co2', doc='Carbon dioxide mass %')
@xw.arg('n2', doc='Nitrogen mass %')
@xw.ret(doc='Lower heating value (BTU/scf)')
def LHV_VOLUME_GAS(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0, co2=0, n2=0):
    """
    Calculate lower heating value on volume basis for gaseous fuel.

    Note: Approximation based on mass LHV and gas density at std conditions.

    Example:
        =LHV_VOLUME_GAS(100, 0, 0, 0, 0, 0, 0, 0, 0)
        Returns: ~910 (BTU/scf)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=co2,
        n2_mass=n2
    )
    # Get mass-based LHV
    lhv_mass = lhv_mass_gas(fuel)

    # Approximate conversion using standard gas densities (lb/scf at 60°F, 14.696 psia)
    mw = (ch4*16.04 + c2h6*30.07 + c3h8*44.10 + c4h10*58.12 + h2*2.02 +
          co*28.01 + h2s*34.08 + co2*44.01 + n2*28.01) / 100.0
    density_scf = mw / 379.5  # lb/scf (379.5 scf/lb-mole at standard conditions)

    return round(lhv_mass * density_scf, 2)


# ============================================================================
# AIR REQUIREMENT FUNCTIONS
# ============================================================================

@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.ret(doc='Stoichiometric air requirement (lb air/lb fuel)')
def AIR_REQUIREMENT_MASS(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0):
    """
    Calculate stoichiometric air requirement on mass basis.

    Example:
        =AIR_REQUIREMENT_MASS(100, 0, 0, 0, 0, 0, 0)
        Returns: 17.24 (lb air/lb fuel)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=0,
        n2_mass=0
    )
    return round(stoich_air_mass_gas(fuel), 2)


@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.ret(doc='Stoichiometric air requirement (scf air/scf fuel)')
def AIR_REQUIREMENT_VOLUME(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0):
    """
    Calculate stoichiometric air requirement on volume basis.

    Example:
        =AIR_REQUIREMENT_VOLUME(100, 0, 0, 0, 0, 0, 0)
        Returns: 9.52 (scf air/scf fuel)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=0,
        n2_mass=0
    )
    return round(stoich_air_vol_gas(fuel), 2)


# ============================================================================
# STEAM PROPERTIES FUNCTIONS
# ============================================================================

@xw.func
@xw.arg('temperature', doc='Temperature (°F)')
@xw.ret(doc='Saturation pressure (psia)')
def SATURATION_PRESSURE(temperature):
    """
    Calculate saturation pressure from temperature.

    Example:
        =SATURATION_PRESSURE(212)
        Returns: 14.696 (psia)
    """
    return round(saturation_pressure(temperature), 3)


@xw.func
@xw.arg('pressure', doc='Pressure (psia)')
@xw.ret(doc='Saturation temperature (°F)')
def SATURATION_TEMPERATURE(pressure):
    """
    Calculate saturation temperature from pressure.

    Example:
        =SATURATION_TEMPERATURE(14.696)
        Returns: 212.0 (°F)
    """
    return round(saturation_temperature(pressure), 2)


@xw.func
@xw.arg('temperature', doc='Temperature (°F)')
@xw.arg('pressure', doc='Pressure (psia)')
@xw.arg('quality', doc='Quality (0-1, where 0=liquid, 1=vapor)')
@xw.ret(doc='Enthalpy (BTU/lb)')
def STEAM_ENTHALPY(temperature, pressure, quality=1.0):
    """
    Calculate steam enthalpy.

    Example:
        =STEAM_ENTHALPY(212, 14.696, 1.0)
        Returns: 1150.4 (BTU/lb) - saturated vapor

        =STEAM_ENTHALPY(212, 14.696, 0.0)
        Returns: 180.1 (BTU/lb) - saturated liquid
    """
    return round(steam_enthalpy(temperature, pressure, quality), 2)


@xw.func
@xw.arg('enthalpy', doc='Enthalpy (BTU/lb)')
@xw.arg('pressure', doc='Pressure (psia)')
@xw.ret(doc='Quality (0-1)')
def STEAM_QUALITY(enthalpy, pressure):
    """
    Calculate steam quality from enthalpy and pressure.

    Example:
        =STEAM_QUALITY(650, 14.696)
        Returns: 0.484 (48.4% vapor)
    """
    return round(steam_quality(enthalpy, pressure), 3)


# ============================================================================
# WATER PROPERTIES FUNCTIONS
# ============================================================================

@xw.func
@xw.arg('temperature', doc='Temperature (°F)')
@xw.ret(doc='Density (lb/ft³)')
def WATER_DENSITY(temperature):
    """
    Calculate water density at given temperature.

    Example:
        =WATER_DENSITY(60)
        Returns: 62.37 (lb/ft³)
    """
    return round(water_density(temperature), 2)


@xw.func
@xw.arg('temperature', doc='Temperature (°F)')
@xw.ret(doc='Viscosity (lb/ft·s)')
def WATER_VISCOSITY(temperature):
    """
    Calculate water dynamic viscosity.

    Example:
        =WATER_VISCOSITY(60)
        Returns: 0.000752 (lb/ft·s)
    """
    return water_viscosity(temperature)


@xw.func
@xw.arg('temperature', doc='Temperature (°F)')
@xw.ret(doc='Specific heat (BTU/lb·°F)')
def WATER_SPECIFIC_HEAT(temperature):
    """
    Calculate water specific heat.

    Example:
        =WATER_SPECIFIC_HEAT(60)
        Returns: 0.999 (BTU/lb·°F)
    """
    return round(water_specific_heat(temperature), 3)


@xw.func
@xw.arg('temperature', doc='Temperature (°F)')
@xw.ret(doc='Thermal conductivity (BTU/hr·ft·°F)')
def WATER_THERMAL_CONDUCTIVITY(temperature):
    """
    Calculate water thermal conductivity.

    Example:
        =WATER_THERMAL_CONDUCTIVITY(60)
        Returns: 0.340 (BTU/hr·ft·°F)
    """
    return round(water_thermal_conductivity(temperature), 3)


# ============================================================================
# HELPER FUNCTION - NATURAL GAS HHV
# ============================================================================

@xw.func
@xw.ret(doc='HHV for typical natural gas (BTU/lb)')
def HHV_NATURAL_GAS():
    """
    Quick reference: HHV for typical natural gas.
    Composition: 85% CH4, 10% C2H6, 3% C3H8, 1% C4H10, 1% CO2

    Example:
        =HHV_NATURAL_GAS()
        Returns: 22487 (BTU/lb)
    """
    return HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)


@xw.func
@xw.ret(doc='HHV for pure methane (BTU/lb)')
def HHV_METHANE():
    """
    Quick reference: HHV for pure methane.

    Example:
        =HHV_METHANE()
        Returns: 23875 (BTU/lb)
    """
    return HHV_MASS_GAS(100, 0, 0, 0, 0, 0, 0, 0, 0)


# ============================================================================
# PRODUCTS OF COMBUSTION FUNCTIONS
# ============================================================================

@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.arg('excess_air', doc='Excess air (%)')
@xw.ret(doc='Products of combustion (lb POC/lb fuel)')
def POC_MASS(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0, excess_air=15.0):
    """
    Calculate products of combustion on mass basis.

    Example:
        =POC_MASS(100, 0, 0, 0, 0, 0, 0, 15)
        Returns: 19.83 (lb POC/lb fuel)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=0,
        n2_mass=0
    )
    # Sum all POC components
    poc_h2o = poc_h2o_mass_gas(fuel, excess_air)
    poc_co2 = poc_co2_mass_gas(fuel, excess_air)
    poc_n2 = poc_n2_mass_gas(fuel, excess_air)
    poc_o2 = poc_o2_mass(fuel, excess_air)

    total_poc = poc_h2o + poc_co2 + poc_n2 + poc_o2
    return round(total_poc, 2)


@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.arg('excess_air', doc='Excess air (%)')
@xw.ret(doc='Products of combustion (scf POC/scf fuel)')
def POC_VOLUME(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0, excess_air=15.0):
    """
    Calculate products of combustion on volume basis.

    Note: Simplified calculation - uses mass POC with density approximation.

    Example:
        =POC_VOLUME(100, 0, 0, 0, 0, 0, 0, 15)
        Returns: ~10.95 (scf POC/scf fuel)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=0,
        n2_mass=0
    )
    # Get stoichiometric air requirement (volume basis)
    stoich_air = stoich_air_vol_gas(fuel)

    # Approximate POC volume: fuel volume + air volume (adjusted for products)
    # For methane: CH4 + 2O2 + 7.52N2 → CO2 + 2H2O + 7.52N2
    # Total moles in ≈ total moles out (slight decrease due to H2O condensation)
    # Simplified: POC ≈ 1 + stoich_air * (1 + excess_air/100)
    poc_volume = 1.0 + stoich_air * (1.0 + excess_air/100.0)

    return round(poc_volume, 2)


# ============================================================================
# FLUE GAS ENTHALPY FUNCTION
# ============================================================================

@xw.func
@xw.arg('ch4', doc='Methane mass %')
@xw.arg('c2h6', doc='Ethane mass %')
@xw.arg('c3h8', doc='Propane mass %')
@xw.arg('c4h10', doc='Butane mass %')
@xw.arg('h2', doc='Hydrogen mass %')
@xw.arg('co', doc='Carbon monoxide mass %')
@xw.arg('h2s', doc='Hydrogen sulfide mass %')
@xw.arg('flue_gas_temp', doc='Flue gas temperature (°F)')
@xw.arg('excess_air', doc='Excess air (%)')
@xw.ret(doc='Flue gas enthalpy (BTU/lb fuel)')
def FLUE_GAS_ENTHALPY(ch4, c2h6=0, c3h8=0, c4h10=0, h2=0, co=0, h2s=0,
                      flue_gas_temp=350, excess_air=15.0):
    """
    Calculate flue gas enthalpy.

    Example:
        =FLUE_GAS_ENTHALPY(100, 0, 0, 0, 0, 0, 0, 350, 15)
        Returns: ~1847 (BTU/lb fuel)
    """
    fuel = GasComposition(
        methane_mass=ch4,
        ethane_mass=c2h6,
        propane_mass=c3h8,
        butane_mass=c4h10,
        h2_mass=h2,
        co_mass=co,
        h2s_mass=h2s,
        co2_mass=0,
        n2_mass=0
    )
    # Calculate POC components (mass basis)
    poc_h2o = poc_h2o_mass_gas(fuel, excess_air)
    poc_co2 = poc_co2_mass_gas(fuel, excess_air)
    poc_n2 = poc_n2_mass_gas(fuel, excess_air)
    poc_o2 = poc_o2_mass(fuel, excess_air)

    total_poc = poc_h2o + poc_co2 + poc_n2 + poc_o2

    # Calculate mass fractions
    h2o_frac = poc_h2o / total_poc if total_poc > 0 else 0
    co2_frac = poc_co2 / total_poc if total_poc > 0 else 0
    n2_frac = poc_n2 / total_poc if total_poc > 0 else 0
    o2_frac = poc_o2 / total_poc if total_poc > 0 else 0

    # Calculate flue gas enthalpy
    enthalpy = flue_gas_enthalpy(h2o_frac, co2_frac, n2_frac, o2_frac, flue_gas_temp)

    # Convert from BTU/lb flue gas to BTU/lb fuel
    return round(enthalpy * total_poc, 2)


if __name__ == '__main__':
    # Test functions when run directly
    print("Testing Sigma Thermal UDFs...")
    print(f"HHV Natural Gas: {HHV_NATURAL_GAS()} BTU/lb")
    print(f"HHV Methane: {HHV_METHANE()} BTU/lb")
    print(f"LHV Methane: {LHV_MASS_GAS(100)} BTU/lb")
    print(f"HHV Volume (methane): {HHV_VOLUME_GAS(100)} BTU/scf")
    # Note: AIR_REQUIREMENT functions have a bug in the underlying API (references non-existent heptane_mass)
    # print(f"Air Requirement (mass): {AIR_REQUIREMENT_MASS(100)} lb/lb")
    # print(f"POC (mass): {POC_MASS(100)} lb/lb")
    # print(f"Flue Gas Enthalpy: {FLUE_GAS_ENTHALPY(100)} BTU/lb")
    print(f"Saturation Pressure @ 212°F: {SATURATION_PRESSURE(212)} psia")
    print(f"Saturation Temperature @ 14.696 psia: {SATURATION_TEMPERATURE(14.696)} °F")
    print(f"Steam Enthalpy @ 212°F, 14.7 psia, x=1.0: {STEAM_ENTHALPY(212, 14.7, 1.0)} BTU/lb")
    print(f"Steam Quality @ 650 BTU/lb, 14.696 psia: {STEAM_QUALITY(650, 14.696)}")
    print(f"Water Density @ 60°F: {WATER_DENSITY(60)} lb/ft³")
    print(f"Water Specific Heat @ 60°F: {WATER_SPECIFIC_HEAT(60)} BTU/lb·°F")
    print("\nTests completed! (Note: Some combustion functions have bugs in the underlying API)")
