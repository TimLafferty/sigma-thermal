"""
Integration Test: Complete Boiler Efficiency Calculation

This integration test validates the complete workflow for calculating
boiler combustion efficiency, testing the integration of:
1. Heating value calculations
2. Stoichiometric air requirements
3. Products of combustion calculations
4. Flue gas enthalpy calculations
5. Stack loss and efficiency calculations

These tests represent realistic industrial scenarios and verify that
all combustion functions work together correctly.

References:
- ASME PTC 4 (Fired Steam Generators)
- GPSA Engineering Data Book, 13th Edition
"""

import pytest
from sigma_thermal.combustion import (
    GasComposition,
    GasCompositionMass,
    hhv_mass_gas,
    lhv_mass_gas,
    hhv_mass_liquid,
    lhv_mass_liquid,
    poc_h2o_mass_gas,
    poc_co2_mass_gas,
    poc_n2_mass_gas,
    poc_h2o_mass_liquid,
    poc_co2_mass_liquid,
    poc_n2_mass_liquid,
    poc_o2_mass,
    flue_gas_enthalpy,
)


class BoilerEfficiencyCalculator:
    """
    Helper class for complete boiler efficiency calculations.

    This demonstrates the full workflow that end users would implement
    to calculate combustion efficiency for gas or liquid fuels.
    """

    def __init__(
        self,
        fuel_type: str,
        fuel_flow: float,
        excess_air_percent: float,
        stack_temp: float,
        ambient_temp: float = 77.0,
        humidity: float = 0.013
    ):
        """
        Initialize boiler calculator.

        Parameters
        ----------
        fuel_type : str
            Either a gas composition dict or liquid fuel type string
        fuel_flow : float
            Fuel mass flow rate (lb/hr)
        excess_air_percent : float
            Excess air percentage (%)
        stack_temp : float
            Stack gas temperature (°F)
        ambient_temp : float
            Ambient air temperature (°F)
        humidity : float
            Humidity ratio (lb H2O / lb dry air)
        """
        self.fuel_type = fuel_type
        self.fuel_flow = fuel_flow
        self.excess_air_percent = excess_air_percent
        self.stack_temp = stack_temp
        self.ambient_temp = ambient_temp
        self.humidity = humidity

        # Results storage
        self.results = {}

    def calculate_gas_fuel(
        self,
        composition: GasComposition,
        composition_mass: GasCompositionMass,
        stoich_air_ratio: float
    ):
        """Calculate efficiency for gaseous fuel."""
        # 1. Heating values
        hhv = hhv_mass_gas(composition)
        lhv = lhv_mass_gas(composition)

        # 2. Air requirements
        stoich_air_flow = self.fuel_flow * stoich_air_ratio
        actual_air_flow = stoich_air_flow * (1 + self.excess_air_percent / 100)

        # 3. Products of combustion
        h2o = poc_h2o_mass_gas(
            composition_mass, self.fuel_flow, self.humidity, actual_air_flow
        )
        co2 = poc_co2_mass_gas(composition_mass, self.fuel_flow)
        n2 = poc_n2_mass_gas(
            composition_mass, self.fuel_flow, actual_air_flow, stoich_air_flow
        )
        o2 = poc_o2_mass(
            self.fuel_flow, actual_air_flow, stoich_air_flow, 0.0
        )

        total_flue = h2o + co2 + n2 + o2

        # 4. Flue gas enthalpy
        flue_enthalpy = flue_gas_enthalpy(
            h2o_fraction=h2o / total_flue,
            co2_fraction=co2 / total_flue,
            n2_fraction=n2 / total_flue,
            o2_fraction=o2 / total_flue,
            gas_temp=self.stack_temp,
            ambient_temp=self.ambient_temp
        )

        # 5. Efficiency calculation
        heat_input = hhv * self.fuel_flow
        stack_loss = flue_enthalpy * total_flue
        combustion_efficiency = (heat_input - stack_loss) / heat_input * 100

        # Store results
        self.results = {
            'fuel_type': 'gas',
            'hhv_btu_lb': hhv,
            'lhv_btu_lb': lhv,
            'heat_input_btu_hr': heat_input,
            'stoich_air_lb_hr': stoich_air_flow,
            'actual_air_lb_hr': actual_air_flow,
            'excess_air_percent': self.excess_air_percent,
            'h2o_lb_hr': h2o,
            'co2_lb_hr': co2,
            'n2_lb_hr': n2,
            'o2_lb_hr': o2,
            'total_flue_lb_hr': total_flue,
            'flue_enthalpy_btu_lb': flue_enthalpy,
            'stack_loss_btu_hr': stack_loss,
            'stack_loss_percent': stack_loss / heat_input * 100,
            'combustion_efficiency_percent': combustion_efficiency,
            'stack_temp_degf': self.stack_temp
        }

        return self.results

    def calculate_liquid_fuel(
        self,
        fuel_name: str,
        stoich_air_ratio: float
    ):
        """Calculate efficiency for liquid fuel."""
        # 1. Heating values
        hhv = hhv_mass_liquid(fuel_name)
        lhv = lhv_mass_liquid(fuel_name)

        # 2. Air requirements
        stoich_air_flow = self.fuel_flow * stoich_air_ratio
        actual_air_flow = stoich_air_flow * (1 + self.excess_air_percent / 100)

        # 3. Products of combustion
        h2o = poc_h2o_mass_liquid(
            fuel_name, self.fuel_flow, self.humidity, actual_air_flow
        )
        co2 = poc_co2_mass_liquid(fuel_name, self.fuel_flow)
        n2 = poc_n2_mass_liquid(
            fuel_name, self.fuel_flow, actual_air_flow, stoich_air_flow
        )
        o2 = poc_o2_mass(
            self.fuel_flow, actual_air_flow, stoich_air_flow, 0.0
        )

        total_flue = h2o + co2 + n2 + o2

        # 4. Flue gas enthalpy
        flue_enthalpy = flue_gas_enthalpy(
            h2o_fraction=h2o / total_flue,
            co2_fraction=co2 / total_flue,
            n2_fraction=n2 / total_flue,
            o2_fraction=o2 / total_flue,
            gas_temp=self.stack_temp,
            ambient_temp=self.ambient_temp
        )

        # 5. Efficiency calculation
        heat_input = hhv * self.fuel_flow
        stack_loss = flue_enthalpy * total_flue
        combustion_efficiency = (heat_input - stack_loss) / heat_input * 100

        # Store results
        self.results = {
            'fuel_type': 'liquid',
            'fuel_name': fuel_name,
            'hhv_btu_lb': hhv,
            'lhv_btu_lb': lhv,
            'heat_input_btu_hr': heat_input,
            'stoich_air_lb_hr': stoich_air_flow,
            'actual_air_lb_hr': actual_air_flow,
            'excess_air_percent': self.excess_air_percent,
            'h2o_lb_hr': h2o,
            'co2_lb_hr': co2,
            'n2_lb_hr': n2,
            'o2_lb_hr': o2,
            'total_flue_lb_hr': total_flue,
            'flue_enthalpy_btu_lb': flue_enthalpy,
            'stack_loss_btu_hr': stack_loss,
            'stack_loss_percent': stack_loss / heat_input * 100,
            'combustion_efficiency_percent': combustion_efficiency,
            'stack_temp_degf': self.stack_temp
        }

        return self.results


class TestBoilerEfficiencyIntegration:
    """Integration tests for complete boiler efficiency calculations."""

    def test_natural_gas_efficient_boiler(self):
        """
        Test complete workflow for efficient natural gas boiler.

        Scenario: Modern condensing boiler with low stack temperature
        """
        # Setup
        calc = BoilerEfficiencyCalculator(
            fuel_type='natural_gas',
            fuel_flow=100.0,  # lb/hr
            excess_air_percent=10.0,
            stack_temp=300.0,  # Very efficient
            ambient_temp=77.0,
            humidity=0.013
        )

        # Pure methane composition
        comp_heating = GasComposition(methane_mass=100.0)
        comp_poc = GasCompositionMass(methane_mass=100.0)
        stoich_air = 17.24  # lb air / lb CH4

        # Execute complete calculation
        results = calc.calculate_gas_fuel(comp_heating, comp_poc, stoich_air)

        # Validate workflow integration
        assert results['fuel_type'] == 'gas'

        # Heating values should be standard for methane
        assert results['hhv_btu_lb'] == pytest.approx(23875, rel=0.001)
        assert results['lhv_btu_lb'] == pytest.approx(21495, rel=0.001)

        # Air flow calculations
        assert results['stoich_air_lb_hr'] == pytest.approx(1724, rel=0.01)
        assert results['actual_air_lb_hr'] == pytest.approx(1896.4, rel=0.01)

        # Products of combustion mass balance
        total_input = (
            calc.fuel_flow +
            results['actual_air_lb_hr'] +
            calc.humidity * results['actual_air_lb_hr']
        )
        total_products = (
            results['h2o_lb_hr'] +
            results['co2_lb_hr'] +
            results['n2_lb_hr'] +
            results['o2_lb_hr']
        )
        assert total_products == pytest.approx(total_input, rel=0.01)

        # Efficiency at 300°F should be very high (>94%)
        assert 94 < results['combustion_efficiency_percent'] < 97
        assert results['stack_loss_percent'] < 6

    def test_natural_gas_inefficient_boiler(self):
        """
        Test complete workflow for older natural gas boiler.

        Scenario: Older atmospheric boiler with high stack temperature
        """
        calc = BoilerEfficiencyCalculator(
            fuel_type='natural_gas',
            fuel_flow=100.0,
            excess_air_percent=25.0,  # Higher excess air
            stack_temp=600.0,  # High stack temp
            ambient_temp=77.0,
            humidity=0.013
        )

        comp_heating = GasComposition(methane_mass=100.0)
        comp_poc = GasCompositionMass(methane_mass=100.0)
        stoich_air = 17.24

        results = calc.calculate_gas_fuel(comp_heating, comp_poc, stoich_air)

        # Efficiency should be lower due to high stack temp and excess air
        assert 84 < results['combustion_efficiency_percent'] < 88
        assert 12 < results['stack_loss_percent'] < 16

        # Higher excess air means more O2 in products
        o2_percent = results['o2_lb_hr'] / results['total_flue_lb_hr'] * 100
        assert 3 < o2_percent < 6

    def test_oil_fired_boiler(self):
        """
        Test complete workflow for #2 oil-fired boiler.

        Scenario: Industrial oil-fired boiler
        """
        calc = BoilerEfficiencyCalculator(
            fuel_type='#2_oil',
            fuel_flow=500.0,  # lb/hr
            excess_air_percent=20.0,
            stack_temp=450.0,
            ambient_temp=77.0,
            humidity=0.013
        )

        results = calc.calculate_liquid_fuel('#2 oil', stoich_air_ratio=14.5)

        # Validate workflow
        assert results['fuel_type'] == 'liquid'
        assert results['fuel_name'] == '#2 oil'

        # Oil heating values
        assert 18900 < results['hhv_btu_lb'] < 19100
        assert results['lhv_btu_lb'] < results['hhv_btu_lb']

        # Efficiency typical for oil at 450°F
        assert 88 < results['combustion_efficiency_percent'] < 92

        # CO2 emissions check
        co2_per_mmbtu = (
            results['co2_lb_hr'] / (results['heat_input_btu_hr'] / 1e6)
        )
        assert 155 < co2_per_mmbtu < 170  # Oil produces more CO2 than gas

    def test_fuel_comparison_same_heat_input(self):
        """
        Compare natural gas vs #2 oil at same heat input.

        This integration test validates that fuel switching
        calculations work correctly.
        """
        # Target heat input: 10 MMBtu/hr
        target_heat_input = 10e6  # BTU/hr

        # Natural gas (methane)
        ng_hhv = 23875  # BTU/lb
        ng_fuel_flow = target_heat_input / ng_hhv

        ng_calc = BoilerEfficiencyCalculator(
            fuel_type='natural_gas',
            fuel_flow=ng_fuel_flow,
            excess_air_percent=15.0,
            stack_temp=400.0
        )

        ng_comp_heating = GasComposition(methane_mass=100.0)
        ng_comp_poc = GasCompositionMass(methane_mass=100.0)
        ng_results = ng_calc.calculate_gas_fuel(
            ng_comp_heating, ng_comp_poc, stoich_air_ratio=17.24
        )

        # #2 Oil
        oil_hhv = 18993  # BTU/lb (approx)
        oil_fuel_flow = target_heat_input / oil_hhv

        oil_calc = BoilerEfficiencyCalculator(
            fuel_type='#2_oil',
            fuel_flow=oil_fuel_flow,
            excess_air_percent=15.0,
            stack_temp=400.0
        )

        oil_results = oil_calc.calculate_liquid_fuel('#2 oil', stoich_air_ratio=14.5)

        # Same heat input
        assert ng_results['heat_input_btu_hr'] == pytest.approx(
            oil_results['heat_input_btu_hr'], rel=0.01
        )

        # Natural gas should be slightly more efficient (less CO2 in flue gas)
        assert ng_results['combustion_efficiency_percent'] > \
               oil_results['combustion_efficiency_percent']

        # Oil requires less air per unit fuel but more fuel for same heat
        assert ng_results['actual_air_lb_hr'] < oil_results['actual_air_lb_hr']

        # Oil produces more CO2 per MMBtu
        ng_co2_per_mmbtu = ng_results['co2_lb_hr'] / (
            ng_results['heat_input_btu_hr'] / 1e6
        )
        oil_co2_per_mmbtu = oil_results['co2_lb_hr'] / (
            oil_results['heat_input_btu_hr'] / 1e6
        )
        assert oil_co2_per_mmbtu > ng_co2_per_mmbtu

    def test_stack_temperature_impact(self):
        """
        Integration test for stack temperature impact on efficiency.

        Tests that all functions integrate properly to show
        expected efficiency changes with stack temperature.
        """
        fuel_flow = 100.0
        excess_air = 10.0
        stack_temps = [300, 400, 500, 600, 700, 800]

        efficiencies = []

        for stack_temp in stack_temps:
            calc = BoilerEfficiencyCalculator(
                fuel_type='natural_gas',
                fuel_flow=fuel_flow,
                excess_air_percent=excess_air,
                stack_temp=stack_temp
            )

            comp_heating = GasComposition(methane_mass=100.0)
            comp_poc = GasCompositionMass(methane_mass=100.0)

            results = calc.calculate_gas_fuel(comp_heating, comp_poc, 17.24)
            efficiencies.append(results['combustion_efficiency_percent'])

        # Efficiency should decrease monotonically with stack temperature
        for i in range(len(efficiencies) - 1):
            assert efficiencies[i] > efficiencies[i+1], \
                f"Efficiency should decrease with stack temp: " \
                f"{efficiencies[i]:.2f}% at {stack_temps[i]}°F vs " \
                f"{efficiencies[i+1]:.2f}% at {stack_temps[i+1]}°F"

        # Range check
        assert 94 < efficiencies[0] < 97  # 300°F - very efficient
        assert 80 < efficiencies[-1] < 85  # 800°F - still reasonable efficiency

    def test_excess_air_impact_integration(self):
        """
        Integration test for excess air impact across all functions.

        Validates that excess air flows through all calculations correctly.
        """
        fuel_flow = 100.0
        stack_temp = 400.0
        excess_air_levels = [5, 10, 15, 20, 30, 50]

        results_list = []

        for excess_air in excess_air_levels:
            calc = BoilerEfficiencyCalculator(
                fuel_type='natural_gas',
                fuel_flow=fuel_flow,
                excess_air_percent=excess_air,
                stack_temp=stack_temp
            )

            comp_heating = GasComposition(methane_mass=100.0)
            comp_poc = GasCompositionMass(methane_mass=100.0)

            results = calc.calculate_gas_fuel(comp_heating, comp_poc, 17.24)
            results_list.append(results)

        # Efficiency should decrease with excess air
        for i in range(len(results_list) - 1):
            assert results_list[i]['combustion_efficiency_percent'] > \
                   results_list[i+1]['combustion_efficiency_percent']

        # O2 in flue gas should increase with excess air
        for i in range(len(results_list) - 1):
            o2_pct_i = results_list[i]['o2_lb_hr'] / results_list[i]['total_flue_lb_hr']
            o2_pct_next = results_list[i+1]['o2_lb_hr'] / results_list[i+1]['total_flue_lb_hr']
            assert o2_pct_i < o2_pct_next

        # Air flow should increase linearly with excess air
        for i, excess_air in enumerate(excess_air_levels):
            expected_air = results_list[0]['stoich_air_lb_hr'] * (1 + excess_air / 100)
            assert results_list[i]['actual_air_lb_hr'] == pytest.approx(
                expected_air, rel=0.01
            )


class TestRealWorldScenarios:
    """Integration tests based on real-world boiler scenarios."""

    def test_process_heater_natural_gas(self):
        """
        Process heater firing natural gas.

        Typical refinery/chemical plant process heater:
        - High firing rate
        - Moderate stack temperature
        - Low excess air (good control)
        """
        calc = BoilerEfficiencyCalculator(
            fuel_type='natural_gas',
            fuel_flow=500.0,  # lb/hr
            excess_air_percent=12.0,
            stack_temp=550.0,  # °F
            ambient_temp=80.0,  # Summer conditions
            humidity=0.015
        )

        # Typical pipeline gas composition
        comp_heating = GasComposition(
            methane_mass=92.0,
            ethane_mass=4.0,
            propane_mass=2.5,
            n2_mass=1.5
        )
        comp_poc = GasCompositionMass(
            methane_mass=92.0,
            ethane_mass=4.0,
            propane_mass=2.5,
            n2_mass=1.5
        )

        # Approximate stoich air for this mixture
        stoich_air = (
            0.92 * 17.24 +  # CH4
            0.04 * 16.12 +  # C2H6
            0.025 * 15.69   # C3H8
        )

        results = calc.calculate_gas_fuel(comp_heating, comp_poc, stoich_air)

        # Validate realistic process heater performance
        assert 87 < results['combustion_efficiency_percent'] < 90
        assert 11.5e6 < results['heat_input_btu_hr'] < 12e6

        # Heat release
        heat_release_mmbtu_hr = results['heat_input_btu_hr'] / 1e6
        assert 11 < heat_release_mmbtu_hr < 12

    def test_package_boiler_oil_fired(self):
        """
        Package boiler firing #2 oil.

        Typical commercial/industrial package boiler:
        - Moderate firing rate
        - Relatively efficient
        - Higher excess air for safety
        """
        calc = BoilerEfficiencyCalculator(
            fuel_type='#2_oil',
            fuel_flow=300.0,  # lb/hr
            excess_air_percent=22.0,  # Conservative for oil
            stack_temp=420.0,
            ambient_temp=75.0,
            humidity=0.012
        )

        results = calc.calculate_liquid_fuel('#2 oil', stoich_air_ratio=14.5)

        # Package boiler performance
        assert 88 < results['combustion_efficiency_percent'] < 91

        # Heat output
        heat_release_mmbtu_hr = results['heat_input_btu_hr'] / 1e6
        assert 5.5 < heat_release_mmbtu_hr < 6.0

        # Stack loss
        assert 9 < results['stack_loss_percent'] < 12
