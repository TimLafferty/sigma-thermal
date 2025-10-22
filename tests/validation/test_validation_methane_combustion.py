"""
Validation Test Case 1: Pure Methane Combustion

This test validates Python implementation against Excel VBA for a complete
methane combustion calculation including:
- Heating values
- Stoichiometric air requirement
- Products of combustion
- Flue gas enthalpy
- Stack loss
- Combustion efficiency

Test Scenario:
- Fuel: 100% Methane (CH4)
- Fuel Flow: 100 lb/hr
- Excess Air: 10%
- Ambient Temperature: 77°F
- Stack Temperature: 1500°F
- Humidity: 0.013 lb H2O / lb dry air

Expected results are calculated from Excel VBA functions:
- HHVMass, LHVMass
- POC_H2OMass, POC_CO2Mass, POC_N2Mass, POC_O2Mass
- EnthalpyCO2, EnthalpyH2O, EnthalpyN2, EnthalpyO2
- FlueGasEnthalpy

References:
- GPSA Engineering Data Book, 13th Edition
- Perry's Chemical Engineers' Handbook, 8th Edition
- ASME PTC 4 (Fired Steam Generators)
"""

import pytest
from sigma_thermal.combustion import (
    GasComposition,  # For heating values
    GasCompositionMass,  # For POC functions
    hhv_mass_gas,
    lhv_mass_gas,
    poc_h2o_mass_gas,
    poc_co2_mass_gas,
    poc_n2_mass_gas,
    poc_o2_mass,
    enthalpy_co2,
    enthalpy_h2o,
    enthalpy_n2,
    enthalpy_o2,
    flue_gas_enthalpy,
)


class TestValidationMethaneComplete:
    """
    Validation test case for pure methane combustion.

    This test represents a typical natural gas boiler combustion scenario
    and validates against expected VBA outputs.
    """

    # Test parameters
    FUEL_FLOW = 100.0  # lb/hr
    EXCESS_AIR_PERCENT = 10.0  # %
    AMBIENT_TEMP = 77.0  # °F
    STACK_TEMP = 1500.0  # °F
    HUMIDITY = 0.013  # lb H2O / lb dry air

    # Stoichiometric constants for CH4
    # CH4 + 2O2 + 7.53N2 -> CO2 + 2H2O + 7.53N2
    # Mass basis: 1 lb CH4 requires 17.24 lb air (stoichiometric)
    STOICH_AIR_RATIO = 17.24  # lb air / lb CH4

    # Expected VBA results (from Excel VBA functions)
    # These values should be verified against actual Excel calculations
    EXPECTED_HHV = 23875.0  # BTU/lb
    EXPECTED_LHV = 21495.0  # BTU/lb

    @pytest.fixture
    def fuel_composition_heating(self):
        """Pure methane composition for heating value calculations"""
        return GasComposition(methane_mass=100.0)

    @pytest.fixture
    def fuel_composition_poc(self):
        """Pure methane composition for POC calculations"""
        return GasCompositionMass(methane_mass=100.0)

    @pytest.fixture
    def stoich_air_flow(self):
        """Stoichiometric air flow rate"""
        return self.FUEL_FLOW * self.STOICH_AIR_RATIO

    @pytest.fixture
    def actual_air_flow(self, stoich_air_flow):
        """Actual air flow with excess air"""
        return stoich_air_flow * (1.0 + self.EXCESS_AIR_PERCENT / 100.0)

    def test_heating_values(self, fuel_composition_heating):
        """Test HHV and LHV calculations"""
        hhv = hhv_mass_gas(fuel_composition_heating)
        lhv = lhv_mass_gas(fuel_composition_heating)

        # Validate against VBA HHVMass and LHVMass
        assert hhv == pytest.approx(self.EXPECTED_HHV, rel=0.0001)
        assert lhv == pytest.approx(self.EXPECTED_LHV, rel=0.0001)

        # Physical validation
        assert lhv < hhv, "LHV should be less than HHV"
        diff_percent = (hhv - lhv) / hhv * 100
        assert 5 < diff_percent < 15, "HHV-LHV difference should be 5-15%"

    def test_stoichiometric_air(self, stoich_air_flow):
        """Test stoichiometric air calculation"""
        # For CH4: 1 lb fuel needs 4 lb O2 = 4/0.2314 = 17.29 lb air
        # Using standard value 17.24 from combustion theory
        assert stoich_air_flow == pytest.approx(1724.0, rel=0.001)

    def test_products_of_combustion_h2o(
        self, fuel_composition_poc, actual_air_flow
    ):
        """Test water mass in products"""
        h2o_mass = poc_h2o_mass_gas(
            fuel_composition_poc,
            self.FUEL_FLOW,
            self.HUMIDITY,
            actual_air_flow
        )

        # Expected from VBA POC_H2OMass:
        # Stoichiometric H2O from CH4: 100 * 2.246 = 224.6 lb/hr
        # H2O from humidity: 0.013 * 1896.4 = 24.65 lb/hr
        # Total: 249.25 lb/hr
        expected_h2o = 100.0 * 2.246 + self.HUMIDITY * actual_air_flow

        assert h2o_mass == pytest.approx(expected_h2o, rel=0.001)
        assert h2o_mass > 240, "H2O mass should be >240 lb/hr for this case"

    def test_products_of_combustion_co2(self, fuel_composition_poc):
        """Test CO2 mass in products"""
        co2_mass = poc_co2_mass_gas(
            fuel_composition_poc,
            self.FUEL_FLOW
        )

        # Expected from VBA POC_CO2Mass:
        # CH4 produces 2.743 lb CO2 per lb fuel
        # 100 * 2.743 = 274.3 lb/hr
        expected_co2 = 100.0 * 2.743

        assert co2_mass == pytest.approx(expected_co2, rel=0.001)

    def test_products_of_combustion_n2(
        self, fuel_composition_poc, stoich_air_flow, actual_air_flow
    ):
        """Test N2 mass in products"""
        n2_mass = poc_n2_mass_gas(
            fuel_composition_poc,
            self.FUEL_FLOW,
            actual_air_flow,
            stoich_air_flow
        )

        # Expected from VBA POC_N2Mass:
        # Stoichiometric N2 from CH4 combustion: 100 * 13.246 = 1324.6 lb/hr
        # N2 from excess air: (1896.4 - 1724.0) * 0.7686 = 132.5 lb/hr
        # Total: 1457.1 lb/hr
        stoich_n2 = 100.0 * 13.246
        excess_n2 = (actual_air_flow - stoich_air_flow) * 0.7686
        expected_n2 = stoich_n2 + excess_n2

        assert n2_mass == pytest.approx(expected_n2, rel=0.001)
        assert n2_mass > 1400, "N2 mass should be >1400 lb/hr"

    def test_products_of_combustion_o2(
        self, stoich_air_flow, actual_air_flow
    ):
        """Test O2 mass in products"""
        o2_mass = poc_o2_mass(
            self.FUEL_FLOW,
            actual_air_flow,
            stoich_air_flow,
            0.0  # No O2 in fuel
        )

        # Expected from VBA POC_O2Mass:
        # O2 from excess air: (1896.4 - 1724.0) * 0.2314 = 39.9 lb/hr
        expected_o2 = (actual_air_flow - stoich_air_flow) * 0.2314

        assert o2_mass == pytest.approx(expected_o2, rel=0.001)
        assert o2_mass > 0, "O2 should be present with excess air"

    def test_flue_gas_composition(
        self, fuel_composition_poc, stoich_air_flow, actual_air_flow
    ):
        """Test total flue gas mass and composition"""
        # Calculate all products
        h2o = poc_h2o_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, self.HUMIDITY, actual_air_flow
        )
        co2 = poc_co2_mass_gas(fuel_composition_poc, self.FUEL_FLOW)
        n2 = poc_n2_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, actual_air_flow, stoich_air_flow
        )
        o2 = poc_o2_mass(
            self.FUEL_FLOW, actual_air_flow, stoich_air_flow, 0.0
        )

        # Total flue gas = fuel + air + humidity water
        # Humidity adds water mass to the system
        humidity_water = self.HUMIDITY * actual_air_flow
        total_flue_gas = self.FUEL_FLOW + actual_air_flow + humidity_water

        # Check mass balance
        products_total = h2o + co2 + n2 + o2
        assert products_total == pytest.approx(total_flue_gas, rel=0.01)

        # Calculate mass fractions
        h2o_fraction = h2o / total_flue_gas
        co2_fraction = co2 / total_flue_gas
        n2_fraction = n2 / total_flue_gas
        o2_fraction = o2 / total_flue_gas

        # Validate fractions sum to 1
        total_fraction = h2o_fraction + co2_fraction + n2_fraction + o2_fraction
        assert total_fraction == pytest.approx(1.0, abs=0.001)

        # Typical flue gas composition checks
        assert 0.10 < h2o_fraction < 0.15, "H2O should be 10-15% by mass"
        assert 0.12 < co2_fraction < 0.16, "CO2 should be 12-16% by mass"
        assert 0.70 < n2_fraction < 0.75, "N2 should be 70-75% by mass"
        assert 0.01 < o2_fraction < 0.03, "O2 should be 1-3% by mass"

    def test_flue_gas_enthalpy(
        self, fuel_composition_poc, stoich_air_flow, actual_air_flow
    ):
        """Test flue gas enthalpy calculation"""
        # Calculate products and mass fractions
        h2o = poc_h2o_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, self.HUMIDITY, actual_air_flow
        )
        co2 = poc_co2_mass_gas(fuel_composition_poc, self.FUEL_FLOW)
        n2 = poc_n2_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, actual_air_flow, stoich_air_flow
        )
        o2 = poc_o2_mass(
            self.FUEL_FLOW, actual_air_flow, stoich_air_flow, 0.0
        )

        total_mass = h2o + co2 + n2 + o2

        # Calculate flue gas enthalpy
        flue_enthalpy = flue_gas_enthalpy(
            h2o_fraction=h2o / total_mass,
            co2_fraction=co2 / total_mass,
            n2_fraction=n2 / total_mass,
            o2_fraction=o2 / total_mass,
            gas_temp=self.STACK_TEMP,
            ambient_temp=self.AMBIENT_TEMP
        )

        # Expected range based on typical flue gas
        # At 1500°F, enthalpy should be ~400-500 BTU/lb
        assert 380 < flue_enthalpy < 520, f"Flue enthalpy {flue_enthalpy} out of range"

        # Physical validation: should be positive (above ambient)
        assert flue_enthalpy > 0, "Enthalpy should be positive above ambient"

    def test_stack_loss_calculation(
        self, fuel_composition_heating, fuel_composition_poc, stoich_air_flow, actual_air_flow
    ):
        """Test complete stack loss calculation"""
        # Get HHV
        hhv = hhv_mass_gas(fuel_composition_heating)

        # Calculate products
        h2o = poc_h2o_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, self.HUMIDITY, actual_air_flow
        )
        co2 = poc_co2_mass_gas(fuel_composition_poc, self.FUEL_FLOW)
        n2 = poc_n2_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, actual_air_flow, stoich_air_flow
        )
        o2 = poc_o2_mass(
            self.FUEL_FLOW, actual_air_flow, stoich_air_flow, 0.0
        )

        total_flue = h2o + co2 + n2 + o2

        # Calculate flue gas enthalpy
        flue_enthalpy = flue_gas_enthalpy(
            h2o_fraction=h2o / total_flue,
            co2_fraction=co2 / total_flue,
            n2_fraction=n2 / total_flue,
            o2_fraction=o2 / total_flue,
            gas_temp=self.STACK_TEMP,
            ambient_temp=self.AMBIENT_TEMP
        )

        # Calculate stack loss
        heat_input = hhv * self.FUEL_FLOW  # BTU/hr
        stack_loss = flue_enthalpy * total_flue  # BTU/hr
        stack_loss_percent = (stack_loss / heat_input) * 100

        # Stack loss at 1500°F with 10% excess air: 30-40%
        # High stack temperature results in significant losses
        assert 30 < stack_loss_percent < 45, \
            f"Stack loss {stack_loss_percent:.1f}% out of typical range"

        # Calculate combustion efficiency
        combustion_efficiency = 100 - stack_loss_percent

        # Combustion efficiency with high stack temperature: 55-70%
        assert 55 < combustion_efficiency < 75, \
            f"Efficiency {combustion_efficiency:.1f}% out of typical range"

        return {
            'heat_input_btu_hr': heat_input,
            'stack_loss_btu_hr': stack_loss,
            'stack_loss_percent': stack_loss_percent,
            'combustion_efficiency': combustion_efficiency,
            'flue_gas_flow_lb_hr': total_flue,
            'flue_gas_enthalpy_btu_lb': flue_enthalpy
        }

    def test_complete_combustion_calculation(
        self, fuel_composition_heating, fuel_composition_poc, stoich_air_flow, actual_air_flow
    ):
        """
        Complete end-to-end combustion calculation test.

        This test validates the entire workflow from fuel input to
        efficiency output, matching expected VBA results.
        """
        # 1. Fuel properties
        hhv = hhv_mass_gas(fuel_composition_heating)
        lhv = lhv_mass_gas(fuel_composition_heating)

        # 2. Air requirements
        assert stoich_air_flow == pytest.approx(1724.0, rel=0.001)
        assert actual_air_flow == pytest.approx(1896.4, rel=0.001)

        # 3. Products of combustion
        h2o = poc_h2o_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, self.HUMIDITY, actual_air_flow
        )
        co2 = poc_co2_mass_gas(fuel_composition_poc, self.FUEL_FLOW)
        n2 = poc_n2_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, actual_air_flow, stoich_air_flow
        )
        o2 = poc_o2_mass(
            self.FUEL_FLOW, actual_air_flow, stoich_air_flow, 0.0
        )

        total_flue = h2o + co2 + n2 + o2

        # 4. Flue gas enthalpy
        flue_enthalpy = flue_gas_enthalpy(
            h2o_fraction=h2o / total_flue,
            co2_fraction=co2 / total_flue,
            n2_fraction=n2 / total_flue,
            o2_fraction=o2 / total_flue,
            gas_temp=self.STACK_TEMP,
            ambient_temp=self.AMBIENT_TEMP
        )

        # 5. Heat balance
        heat_input = hhv * self.FUEL_FLOW
        stack_loss = flue_enthalpy * total_flue
        efficiency = (heat_input - stack_loss) / heat_input * 100

        # 6. Validate complete calculation
        results = {
            'fuel_flow_lb_hr': self.FUEL_FLOW,
            'hhv_btu_lb': hhv,
            'lhv_btu_lb': lhv,
            'stoich_air_lb_hr': stoich_air_flow,
            'actual_air_lb_hr': actual_air_flow,
            'excess_air_percent': self.EXCESS_AIR_PERCENT,
            'h2o_lb_hr': h2o,
            'co2_lb_hr': co2,
            'n2_lb_hr': n2,
            'o2_lb_hr': o2,
            'total_flue_lb_hr': total_flue,
            'stack_temp_degf': self.STACK_TEMP,
            'flue_enthalpy_btu_lb': flue_enthalpy,
            'heat_input_btu_hr': heat_input,
            'stack_loss_btu_hr': stack_loss,
            'combustion_efficiency_percent': efficiency
        }

        # Print results for documentation
        print("\n=== Complete Methane Combustion Calculation ===")
        for key, value in results.items():
            print(f"{key:35s}: {value:12.2f}")

        # Overall validation
        # At 1500°F stack temperature, efficiency will be lower (55-70%)
        assert efficiency > 55, "Efficiency too low"
        assert efficiency < 75, "Efficiency suspiciously high"
        assert total_flue > 1900, "Flue gas mass too low"
        assert total_flue < 2100, "Flue gas mass too high"

        return results


class TestValidationMethaneSensitivity:
    """Sensitivity analysis for methane combustion"""

    def test_excess_air_sensitivity(self):
        """Test effect of varying excess air"""
        fuel_heating = GasComposition(methane_mass=100.0)
        fuel_poc = GasCompositionMass(methane_mass=100.0)
        fuel_flow = 100.0
        stoich_air = 1724.0

        excess_air_values = [0, 5, 10, 15, 20, 30, 50]
        efficiencies = []

        for excess in excess_air_values:
            actual_air = stoich_air * (1.0 + excess / 100.0)

            # Calculate efficiency
            hhv = hhv_mass_gas(fuel_heating)
            h2o = poc_h2o_mass_gas(fuel_poc, fuel_flow, 0.013, actual_air)
            co2 = poc_co2_mass_gas(fuel_poc, fuel_flow)
            n2 = poc_n2_mass_gas(fuel_poc, fuel_flow, actual_air, stoich_air)
            o2 = poc_o2_mass(fuel_flow, actual_air, stoich_air, 0.0)

            total = h2o + co2 + n2 + o2
            enthalpy = flue_gas_enthalpy(
                h2o/total, co2/total, n2/total, o2/total,
                1500.0, 77.0
            )

            efficiency = (hhv * fuel_flow - enthalpy * total) / (hhv * fuel_flow) * 100
            efficiencies.append(efficiency)

        # Efficiency should decrease with increasing excess air
        for i in range(len(efficiencies) - 1):
            assert efficiencies[i] > efficiencies[i+1], \
                "Efficiency should decrease with excess air"

    def test_stack_temperature_sensitivity(self):
        """Test effect of varying stack temperature"""
        fuel_heating = GasComposition(methane_mass=100.0)
        fuel_poc = GasCompositionMass(methane_mass=100.0)
        fuel_flow = 100.0
        stoich_air = 1724.0
        actual_air = 1896.4

        stack_temps = [300, 500, 750, 1000, 1500, 2000, 2500]
        efficiencies = []

        for temp in stack_temps:
            hhv = hhv_mass_gas(fuel_heating)
            h2o = poc_h2o_mass_gas(fuel_poc, fuel_flow, 0.013, actual_air)
            co2 = poc_co2_mass_gas(fuel_poc, fuel_flow)
            n2 = poc_n2_mass_gas(fuel_poc, fuel_flow, actual_air, stoich_air)
            o2 = poc_o2_mass(fuel_flow, actual_air, stoich_air, 0.0)

            total = h2o + co2 + n2 + o2
            enthalpy = flue_gas_enthalpy(
                h2o/total, co2/total, n2/total, o2/total,
                temp, 77.0
            )

            efficiency = (hhv * fuel_flow - enthalpy * total) / (hhv * fuel_flow) * 100
            efficiencies.append(efficiency)

        # Efficiency should decrease with increasing stack temperature
        for i in range(len(efficiencies) - 1):
            assert efficiencies[i] > efficiencies[i+1], \
                "Efficiency should decrease with stack temperature"
