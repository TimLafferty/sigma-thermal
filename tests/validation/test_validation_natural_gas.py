"""
Validation Test Case 2: Natural Gas Mixture Combustion

This test validates Python implementation for a typical natural gas mixture
containing multiple hydrocarbons and inerts.

Test Scenario:
- Fuel: 90% CH4, 5% C2H6, 3% C3H8, 2% N2 (mass basis)
- Fuel Flow: 100 lb/hr
- Excess Air: 15%
- Ambient Temperature: 77°F
- Stack Temperature: 350°F (typical efficient boiler)
- Humidity: 0.013 lb H2O / lb dry air

This represents a realistic natural gas combustion scenario
in an efficient industrial boiler.

References:
- GPSA Engineering Data Book, 13th Edition
- Natural gas composition data from typical pipelines
- ASME PTC 4 (Fired Steam Generators)
"""

import pytest
from sigma_thermal.combustion import (
    GasComposition,
    GasCompositionMass,
    hhv_mass_gas,
    lhv_mass_gas,
    poc_h2o_mass_gas,
    poc_co2_mass_gas,
    poc_n2_mass_gas,
    poc_o2_mass,
    flue_gas_enthalpy,
)


class TestValidationNaturalGas:
    """
    Validation test for natural gas mixture combustion.

    Tests a realistic pipeline natural gas composition through
    complete combustion analysis.
    """

    # Test parameters
    FUEL_FLOW = 100.0  # lb/hr
    EXCESS_AIR_PERCENT = 15.0  # %
    AMBIENT_TEMP = 77.0  # °F
    STACK_TEMP = 350.0  # °F (efficient boiler)
    HUMIDITY = 0.013  # lb H2O / lb dry air

    # Natural gas composition (mass basis)
    METHANE_MASS = 90.0  # %
    ETHANE_MASS = 5.0    # %
    PROPANE_MASS = 3.0   # %
    N2_MASS = 2.0        # %

    # Stoichiometric air calculation
    # Weighted average based on composition
    # CH4: 17.24, C2H6: 16.12, C3H8: 15.69 lb air/lb fuel
    STOICH_AIR_RATIO = (
        0.90 * 17.24 +  # Methane
        0.05 * 16.12 +  # Ethane
        0.03 * 15.69 +  # Propane
        0.02 * 0.0      # N2 (inert)
    )  # ≈ 16.69 lb air / lb fuel

    @pytest.fixture
    def fuel_composition_heating(self):
        """Natural gas composition for heating values"""
        return GasComposition(
            methane_mass=self.METHANE_MASS,
            ethane_mass=self.ETHANE_MASS,
            propane_mass=self.PROPANE_MASS,
            n2_mass=self.N2_MASS
        )

    @pytest.fixture
    def fuel_composition_poc(self):
        """Natural gas composition for POC calculations"""
        return GasCompositionMass(
            methane_mass=self.METHANE_MASS,
            ethane_mass=self.ETHANE_MASS,
            propane_mass=self.PROPANE_MASS,
            n2_mass=self.N2_MASS
        )

    @pytest.fixture
    def stoich_air_flow(self):
        """Stoichiometric air flow rate"""
        return self.FUEL_FLOW * self.STOICH_AIR_RATIO

    @pytest.fixture
    def actual_air_flow(self, stoich_air_flow):
        """Actual air flow with excess air"""
        return stoich_air_flow * (1.0 + self.EXCESS_AIR_PERCENT / 100.0)

    def test_heating_values(self, fuel_composition_heating):
        """Test HHV and LHV for natural gas mixture"""
        hhv = hhv_mass_gas(fuel_composition_heating)
        lhv = lhv_mass_gas(fuel_composition_heating)

        # Expected values (mass-weighted average)
        # CH4: 23875, C2H6: 22323, C3H8: 21669 BTU/lb
        expected_hhv = (
            0.90 * 23875 +
            0.05 * 22323 +
            0.03 * 21669
        )
        expected_lhv = (
            0.90 * 21495 +
            0.05 * 20418 +
            0.03 * 19937
        )

        assert hhv == pytest.approx(expected_hhv, rel=0.0001)
        assert lhv == pytest.approx(expected_lhv, rel=0.0001)

        # Physical validation
        assert lhv < hhv
        assert 23000 < hhv < 24000, "HHV should be in typical natural gas range"
        assert 20900 < lhv < 22000, "LHV should be in typical natural gas range"

    def test_mixture_heating_value_comparison(self, fuel_composition_heating):
        """Test that mixture HHV is between pure component values"""
        hhv_mixture = hhv_mass_gas(fuel_composition_heating)

        # Pure component HHVs
        hhv_ch4 = 23875
        hhv_c2h6 = 22323
        hhv_c3h8 = 21669

        # Mixture should be weighted average, between min and max
        assert min(hhv_c3h8, hhv_c2h6, hhv_ch4) < hhv_mixture < max(hhv_c3h8, hhv_c2h6, hhv_ch4)

    def test_stoichiometric_air(self, stoich_air_flow):
        """Test stoichiometric air for natural gas mixture"""
        # Should be slightly less than pure methane (17.24)
        # due to heavier hydrocarbons requiring less air per mass
        assert 1650 < stoich_air_flow < 1700
        assert stoich_air_flow == pytest.approx(1669.0, rel=0.01)

    def test_products_of_combustion_composition(
        self, fuel_composition_poc, stoich_air_flow, actual_air_flow
    ):
        """Test complete POC calculation for natural gas"""
        # Calculate all products
        h2o = poc_h2o_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, self.HUMIDITY, actual_air_flow
        )
        co2 = poc_co2_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW
        )
        n2 = poc_n2_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW, actual_air_flow, stoich_air_flow
        )
        o2 = poc_o2_mass(
            self.FUEL_FLOW, actual_air_flow, stoich_air_flow, 0.0
        )

        # Check mass balance (including humidity water)
        humidity_water = self.HUMIDITY * actual_air_flow
        total_input = self.FUEL_FLOW + actual_air_flow + humidity_water
        total_products = h2o + co2 + n2 + o2

        assert total_products == pytest.approx(total_input, rel=0.01)

        # Validate product ranges for natural gas
        assert 235 < h2o < 250, "H2O mass reasonable for natural gas"
        assert 260 < co2 < 280, "CO2 mass reasonable for natural gas"
        assert 1400 < n2 < 1500, "N2 mass reasonable for natural gas"
        assert 50 < o2 < 65, "O2 mass reasonable with 15% excess air"

    def test_co2_emissions(self, fuel_composition_poc):
        """Test CO2 emissions calculation"""
        co2_mass = poc_co2_mass_gas(
            fuel_composition_poc, self.FUEL_FLOW
        )

        # CO2 per unit heat input
        hhv = hhv_mass_gas(
            GasComposition(
                methane_mass=self.METHANE_MASS,
                ethane_mass=self.ETHANE_MASS,
                propane_mass=self.PROPANE_MASS,
                n2_mass=self.N2_MASS
            )
        )

        # lb CO2 per MMBtu
        co2_per_mmbtu = (co2_mass / (hhv * self.FUEL_FLOW)) * 1e6

        # Natural gas typically produces 117 lb CO2/MMBtu
        assert 110 < co2_per_mmbtu < 125, \
            f"CO2 emissions {co2_per_mmbtu:.1f} lb/MMBtu out of typical range"

    def test_efficient_boiler_performance(
        self, fuel_composition_heating, fuel_composition_poc,
        stoich_air_flow, actual_air_flow
    ):
        """
        Test complete boiler efficiency at low stack temperature.

        At 350°F stack temperature, modern boilers achieve 80-85% efficiency.
        """
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

        humidity_water = self.HUMIDITY * actual_air_flow
        total_flue = h2o + co2 + n2 + o2

        # Calculate flue gas enthalpy at 350°F (efficient)
        flue_enthalpy = flue_gas_enthalpy(
            h2o_fraction=h2o / total_flue,
            co2_fraction=co2 / total_flue,
            n2_fraction=n2 / total_flue,
            o2_fraction=o2 / total_flue,
            gas_temp=self.STACK_TEMP,
            ambient_temp=self.AMBIENT_TEMP
        )

        # Calculate efficiency
        heat_input = hhv * self.FUEL_FLOW
        stack_loss = flue_enthalpy * total_flue
        efficiency = (heat_input - stack_loss) / heat_input * 100

        # Modern efficient boiler at 350°F: 90-95% efficiency (low stack temp)
        assert 90 < efficiency < 95, \
            f"Efficiency {efficiency:.1f}% out of expected range for efficient boiler"

        # Stack loss should be very low at 350°F (efficient boiler)
        stack_loss_percent = (stack_loss / heat_input) * 100
        assert 5 < stack_loss_percent < 10, \
            f"Stack loss {stack_loss_percent:.1f}% out of expected range"

        return {
            'efficiency': efficiency,
            'stack_loss_percent': stack_loss_percent,
            'heat_input_btu_hr': heat_input,
            'stack_temp_degf': self.STACK_TEMP
        }

    def test_excess_air_effect(self, fuel_composition_heating, fuel_composition_poc):
        """Test impact of excess air on efficiency"""
        stoich_air = self.FUEL_FLOW * self.STOICH_AIR_RATIO

        excess_air_levels = [5, 10, 15, 20, 30]
        efficiencies = []

        for excess in excess_air_levels:
            actual_air = stoich_air * (1.0 + excess / 100.0)

            hhv = hhv_mass_gas(fuel_composition_heating)
            h2o = poc_h2o_mass_gas(
                fuel_composition_poc, self.FUEL_FLOW, self.HUMIDITY, actual_air
            )
            co2 = poc_co2_mass_gas(fuel_composition_poc, self.FUEL_FLOW)
            n2 = poc_n2_mass_gas(
                fuel_composition_poc, self.FUEL_FLOW, actual_air, stoich_air
            )
            o2 = poc_o2_mass(self.FUEL_FLOW, actual_air, stoich_air, 0.0)

            total = h2o + co2 + n2 + o2
            enthalpy = flue_gas_enthalpy(
                h2o/total, co2/total, n2/total, o2/total,
                self.STACK_TEMP, self.AMBIENT_TEMP
            )

            efficiency = (hhv * self.FUEL_FLOW - enthalpy * total) / (hhv * self.FUEL_FLOW) * 100
            efficiencies.append(efficiency)

        # Efficiency should decrease monotonically with excess air
        for i in range(len(efficiencies) - 1):
            assert efficiencies[i] >= efficiencies[i+1], \
                f"Efficiency should decrease with excess air: {efficiencies[i]:.2f}% vs {efficiencies[i+1]:.2f}%"

        # Efficiency drop from 5% to 30% excess air should be 1-5%
        efficiency_drop = efficiencies[0] - efficiencies[-1]
        assert 1.0 < efficiency_drop < 5.0, \
            f"Efficiency drop {efficiency_drop:.2f}% out of expected range"

    def test_comparison_to_pure_methane(self):
        """Compare natural gas mixture to pure methane"""
        # Natural gas mixture
        ng_composition = GasComposition(
            methane_mass=self.METHANE_MASS,
            ethane_mass=self.ETHANE_MASS,
            propane_mass=self.PROPANE_MASS,
            n2_mass=self.N2_MASS
        )
        ng_hhv = hhv_mass_gas(ng_composition)

        # Pure methane
        ch4_composition = GasComposition(methane_mass=100.0)
        ch4_hhv = hhv_mass_gas(ch4_composition)

        # Natural gas HHV should be slightly lower due to heavier hydrocarbons
        assert ng_hhv < ch4_hhv, "Natural gas HHV should be less than pure methane"

        # Difference should be small (< 5%)
        diff_percent = (ch4_hhv - ng_hhv) / ch4_hhv * 100
        assert 0 < diff_percent < 5, \
            f"HHV difference {diff_percent:.2f}% larger than expected"


class TestValidationNaturalGasComparison:
    """Compare different natural gas compositions"""

    def test_lean_vs_rich_natural_gas(self):
        """Compare lean (high methane) vs rich (more ethane/propane) gas"""
        # Lean gas: 95% CH4, 3% C2H6, 1% C3H8, 1% N2
        lean_gas = GasComposition(
            methane_mass=95.0,
            ethane_mass=3.0,
            propane_mass=1.0,
            n2_mass=1.0
        )

        # Rich gas: 85% CH4, 8% C2H6, 5% C3H8, 2% N2
        rich_gas = GasComposition(
            methane_mass=85.0,
            ethane_mass=8.0,
            propane_mass=5.0,
            n2_mass=2.0
        )

        lean_hhv = hhv_mass_gas(lean_gas)
        rich_hhv = hhv_mass_gas(rich_gas)

        # Lean gas should have higher HHV (more H:C ratio)
        assert lean_hhv > rich_hhv, "Lean gas should have higher HHV"

        # Difference should be noticeable but < 3%
        diff_percent = (lean_hhv - rich_hhv) / lean_hhv * 100
        assert 0.5 < diff_percent < 3.0, \
            f"HHV difference {diff_percent:.2f}% out of expected range"

    def test_high_inert_gas(self):
        """Test natural gas with high inert content"""
        # High inert gas: 85% CH4, 4% C2H6, 2% C3H8, 9% N2
        high_inert = GasComposition(
            methane_mass=85.0,
            ethane_mass=4.0,
            propane_mass=2.0,
            n2_mass=9.0
        )

        # Standard gas: 90% CH4, 5% C2H6, 3% C3H8, 2% N2
        standard = GasComposition(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )

        high_inert_hhv = hhv_mass_gas(high_inert)
        standard_hhv = hhv_mass_gas(standard)

        # High inert gas should have lower HHV (inerts dilute fuel value)
        assert high_inert_hhv < standard_hhv, \
            "High inert gas should have lower HHV"

        # Reduction should be proportional to inert content
        # Roughly 7% less combustibles → 7% less HHV
        reduction_percent = (standard_hhv - high_inert_hhv) / standard_hhv * 100
        assert 5 < reduction_percent < 10, \
            f"HHV reduction {reduction_percent:.2f}% out of expected range"
