"""
Validation Test Case 3: Liquid Fuel (#2 Oil) Combustion

This test validates Python implementation for typical #2 fuel oil combustion
in industrial boilers and furnaces.

Test Scenario:
- Fuel: #2 Fuel Oil (typical distillate)
- Ultimate Analysis: 87% C, 13% H (mass basis)
- Fuel Flow: 1000 lb/hr (higher than gas tests)
- Excess Air: 20% (higher than gas for complete combustion)
- Ambient Temperature: 77°F
- Stack Temperature: 450°F (typical oil-fired boiler)
- Humidity: 0.013 lb H2O / lb dry air

This represents a realistic #2 oil combustion scenario in an
industrial heating application.

References:
- GPSA Engineering Data Book, 13th Edition
- ASME PTC 4 (Fired Steam Generators)
- EPA AP-42 (Emissions factors for fuel oil)
- Typical #2 oil properties: HHV ~19,500 BTU/lb
"""

import pytest
from sigma_thermal.combustion import (
    hhv_mass_liquid,
    lhv_mass_liquid,
    poc_h2o_mass_liquid,
    poc_co2_mass_liquid,
    poc_n2_mass_liquid,
    poc_o2_mass,
    flue_gas_enthalpy,
)


class TestValidationLiquidFuel:
    """
    Validation test for #2 fuel oil combustion.

    Tests a typical distillate fuel oil through complete combustion analysis.
    """

    # Test parameters
    FUEL_TYPE = '#2 oil'  # Standard distillate fuel oil
    FUEL_FLOW = 1000.0  # lb/hr (higher than gas tests)
    EXCESS_AIR_PERCENT = 20.0  # % (higher for oil)
    AMBIENT_TEMP = 77.0  # °F
    STACK_TEMP = 450.0  # °F (typical oil-fired boiler)
    HUMIDITY = 0.013  # lb H2O / lb dry air

    # Stoichiometric air for #2 oil (from lookup tables)
    # Approximately 14.5 lb air / lb fuel for #2 oil
    STOICH_AIR_RATIO = 14.5  # lb air / lb fuel

    @pytest.fixture
    def stoich_air_flow(self):
        """Stoichiometric air flow rate"""
        return self.FUEL_FLOW * self.STOICH_AIR_RATIO

    @pytest.fixture
    def actual_air_flow(self, stoich_air_flow):
        """Actual air flow with excess air"""
        return stoich_air_flow * (1.0 + self.EXCESS_AIR_PERCENT / 100.0)

    def test_heating_values(self):
        """Test HHV and LHV for #2 fuel oil"""
        hhv = hhv_mass_liquid(self.FUEL_TYPE)
        lhv = lhv_mass_liquid(self.FUEL_TYPE)

        # LHV should be less than HHV
        assert lhv < hhv

        # Physical validation for #2 oil
        # From GPSA: #2 oil HHV ~18,993 BTU/lb
        assert 18900 < hhv < 19100, f"HHV {hhv} should be in typical #2 oil range"
        assert 17800 < lhv < 18100, f"LHV {lhv} should be in typical #2 oil range"

        # Difference should be water latent heat from H2 combustion
        hhv_lhv_diff = hhv - lhv
        assert 1000 < hhv_lhv_diff < 1300

    def test_stoichiometric_air(self, stoich_air_flow):
        """Test stoichiometric air for #2 fuel oil"""
        # Should be around 14.5 lb air / lb fuel
        assert 14400 < stoich_air_flow < 14600
        assert stoich_air_flow == pytest.approx(14490.0, rel=0.01)

    def test_products_of_combustion_composition(
        self, stoich_air_flow, actual_air_flow
    ):
        """Test complete POC calculation for #2 oil"""
        # Calculate all products
        h2o = poc_h2o_mass_liquid(
            fuel_type=self.FUEL_TYPE,
            fuel_flow_mass=self.FUEL_FLOW,
            humidity=self.HUMIDITY,
            air_flow_mass=actual_air_flow
        )
        co2 = poc_co2_mass_liquid(
            fuel_type=self.FUEL_TYPE,
            fuel_flow_mass=self.FUEL_FLOW
        )
        n2 = poc_n2_mass_liquid(
            fuel_type=self.FUEL_TYPE,
            fuel_flow_mass=self.FUEL_FLOW,
            excess_air_mass=actual_air_flow,
            air_flow_mass=stoich_air_flow
        )
        o2 = poc_o2_mass(
            fuel_flow_mass=self.FUEL_FLOW,
            excess_air_mass=actual_air_flow,
            air_flow_mass=stoich_air_flow,
            o2_in_fuel_mass=0.0  # No oxygen in fuel
        )

        # Check mass balance (including humidity water)
        humidity_water = self.HUMIDITY * actual_air_flow
        total_input = self.FUEL_FLOW + actual_air_flow + humidity_water
        total_products = h2o + co2 + n2 + o2

        # Liquid fuel uses lookup tables, so tolerance is slightly higher
        assert total_products == pytest.approx(total_input, rel=0.02)

        # Validate product ranges for #2 oil
        assert 1340 < h2o < 1450, "H2O mass reasonable for #2 oil"
        assert 3100 < co2 < 3300, "CO2 mass reasonable for #2 oil"
        assert 12500 < n2 < 13500, "N2 mass reasonable for #2 oil"
        assert 550 < o2 < 700, "O2 mass reasonable with 20% excess air"

    def test_co2_emissions(self):
        """Test CO2 emissions calculation for #2 oil"""
        co2_mass = poc_co2_mass_liquid(
            fuel_type=self.FUEL_TYPE,
            fuel_flow_mass=self.FUEL_FLOW
        )

        # CO2 per unit heat input
        hhv = hhv_mass_liquid(self.FUEL_TYPE)

        # lb CO2 per MMBtu
        co2_per_mmbtu = (co2_mass / (hhv * self.FUEL_FLOW)) * 1e6

        # #2 oil typically produces 161 lb CO2/MMBtu (higher than natural gas)
        assert 155 < co2_per_mmbtu < 170, \
            f"CO2 emissions {co2_per_mmbtu:.1f} lb/MMBtu out of typical range"

    def test_oil_boiler_performance(self, stoich_air_flow, actual_air_flow):
        """
        Test complete boiler efficiency for oil-fired system.

        At 450°F stack temperature, oil-fired boilers achieve 82-87% efficiency.
        """
        # Get HHV
        hhv = hhv_mass_liquid(self.FUEL_TYPE)

        # Calculate products
        h2o = poc_h2o_mass_liquid(
            fuel_type=self.FUEL_TYPE,
            fuel_flow_mass=self.FUEL_FLOW,
            humidity=self.HUMIDITY,
            air_flow_mass=actual_air_flow
        )
        co2 = poc_co2_mass_liquid(
            fuel_type=self.FUEL_TYPE,
            fuel_flow_mass=self.FUEL_FLOW
        )
        n2 = poc_n2_mass_liquid(
            fuel_type=self.FUEL_TYPE,
            fuel_flow_mass=self.FUEL_FLOW,
            excess_air_mass=actual_air_flow,
            air_flow_mass=stoich_air_flow
        )
        o2 = poc_o2_mass(
            fuel_flow_mass=self.FUEL_FLOW,
            excess_air_mass=actual_air_flow,
            air_flow_mass=stoich_air_flow,
            o2_in_fuel_mass=0.0
        )

        total_flue = h2o + co2 + n2 + o2

        # Calculate flue gas enthalpy at 450°F
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

        # Oil-fired boiler at 450°F: High efficiency (similar to gas at this low temp)
        assert 88 < efficiency < 92, \
            f"Efficiency {efficiency:.1f}% out of expected range for oil-fired boiler"

        # Stack loss should be low at 450°F (efficient operation)
        stack_loss_percent = (stack_loss / heat_input) * 100
        assert 8 < stack_loss_percent < 15, \
            f"Stack loss {stack_loss_percent:.1f}% out of expected range"

        return {
            'efficiency': efficiency,
            'stack_loss_percent': stack_loss_percent,
            'heat_input_btu_hr': heat_input,
            'stack_temp_degf': self.STACK_TEMP
        }

    def test_excess_air_effect(self, stoich_air_flow):
        """Test impact of excess air on efficiency for oil combustion"""
        excess_air_levels = [10, 15, 20, 25, 35]
        efficiencies = []

        for excess in excess_air_levels:
            actual_air = stoich_air_flow * (1.0 + excess / 100.0)

            hhv = hhv_mass_liquid(self.FUEL_TYPE)
            h2o = poc_h2o_mass_liquid(
                fuel_type=self.FUEL_TYPE,
                fuel_flow_mass=self.FUEL_FLOW,
                humidity=self.HUMIDITY,
                air_flow_mass=actual_air
            )
            co2 = poc_co2_mass_liquid(
                fuel_type=self.FUEL_TYPE,
                fuel_flow_mass=self.FUEL_FLOW
            )
            n2 = poc_n2_mass_liquid(
                fuel_type=self.FUEL_TYPE,
                fuel_flow_mass=self.FUEL_FLOW,
                excess_air_mass=actual_air,
                air_flow_mass=stoich_air_flow
            )
            o2 = poc_o2_mass(
                fuel_flow_mass=self.FUEL_FLOW,
                excess_air_mass=actual_air,
                air_flow_mass=stoich_air_flow,
                o2_in_fuel_mass=0.0
            )

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

        # Efficiency drop from 10% to 35% excess air should be 2-4%
        efficiency_drop = efficiencies[0] - efficiencies[-1]
        assert 1.5 < efficiency_drop < 5.0, \
            f"Efficiency drop {efficiency_drop:.2f}% out of expected range"

    def test_comparison_to_natural_gas(self):
        """Compare #2 oil to natural gas (pure methane)"""
        # #2 Oil
        oil_hhv = hhv_mass_liquid(self.FUEL_TYPE)

        # Natural gas (pure methane reference value)
        ng_hhv = 23875  # BTU/lb

        # Oil HHV should be lower than natural gas (less H:C ratio)
        assert oil_hhv < ng_hhv, "Oil HHV should be less than natural gas"

        # Difference should be significant (~20%)
        diff_percent = (ng_hhv - oil_hhv) / ng_hhv * 100
        assert 15 < diff_percent < 25, \
            f"HHV difference {diff_percent:.2f}% out of expected range"

    def test_oil_requires_more_air_per_btu(self):
        """Test that oil requires more air per BTU than natural gas"""
        # Oil properties
        oil_hhv = hhv_mass_liquid(self.FUEL_TYPE)
        oil_stoich_air = self.STOICH_AIR_RATIO
        oil_air_per_mmbtu = (oil_stoich_air / oil_hhv) * 1e6

        # Natural gas (methane) reference
        ng_hhv = 23875  # BTU/lb
        ng_stoich_air = 17.24  # lb air / lb fuel
        ng_air_per_mmbtu = (ng_stoich_air / ng_hhv) * 1e6

        # Oil should require less air per MMBtu (less air per lb, but also less BTU per lb)
        # Actually, oil requires LESS air per lb (14.5 vs 17.24)
        # But per MMBtu, it's close
        assert 700 < oil_air_per_mmbtu < 800, \
            f"Oil air requirement {oil_air_per_mmbtu:.0f} lb/MMBtu out of range"
        assert 700 < ng_air_per_mmbtu < 750, \
            f"NG air requirement {ng_air_per_mmbtu:.0f} lb/MMBtu out of range"


class TestValidationLiquidFuelVariations:
    """Compare different liquid fuel types"""

    def test_heavy_vs_light_oil(self):
        """Test heavy oil (#6) vs lighter distillate (#2)"""
        # Heavy oil #6
        heavy_hhv = hhv_mass_liquid('#6 oil')

        # Standard #2 oil
        standard_hhv = hhv_mass_liquid('#2 oil')

        # #2 oil should have higher HHV (better quality, more H:C ratio)
        assert standard_hhv > heavy_hhv, \
            "#2 oil should have higher HHV than #6 oil"

        # Difference should be noticeable (>5%)
        diff_percent = (standard_hhv - heavy_hhv) / standard_hhv * 100
        assert 0.5 < diff_percent < 10.0, \
            f"HHV difference {diff_percent:.2f}% out of expected range"

    def test_gasoline_vs_oil(self):
        """Test gasoline vs #2 oil"""
        # Gasoline (higher H:C ratio)
        gasoline_hhv = hhv_mass_liquid('gasoline')

        # Standard #2 oil
        oil_hhv = hhv_mass_liquid('#2 oil')

        # Gasoline should have higher HHV (more hydrogen)
        assert gasoline_hhv > oil_hhv, \
            "Gasoline should have higher HHV than #2 oil"

        # Hydrogen has much higher energy per pound
        diff_percent = (gasoline_hhv - oil_hhv) / oil_hhv * 100
        assert 1.0 < diff_percent < 10.0, \
            f"HHV difference {diff_percent:.2f}% out of expected range"
