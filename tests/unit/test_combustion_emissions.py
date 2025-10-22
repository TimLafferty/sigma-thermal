"""
Unit tests for combustion emissions calculations.

Tests NOx and CO2 emissions from combustion processes.
"""

import pytest
from sigma_thermal.combustion.emissions import (
    nox_emissions,
    co2_emissions,
    NOxEmissions,
    CO2Emissions,
)


class TestNOxEmissions:
    """Tests for nox_emissions function"""

    def test_natural_gas_baseline(self):
        """Test baseline natural gas NOx emissions"""
        # 100 MMBtu/hr, 3200°F (reference), 15% excess air
        result = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )
        # Should be close to baseline emission factor (0.10 lb/MMBtu)
        assert result > 8.0
        assert result < 12.0

    def test_high_temperature_increases_nox(self):
        """Test that higher flame temperature increases NOx"""
        # Same conditions, different temperatures
        result_low = nox_emissions(
            fuel_rate=100,
            flame_temp=3000,
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )

        result_high = nox_emissions(
            fuel_rate=100,
            flame_temp=3600,
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )

        # Higher temperature should produce more NOx (thermal NOx)
        assert result_high > result_low
        assert result_high > result_low * 1.2  # At least 20% increase at +600°F

    def test_excess_air_increases_nox(self):
        """Test that excess air increases NOx"""
        result_low_ea = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=5.0,
            fuel_type="natural_gas"
        )

        result_high_ea = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=25.0,
            fuel_type="natural_gas"
        )

        # More excess air = more O2 = more NOx
        assert result_high_ea > result_low_ea

    def test_oil_higher_nox_than_gas(self):
        """Test that oil produces more NOx than natural gas"""
        result_gas = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )

        result_oil = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="oil"
        )

        # Oil should produce more NOx
        assert result_oil > result_gas

    def test_coal_highest_nox(self):
        """Test that coal produces highest NOx"""
        result_coal = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="coal"
        )

        result_gas = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )

        # Coal has highest base emission factor
        assert result_coal > result_gas * 3  # At least 3x

    def test_fuel_nitrogen_increases_nox(self):
        """Test that fuel nitrogen contributes to NOx"""
        result_no_n = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="oil",
            fuel_nitrogen_pct=0.0
        )

        result_with_n = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="oil",
            fuel_nitrogen_pct=0.5  # 0.5% nitrogen
        )

        # Fuel nitrogen should add to NOx
        assert result_with_n > result_no_n

    def test_residence_time_increases_nox(self):
        """Test that longer residence time increases NOx"""
        result_short = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="natural_gas",
            residence_time=0.3
        )

        result_long = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="natural_gas",
            residence_time=1.0
        )

        # Longer residence time allows more NOx formation
        assert result_long > result_short

    def test_low_temperature_minimal_nox(self):
        """Test minimal NOx formation at low temperatures"""
        result = nox_emissions(
            fuel_rate=100,
            flame_temp=2500,  # Below thermal NOx threshold
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )
        # Should be very low at this temperature
        assert result < 5.0

    def test_zero_fuel_rate_error(self):
        """Test error when fuel rate is zero"""
        with pytest.raises(ValueError, match="Fuel rate must be positive"):
            nox_emissions(
                fuel_rate=0,
                flame_temp=3200,
                excess_air_pct=15.0
            )

    def test_negative_fuel_rate_error(self):
        """Test error when fuel rate is negative"""
        with pytest.raises(ValueError, match="Fuel rate must be positive"):
            nox_emissions(
                fuel_rate=-100,
                flame_temp=3200,
                excess_air_pct=15.0
            )

    def test_low_flame_temp_error(self):
        """Test error when flame temperature is too low"""
        with pytest.raises(ValueError, match="Flame temperature must be at least"):
            nox_emissions(
                fuel_rate=100,
                flame_temp=500,
                excess_air_pct=15.0
            )

    def test_high_flame_temp_error(self):
        """Test error when flame temperature is unreasonably high"""
        with pytest.raises(ValueError, match="Flame temperature .* is unreasonably high"):
            nox_emissions(
                fuel_rate=100,
                flame_temp=6000,
                excess_air_pct=15.0
            )

    def test_zero_residence_time_error(self):
        """Test error when residence time is zero"""
        with pytest.raises(ValueError, match="Residence time must be positive"):
            nox_emissions(
                fuel_rate=100,
                flame_temp=3200,
                excess_air_pct=15.0,
                residence_time=0
            )

    def test_unknown_fuel_type_error(self):
        """Test error for unknown fuel type"""
        with pytest.raises(ValueError, match="Unknown fuel type"):
            nox_emissions(
                fuel_rate=100,
                flame_temp=3200,
                excess_air_pct=15.0,
                fuel_type="unknown_fuel"
            )


class TestCO2Emissions:
    """Tests for co2_emissions function"""

    def test_natural_gas_co2(self):
        """Test CO2 emissions from natural gas"""
        # 100 MMBtu/hr, 75% carbon equivalent
        result = co2_emissions(
            fuel_rate=100,
            carbon_content=0.75,
            fuel_type="natural_gas"
        )
        # EPA factor: 116.9 lb CO2/MMBtu
        expected = 116.9 * 100
        assert result == pytest.approx(expected, rel=0.05)

    def test_oil_co2(self):
        """Test CO2 emissions from oil"""
        # 100 lb/hr oil, 87% carbon
        result = co2_emissions(
            fuel_rate=100,
            carbon_content=0.87,
            fuel_type="oil"
        )
        # Stoichiometric: 87 lb C/hr × (44/12) = 319 lb CO2/hr
        expected = 100 * 0.87 * (44.0 / 12.0)
        assert result == pytest.approx(expected, rel=1e-3)

    def test_coal_co2(self):
        """Test CO2 emissions from coal"""
        # 1000 lb/hr coal, 70% carbon
        result = co2_emissions(
            fuel_rate=1000,
            carbon_content=0.70,
            fuel_type="coal"
        )
        # 700 lb C/hr × (44/12) = 2566.67 lb CO2/hr
        expected = 1000 * 0.70 * (44.0 / 12.0)
        assert result == pytest.approx(expected, rel=1e-3)

    def test_propane_co2(self):
        """Test CO2 emissions from propane"""
        # 50 MMBtu/hr propane
        result = co2_emissions(
            fuel_rate=50,
            carbon_content=0.817,
            fuel_type="propane"
        )
        # EPA factor: 139.0 lb CO2/MMBtu
        expected = 139.0 * 50
        assert result == pytest.approx(expected, rel=0.05)

    def test_higher_carbon_more_co2(self):
        """Test that higher carbon content produces more CO2"""
        result_low = co2_emissions(
            fuel_rate=100,
            carbon_content=0.70,
            fuel_type="coal"
        )

        result_high = co2_emissions(
            fuel_rate=100,
            carbon_content=0.85,
            fuel_type="coal"
        )

        # Proportional relationship
        ratio = result_high / result_low
        expected_ratio = 0.85 / 0.70
        assert ratio == pytest.approx(expected_ratio, rel=1e-3)

    def test_co2_stoichiometry(self):
        """Test CO2 stoichiometric conversion"""
        # 12 lb carbon should produce 44 lb CO2
        result = co2_emissions(
            fuel_rate=12.0 / 0.87,  # Pure carbon (adjusted for 87% content)
            carbon_content=0.87,
            fuel_type="coal"
        )
        expected = 44.0
        assert result == pytest.approx(expected, rel=1e-2)

    def test_zero_fuel_rate_error(self):
        """Test error when fuel rate is zero"""
        with pytest.raises(ValueError, match="Fuel rate must be positive"):
            co2_emissions(
                fuel_rate=0,
                carbon_content=0.75
            )

    def test_negative_fuel_rate_error(self):
        """Test error when fuel rate is negative"""
        with pytest.raises(ValueError, match="Fuel rate must be positive"):
            co2_emissions(
                fuel_rate=-100,
                carbon_content=0.75
            )

    def test_carbon_content_below_zero_error(self):
        """Test error when carbon content is negative"""
        with pytest.raises(ValueError, match="Carbon content must be between 0 and 1"):
            co2_emissions(
                fuel_rate=100,
                carbon_content=-0.1
            )

    def test_carbon_content_above_one_error(self):
        """Test error when carbon content exceeds 1"""
        with pytest.raises(ValueError, match="Carbon content must be between 0 and 1"):
            co2_emissions(
                fuel_rate=100,
                carbon_content=1.5
            )

    def test_zero_carbon_minimal_co2(self):
        """Test that zero carbon content produces minimal CO2"""
        # For gaseous fuels, emission factor is used even with 0 carbon content
        result = co2_emissions(
            fuel_rate=100,
            carbon_content=0.0,
            fuel_type="natural_gas"
        )
        # Natural gas uses emission factor, so result will be small but non-zero
        assert result >= 0.0

        # For solid/liquid fuels, zero carbon should give zero CO2
        result_solid = co2_emissions(
            fuel_rate=100,
            carbon_content=0.0,
            fuel_type="coal"
        )
        assert result_solid == 0.0


class TestVBACompatibility:
    """Tests for VBA-compatible wrapper functions"""

    def test_nox_emissions_vba(self):
        """Test VBA wrapper for nox_emissions"""
        result = NOxEmissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )
        assert result > 8.0
        assert result < 12.0

    def test_nox_emissions_vba_all_params(self):
        """Test VBA wrapper with all parameters"""
        result = NOxEmissions(
            fuel_rate=100,
            flame_temp=3200,
            excess_air_pct=15.0,
            fuel_type="oil",
            fuel_nitrogen_pct=0.3,
            residence_time=0.7
        )
        # Oil with fuel nitrogen should have elevated NOx
        assert result > 15.0

    def test_co2_emissions_vba(self):
        """Test VBA wrapper for co2_emissions"""
        result = CO2Emissions(
            fuel_rate=100,
            carbon_content=0.75,
            fuel_type="natural_gas"
        )
        expected = 116.9 * 100
        assert result == pytest.approx(expected, rel=0.05)

    def test_co2_emissions_vba_coal(self):
        """Test VBA wrapper for coal CO2"""
        result = CO2Emissions(
            fuel_rate=1000,
            carbon_content=0.70,
            fuel_type="coal"
        )
        expected = 1000 * 0.70 * (44.0 / 12.0)
        assert result == pytest.approx(expected, rel=1e-3)


class TestIntegrationScenarios:
    """Integration tests with realistic combustion scenarios"""

    def test_natural_gas_boiler_emissions(self):
        """Test complete emissions from natural gas boiler"""
        # 100 MMBtu/hr natural gas boiler, 3600°F, 10% excess air
        fuel_rate = 100  # MMBtu/hr
        flame_temp = 3600  # °F
        excess_air = 10.0  # %
        carbon_content = 0.75  # Natural gas equivalent

        # Calculate NOx
        nox = nox_emissions(
            fuel_rate=fuel_rate,
            flame_temp=flame_temp,
            excess_air_pct=excess_air,
            fuel_type="natural_gas"
        )

        # Calculate CO2
        co2 = co2_emissions(
            fuel_rate=fuel_rate,
            carbon_content=carbon_content,
            fuel_type="natural_gas"
        )

        # Verify reasonable ranges
        assert nox > 5.0 and nox < 20.0  # Typical: 0.08-0.12 lb/MMBtu
        assert co2 > 11000 and co2 < 12000  # ~117 lb CO2/MMBtu

    def test_oil_fired_heater_emissions(self):
        """Test complete emissions from oil-fired heater"""
        # 75 MMBtu/hr oil heater, 3500°F, 15% excess air
        fuel_rate = 75
        flame_temp = 3500
        excess_air = 15.0
        carbon_content = 0.87

        # Calculate NOx
        nox = nox_emissions(
            fuel_rate=fuel_rate,
            flame_temp=flame_temp,
            excess_air_pct=excess_air,
            fuel_type="oil",
            fuel_nitrogen_pct=0.2
        )

        # Calculate CO2 (note: for oil, treat as lb/hr not MMBtu/hr for carbon calc)
        # Typical: ~19,000 BTU/lb for #2 oil, so 75 MMBtu/hr ≈ 3947 lb/hr
        fuel_lb_hr = 75 * 1e6 / 19000  # Convert MMBtu/hr to lb/hr
        co2 = co2_emissions(
            fuel_rate=fuel_lb_hr,
            carbon_content=carbon_content,
            fuel_type="oil"
        )

        # Oil should have higher NOx than gas
        assert nox > 10.0

        # CO2 from carbon mass balance
        assert co2 > 12000  # Oil has more CO2/MMBtu than gas

    def test_emissions_reduction_with_low_nox_burner(self):
        """Test emissions reduction with low NOx burner (lower flame temp)"""
        # Standard burner
        nox_standard = nox_emissions(
            fuel_rate=100,
            flame_temp=3600,
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )

        # Low NOx burner (reduced flame temperature through staging)
        nox_low_nox = nox_emissions(
            fuel_rate=100,
            flame_temp=3200,  # Reduced by 400°F
            excess_air_pct=15.0,
            fuel_type="natural_gas"
        )

        # Should achieve 10-30% NOx reduction with temperature staging
        reduction_pct = ((nox_standard - nox_low_nox) / nox_standard) * 100
        assert reduction_pct > 10  # At least 10% reduction
        assert reduction_pct < 40  # Not more than 40%

    def test_co2_intensity_comparison(self):
        """Test CO2 intensity comparison across fuel types"""
        # Same energy input (100 MMBtu/hr)
        fuel_rate_mmbtu = 100

        # Natural gas
        co2_gas = co2_emissions(
            fuel_rate=fuel_rate_mmbtu,
            carbon_content=0.75,
            fuel_type="natural_gas"
        )

        # Calculate equivalent lb/hr for coal (assuming 12,000 BTU/lb)
        coal_lb_hr = fuel_rate_mmbtu * 1e6 / 12000

        # Coal
        co2_coal = co2_emissions(
            fuel_rate=coal_lb_hr,
            carbon_content=0.70,
            fuel_type="coal"
        )

        # Coal should produce more CO2/MMBtu than natural gas
        assert co2_coal > co2_gas

    def test_emissions_with_air_preheat(self):
        """Test NOx increases with air preheat (higher flame temp)"""
        # Without preheat
        nox_ambient = nox_emissions(
            fuel_rate=100,
            flame_temp=3600,
            excess_air_pct=10.0,
            fuel_type="natural_gas"
        )

        # With air preheat (higher flame temperature)
        nox_preheat = nox_emissions(
            fuel_rate=100,
            flame_temp=4000,  # Higher due to preheat
            excess_air_pct=10.0,
            fuel_type="natural_gas"
        )

        # Preheated air increases NOx moderately
        assert nox_preheat > nox_ambient * 1.1  # At least 10% increase
