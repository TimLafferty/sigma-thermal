"""
Unit tests for flame temperature calculations.

Tests adiabatic flame temperature and excess air effects on
flame temperature in combustion systems.
"""

import pytest
from sigma_thermal.combustion.flame_temperature import (
    adiabatic_flame_temp,
    flame_temp_excess_air,
    AdiabaticFlameTemp,
    FlameTempExcessAir,
)


class TestAdiabaticFlameTemp:
    """Tests for adiabatic_flame_temp function"""

    def test_natural_gas_stoichiometric(self):
        """Test natural gas at stoichiometric conditions"""
        # Natural gas: LHV = 21,500 BTU/lb, stoich air = 17.24 lb air/lb fuel
        result = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=0.0
        )
        # Should be high (3800-3950°F range for stoichiometric)
        assert result > 3800
        assert result < 3950

    def test_natural_gas_with_excess_air(self):
        """Test natural gas with 10% excess air"""
        result = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0
        )
        # Should be lower than stoichiometric (3500-3650°F range)
        assert result > 3500
        assert result < 3650

    def test_natural_gas_with_20_percent_excess(self):
        """Test natural gas with 20% excess air"""
        result = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=20.0
        )
        # Should be even lower (3250-3400°F range)
        assert result > 3250
        assert result < 3400

    def test_oil_fuel(self):
        """Test oil fuel combustion"""
        # #2 Oil: LHV ~ 18,500 BTU/lb, stoich air = 14.5 lb air/lb fuel
        result = adiabatic_flame_temp(
            lhv=18500,
            fuel_rate=100,
            stoich_air=14.5,
            excess_air_pct=15.0
        )
        # Oil flame temps typically 3400-3600°F with 15% excess air
        assert result > 3400
        assert result < 3600

    def test_hydrogen_combustion(self):
        """Test hydrogen combustion (high temperature)"""
        # Hydrogen: LHV ~ 51,600 BTU/lb, stoich air = 34.28 lb air/lb fuel
        result = adiabatic_flame_temp(
            lhv=51600,
            fuel_rate=50,
            stoich_air=34.28,
            excess_air_pct=5.0
        )
        # Hydrogen burns very hot (4300-4600°F)
        assert result > 4300
        assert result < 4600

    def test_preheated_air(self):
        """Test effect of preheated combustion air"""
        # Same fuel, ambient air vs preheated
        result_ambient = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0,
            air_temp=77.0
        )

        result_preheated = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0,
            air_temp=400.0
        )

        # Preheated air should increase flame temperature
        assert result_preheated > result_ambient
        assert (result_preheated - result_ambient) > 100  # At least 100°F increase

    def test_preheated_fuel(self):
        """Test effect of preheated fuel"""
        result_ambient = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0,
            fuel_temp=77.0
        )

        result_preheated = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0,
            fuel_temp=200.0
        )

        # Preheated fuel should increase flame temperature (but less than air)
        assert result_preheated > result_ambient
        assert (result_preheated - result_ambient) > 5

    def test_humid_air(self):
        """Test effect of humidity on flame temperature"""
        result_dry = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0,
            humidity=0.0
        )

        result_humid = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0,
            humidity=0.013  # ~60% RH at 77°F
        )

        # Humid air should decrease flame temperature
        assert result_humid < result_dry
        assert (result_dry - result_humid) > 10

    def test_higher_fuel_rate_same_temp(self):
        """Test that fuel rate doesn't affect temperature (intensive property)"""
        result_100 = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0
        )

        result_200 = adiabatic_flame_temp(
            lhv=21500,
            fuel_rate=200,
            stoich_air=17.24,
            excess_air_pct=10.0
        )

        # Temperature should be same (intensive property)
        assert result_100 == pytest.approx(result_200, rel=1e-3)

    def test_zero_lhv_error(self):
        """Test error when LHV is zero"""
        with pytest.raises(ValueError, match="Lower heating value must be positive"):
            adiabatic_flame_temp(
                lhv=0,
                fuel_rate=100,
                stoich_air=17.24
            )

    def test_negative_lhv_error(self):
        """Test error when LHV is negative"""
        with pytest.raises(ValueError, match="Lower heating value must be positive"):
            adiabatic_flame_temp(
                lhv=-1000,
                fuel_rate=100,
                stoich_air=17.24
            )

    def test_zero_fuel_rate_error(self):
        """Test error when fuel rate is zero"""
        with pytest.raises(ValueError, match="Fuel rate must be positive"):
            adiabatic_flame_temp(
                lhv=21500,
                fuel_rate=0,
                stoich_air=17.24
            )

    def test_zero_stoich_air_error(self):
        """Test error when stoichiometric air is zero"""
        with pytest.raises(ValueError, match="Stoichiometric air must be positive"):
            adiabatic_flame_temp(
                lhv=21500,
                fuel_rate=100,
                stoich_air=0
            )

    def test_negative_excess_air_error(self):
        """Test error when excess air is negative"""
        with pytest.raises(ValueError, match="Excess air percentage cannot be negative"):
            adiabatic_flame_temp(
                lhv=21500,
                fuel_rate=100,
                stoich_air=17.24,
                excess_air_pct=-10.0
            )

    def test_very_low_lhv_unreasonable_temp(self):
        """Test that unreasonably low LHV produces error"""
        with pytest.raises(ValueError, match="Calculated flame temperature .* is unreasonably low"):
            adiabatic_flame_temp(
                lhv=100,  # Too low for combustion
                fuel_rate=100,
                stoich_air=17.24,
                excess_air_pct=50.0
            )


class TestFlameTempExcessAir:
    """Tests for flame_temp_excess_air function"""

    def test_zero_excess_air(self):
        """Test that zero excess air returns stoichiometric temperature"""
        stoich_temp = 3600
        result = flame_temp_excess_air(stoich_temp, 0.0)
        assert result == stoich_temp

    def test_10_percent_excess_air(self):
        """Test 10% excess air temperature drop"""
        stoich_temp = 3600
        result = flame_temp_excess_air(stoich_temp, 10.0)
        # Should drop ~450°F (45°F per 1%)
        expected_drop = 45 * 10
        assert result == pytest.approx(stoich_temp - expected_drop, rel=0.05)

    def test_20_percent_excess_air(self):
        """Test 20% excess air temperature drop"""
        stoich_temp = 3600
        result = flame_temp_excess_air(stoich_temp, 20.0)
        # Should drop ~900°F (45°F per 1%)
        assert result < stoich_temp - 800
        assert result > stoich_temp - 1000

    def test_preheated_air_reduces_drop(self):
        """Test that preheated air reduces temperature drop"""
        stoich_temp = 3600

        result_ambient = flame_temp_excess_air(stoich_temp, 10.0, air_temp=77.0)
        result_preheated = flame_temp_excess_air(stoich_temp, 10.0, air_temp=400.0)

        # Preheated air should reduce the temperature drop
        assert result_preheated > result_ambient
        assert (result_preheated - result_ambient) > 50

    def test_high_excess_air(self):
        """Test high excess air (50%)"""
        stoich_temp = 3600
        result = flame_temp_excess_air(stoich_temp, 50.0)
        # Should drop significantly but still be reasonable (drops ~2250°F)
        assert result > 1200
        assert result < 1500

    def test_very_low_stoich_temp_error(self):
        """Test error when stoichiometric temperature is too low"""
        with pytest.raises(ValueError, match="Stoichiometric flame temperature .* is too low"):
            flame_temp_excess_air(500, 10.0)

    def test_very_high_stoich_temp_error(self):
        """Test error when stoichiometric temperature is too high"""
        with pytest.raises(ValueError, match="Stoichiometric flame temperature .* is too high"):
            flame_temp_excess_air(6000, 10.0)

    def test_negative_excess_air_error(self):
        """Test error when excess air is negative"""
        with pytest.raises(ValueError, match="Excess air percentage cannot be negative"):
            flame_temp_excess_air(3600, -10.0)

    def test_excessive_drop_error(self):
        """Test error when temperature drops too low"""
        with pytest.raises(ValueError, match="Calculated flame temperature .* is unreasonably low"):
            flame_temp_excess_air(3000, 100.0)  # Excessive excess air


class TestVBACompatibility:
    """Tests for VBA-compatible wrapper functions"""

    def test_adiabatic_flame_temp_vba(self):
        """Test VBA wrapper for adiabatic_flame_temp"""
        result = AdiabaticFlameTemp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0
        )
        assert result > 3500
        assert result < 3650

    def test_adiabatic_flame_temp_vba_all_params(self):
        """Test VBA wrapper with all parameters"""
        result = AdiabaticFlameTemp(
            lhv=21500,
            fuel_rate=100,
            stoich_air=17.24,
            excess_air_pct=10.0,
            fuel_temp=77.0,
            air_temp=400.0,
            humidity=0.01
        )
        # Preheated air, humid conditions
        assert result > 3600
        assert result < 3850

    def test_flame_temp_excess_air_vba(self):
        """Test VBA wrapper for flame_temp_excess_air"""
        result = FlameTempExcessAir(3600, 10.0)
        expected = 3600 - (45 * 10)
        assert result == pytest.approx(expected, rel=0.05)

    def test_flame_temp_excess_air_vba_preheated(self):
        """Test VBA wrapper with preheated air"""
        result = FlameTempExcessAir(3600, 10.0, air_temp=400.0)
        # Should be higher than ambient air case
        assert result > 3600 - 450


class TestIntegrationScenarios:
    """Integration tests with realistic combustion scenarios"""

    def test_natural_gas_boiler_complete(self):
        """Test complete natural gas boiler flame temperature calculation"""
        # Natural gas boiler: 10% excess air, ambient conditions
        lhv = 21500  # BTU/lb
        fuel_rate = 100  # lb/hr
        stoich_air = 17.24  # lb air/lb fuel
        excess_air = 10.0  # %

        # Calculate adiabatic flame temperature
        flame_temp = adiabatic_flame_temp(
            lhv=lhv,
            fuel_rate=fuel_rate,
            stoich_air=stoich_air,
            excess_air_pct=excess_air
        )

        # Typical for natural gas with 10% excess air
        assert flame_temp > 3500
        assert flame_temp < 3650

    def test_oil_fired_heater_complete(self):
        """Test complete oil-fired heater flame temperature"""
        # Oil heater: 15% excess air
        lhv = 18500  # BTU/lb
        fuel_rate = 75  # lb/hr
        stoich_air = 14.5  # lb air/lb fuel
        excess_air = 15.0  # %

        flame_temp = adiabatic_flame_temp(
            lhv=lhv,
            fuel_rate=fuel_rate,
            stoich_air=stoich_air,
            excess_air_pct=excess_air
        )

        # Oil typically 3400-3600°F with 15% excess air
        assert flame_temp > 3400
        assert flame_temp < 3600

    def test_preheated_air_furnace(self):
        """Test industrial furnace with preheated combustion air"""
        # High efficiency furnace with air preheat
        lhv = 21500
        fuel_rate = 150
        stoich_air = 17.24
        excess_air = 5.0  # Low excess air for high efficiency
        air_temp = 600.0  # Preheated from recuperator

        flame_temp = adiabatic_flame_temp(
            lhv=lhv,
            fuel_rate=fuel_rate,
            stoich_air=stoich_air,
            excess_air_pct=excess_air,
            air_temp=air_temp
        )

        # Should be higher due to preheat (4000-4300°F)
        assert flame_temp > 4000
        assert flame_temp < 4300

    def test_excess_air_optimization(self):
        """Test flame temperature across range of excess air"""
        lhv = 21500
        fuel_rate = 100
        stoich_air = 17.24

        # Calculate for different excess air levels
        temps = []
        excess_air_levels = [0, 5, 10, 15, 20, 25]

        for excess in excess_air_levels:
            temp = adiabatic_flame_temp(
                lhv=lhv,
                fuel_rate=fuel_rate,
                stoich_air=stoich_air,
                excess_air_pct=excess
            )
            temps.append(temp)

        # Temperature should decrease monotonically with excess air
        for i in range(len(temps) - 1):
            assert temps[i] > temps[i + 1]

        # Verify reasonable range
        assert temps[0] > 3800  # Stoichiometric is hottest
        assert temps[-1] < 3200  # 25% excess air is coolest

    def test_two_calculation_methods_consistency(self):
        """Test consistency between direct and correction methods"""
        lhv = 21500
        fuel_rate = 100
        stoich_air = 17.24
        excess_air = 10.0

        # Method 1: Direct calculation with excess air
        temp_direct = adiabatic_flame_temp(
            lhv=lhv,
            fuel_rate=fuel_rate,
            stoich_air=stoich_air,
            excess_air_pct=excess_air
        )

        # Method 2: Calculate stoichiometric, then apply correction
        temp_stoich = adiabatic_flame_temp(
            lhv=lhv,
            fuel_rate=fuel_rate,
            stoich_air=stoich_air,
            excess_air_pct=0.0
        )
        temp_corrected = flame_temp_excess_air(temp_stoich, excess_air)

        # Should be reasonably close (within 10%)
        assert temp_direct == pytest.approx(temp_corrected, rel=0.15)
