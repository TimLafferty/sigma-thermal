"""
Unit tests for water and steam property functions.

Tests the saturation pressure and temperature functions against
ASME steam table values and validates accuracy, error handling,
and VBA compatibility.

Author: GTS Energy Inc.
Date: October 2025
"""

import pytest
import math
from sigma_thermal.fluids.water_properties import (
    saturation_pressure,
    saturation_temperature,
    water_density,
    water_viscosity,
    water_specific_heat,
    water_thermal_conductivity,
    steam_enthalpy,
    steam_quality,
    SaturationPressure,
    SaturationTemperature,
    WaterDensity,
    WaterViscosity,
    WaterSpecificHeat,
    WaterThermalConductivity,
    SteamEnthalpy,
    SteamQuality,
)


class TestSaturationPressure:
    """Tests for saturation_pressure() function."""

    def test_freezing_point(self):
        """Test at freezing point (32°F)."""
        result = saturation_pressure(32.0)
        # ASME Steam Tables: 0.08854 psia at 32°F
        assert abs(result - 0.08854) < 0.001

    def test_boiling_point_atmospheric(self):
        """Test at atmospheric boiling point (212°F)."""
        result = saturation_pressure(212.0)
        # ASME Steam Tables: 14.696 psia at 212°F
        assert abs(result - 14.696) < 0.05

    def test_300_degf(self):
        """Test at 300°F (common process temperature)."""
        result = saturation_pressure(300.0)
        # ASME Steam Tables: 67.01 psia at 300°F
        assert abs(result - 67.01) < 1.0

    def test_400_degf(self):
        """Test at 400°F (high pressure steam)."""
        result = saturation_pressure(400.0)
        # ASME Steam Tables: 247.26 psia at 400°F
        assert abs(result - 247.26) < 5.0

    def test_600_degf(self):
        """Test at 600°F (near critical point)."""
        result = saturation_pressure(600.0)
        # ASME Steam Tables: 1542.9 psia at 600°F
        assert abs(result - 1542.9) < 50.0

    def test_monotonic_increase(self):
        """Test that pressure increases with temperature."""
        temps = [100, 150, 200, 250, 300, 350, 400]
        pressures = [saturation_pressure(t) for t in temps]

        # Pressure should increase monotonically
        for i in range(len(pressures) - 1):
            assert pressures[i + 1] > pressures[i]

    def test_below_freezing_error(self):
        """Test error for temperature below freezing point."""
        with pytest.raises(ValueError, match="below freezing point"):
            saturation_pressure(20.0)

    def test_above_critical_error(self):
        """Test error for temperature above critical point."""
        with pytest.raises(ValueError, match="exceeds critical temperature"):
            saturation_pressure(800.0)

    def test_negative_temperature_error(self):
        """Test error for negative temperature."""
        with pytest.raises(ValueError, match="below freezing point"):
            saturation_pressure(-10.0)


class TestSaturationTemperature:
    """Tests for saturation_temperature() function."""

    def test_atmospheric_pressure(self):
        """Test at atmospheric pressure (14.696 psia)."""
        result = saturation_temperature(14.696)
        # Should return 212°F (boiling point)
        assert abs(result - 212.0) < 0.5

    def test_vacuum_conditions(self):
        """Test at vacuum pressure (1 psia)."""
        result = saturation_temperature(1.0)
        # ASME Steam Tables: ~101.74°F at 1 psia
        assert abs(result - 101.74) < 1.0

    def test_100_psia(self):
        """Test at 100 psia (common boiler pressure)."""
        result = saturation_temperature(100.0)
        # ASME Steam Tables: 327.82°F at 100 psia
        assert abs(result - 327.82) < 1.0

    def test_200_psia(self):
        """Test at 200 psia (industrial steam)."""
        result = saturation_temperature(200.0)
        # ASME Steam Tables: 381.80°F at 200 psia
        assert abs(result - 381.80) < 2.0

    def test_1000_psia(self):
        """Test at 1000 psia (high pressure steam)."""
        result = saturation_temperature(1000.0)
        # ASME Steam Tables: 544.58°F at 1000 psia
        assert abs(result - 544.58) < 3.0

    def test_monotonic_increase(self):
        """Test that temperature increases with pressure."""
        pressures = [10, 50, 100, 200, 500, 1000]
        temps = [saturation_temperature(p) for p in pressures]

        # Temperature should increase monotonically
        for i in range(len(temps) - 1):
            assert temps[i + 1] > temps[i]

    def test_below_triple_point_error(self):
        """Test error for pressure below triple point."""
        with pytest.raises(ValueError, match="below triple point"):
            saturation_temperature(0.05)

    def test_above_critical_pressure_error(self):
        """Test error for pressure above critical point."""
        with pytest.raises(ValueError, match="exceeds critical pressure"):
            saturation_temperature(3500.0)

    def test_negative_pressure_error(self):
        """Test error for negative pressure."""
        with pytest.raises(ValueError, match="below triple point"):
            saturation_temperature(-1.0)


class TestCrossValidation:
    """Cross-validation tests between pressure and temperature functions."""

    def test_temp_to_pressure_to_temp_roundtrip(self):
        """Test T→P→T roundtrip accuracy."""
        test_temps = [100, 212, 300, 400, 500, 600]

        for original_temp in test_temps:
            pressure = saturation_pressure(original_temp)
            recovered_temp = saturation_temperature(pressure)

            # Should recover original temperature within 0.5°F
            assert abs(recovered_temp - original_temp) < 0.5, \
                f"T→P→T failed: {original_temp}°F → {pressure} psia → {recovered_temp}°F"

    def test_pressure_to_temp_to_pressure_roundtrip(self):
        """Test P→T→P roundtrip accuracy."""
        test_pressures = [1.0, 14.696, 50.0, 100.0, 200.0, 500.0]

        for original_pressure in test_pressures:
            temp = saturation_temperature(original_pressure)
            recovered_pressure = saturation_pressure(temp)

            # Should recover original pressure within 1%
            relative_error = abs(recovered_pressure - original_pressure) / original_pressure
            assert relative_error < 0.01, \
                f"P→T→P failed: {original_pressure} psia → {temp}°F → {recovered_pressure} psia"


class TestVBACompatibility:
    """Tests for VBA-compatible wrapper functions."""

    def test_saturation_pressure_wrapper(self):
        """Test SaturationPressure() VBA wrapper."""
        result_vba = SaturationPressure(212.0)
        result_python = saturation_pressure(212.0)

        assert result_vba == result_python
        assert abs(result_vba - 14.696) < 0.05

    def test_saturation_temperature_wrapper(self):
        """Test SaturationTemperature() VBA wrapper."""
        result_vba = SaturationTemperature(14.696)
        result_python = saturation_temperature(14.696)

        assert result_vba == result_python
        assert abs(result_vba - 212.0) < 0.5


class TestIntegrationScenarios:
    """Integration tests with realistic engineering scenarios."""

    def test_low_pressure_boiler(self):
        """Test typical low-pressure boiler conditions (15 psig)."""
        # 15 psig = 29.696 psia (gauge + atmospheric)
        pressure = 29.696
        temp = saturation_temperature(pressure)

        # Should be around 250°F
        assert temp > 240 and temp < 260

        # Verify roundtrip
        pressure_check = saturation_pressure(temp)
        assert abs(pressure_check - pressure) / pressure < 0.01

    def test_high_pressure_steam_system(self):
        """Test high-pressure steam system (600 psig)."""
        # 600 psig = 614.7 psia
        pressure = 614.7
        temp = saturation_temperature(pressure)

        # Should be around 486-490°F
        assert temp > 480 and temp < 495

        # Verify roundtrip
        pressure_check = saturation_pressure(temp)
        assert abs(pressure_check - pressure) / pressure < 0.01

    def test_vacuum_deaerator(self):
        """Test vacuum deaerator conditions (5 psia)."""
        pressure = 5.0
        temp = saturation_temperature(pressure)

        # Should be around 162°F
        assert temp > 160 and temp < 165

        # Verify roundtrip
        pressure_check = saturation_pressure(temp)
        assert abs(pressure_check - pressure) / pressure < 0.02

    def test_condenser_operation(self):
        """Test steam condenser at 1.5 psia (typical vacuum)."""
        pressure = 1.5
        temp = saturation_temperature(pressure)

        # Should be around 115-120°F
        assert temp > 110 and temp < 125

        # Verify roundtrip
        pressure_check = saturation_pressure(temp)
        assert abs(pressure_check - pressure) / pressure < 0.02

    def test_moderate_pressure_range(self):
        """Test moderate pressure range (50-150 psia) used in process heating."""
        pressures = [50, 75, 100, 125, 150]

        for p in pressures:
            temp = saturation_temperature(p)

            # Temperatures should be in reasonable process range
            assert temp > 250 and temp < 400

            # Verify accuracy with roundtrip
            p_check = saturation_pressure(temp)
            assert abs(p_check - p) / p < 0.01

    def test_flash_steam_calculation(self):
        """Test flash steam conditions when pressure drops."""
        # Start with 150 psia condensate
        initial_pressure = 150.0
        initial_temp = saturation_temperature(initial_pressure)

        # Flash to 50 psia
        flash_pressure = 50.0
        flash_temp = saturation_temperature(flash_pressure)

        # Flash temperature should be lower
        assert flash_temp < initial_temp

        # Temperature drop should be significant (>70°F)
        temp_drop = initial_temp - flash_temp
        assert temp_drop > 70 and temp_drop < 100


class TestWaterDensity:
    """Tests for water_density() function."""

    def test_standard_conditions(self):
        """Test at standard conditions (60°F, 1 atm)."""
        result = water_density(60.0)
        # Reference: 62.366 lb/ft³ at 60°F
        assert abs(result - 62.366) < 0.1

    def test_freezing_point(self):
        """Test at freezing point (32°F)."""
        result = water_density(32.0)
        # Reference: 62.428 lb/ft³ at 32°F (maximum density ~39°F)
        assert abs(result - 62.428) < 0.1

    def test_boiling_point(self):
        """Test at atmospheric boiling point (212°F)."""
        result = water_density(212.0)
        # Reference: 59.83 lb/ft³ at 212°F
        assert abs(result - 59.83) < 0.2

    def test_hot_water_200f(self):
        """Test at 200°F (common process temperature)."""
        result = water_density(200.0)
        # Reference: ~60.1 lb/ft³ at 200°F
        assert result > 59.8 and result < 60.5

    def test_room_temperature(self):
        """Test at room temperature (68°F)."""
        result = water_density(68.0)
        # Should be very close to 62.3 lb/ft³
        assert abs(result - 62.3) < 0.2

    def test_density_decreases_with_temperature(self):
        """Test that density decreases with increasing temperature."""
        temps = [40, 80, 120, 160, 200]
        densities = [water_density(t) for t in temps]

        # Density should decrease monotonically (after ~39°F)
        for i in range(1, len(densities)):
            assert densities[i] < densities[i-1]

    def test_pressure_effect(self):
        """Test that pressure increases density (compressibility)."""
        temp = 100.0

        density_low = water_density(temp, 14.7)
        density_high = water_density(temp, 1000.0)

        # Higher pressure should give slightly higher density
        assert density_high > density_low

        # But effect should be small (<1% for reasonable pressures)
        percent_change = (density_high - density_low) / density_low * 100
        assert percent_change < 1.0

    def test_below_freezing_error(self):
        """Test error for temperature below freezing."""
        with pytest.raises(ValueError, match="below freezing point"):
            water_density(20.0)

    def test_above_range_error(self):
        """Test error for temperature above valid range."""
        with pytest.raises(ValueError, match="exceeds valid range"):
            water_density(450.0)

    def test_negative_pressure_error(self):
        """Test error for negative pressure."""
        with pytest.raises(ValueError, match="must be positive"):
            water_density(100.0, -5.0)

    def test_zero_pressure_error(self):
        """Test error for zero pressure."""
        with pytest.raises(ValueError, match="must be positive"):
            water_density(100.0, 0.0)


class TestWaterViscosity:
    """Tests for water_viscosity() function."""

    def test_room_temperature(self):
        """Test at room temperature (68°F)."""
        result = water_viscosity(68.0)
        # Reference: ~0.000668 lb/(ft·s) = 1.002 cP
        assert abs(result - 0.000668) < 0.00005

    def test_boiling_point(self):
        """Test at atmospheric boiling point (212°F)."""
        result = water_viscosity(212.0)
        # Reference: ~0.000196 lb/(ft·s) = 0.294 cP
        assert abs(result - 0.000196) < 0.00002

    def test_cold_water(self):
        """Test at 40°F (cold water)."""
        result = water_viscosity(40.0)
        # Reference: ~0.00108 lb/(ft·s) = 1.62 cP
        assert abs(result - 0.00108) < 0.0001

    def test_hot_water_200f(self):
        """Test at 200°F."""
        result = water_viscosity(200.0)
        # Reference: ~0.000204 lb/(ft·s) = 0.305 cP
        assert abs(result - 0.000204) < 0.00002

    def test_viscosity_decreases_with_temperature(self):
        """Test that viscosity decreases with increasing temperature."""
        temps = [40, 80, 120, 160, 200, 300]
        viscosities = [water_viscosity(t) for t in temps]

        # Viscosity should decrease monotonically
        for i in range(1, len(viscosities)):
            assert viscosities[i] < viscosities[i-1]

    def test_exponential_decrease(self):
        """Test that viscosity decreases exponentially (not linearly)."""
        # Viscosity change should be larger at low temperatures
        visc_40 = water_viscosity(40.0)
        visc_80 = water_viscosity(80.0)
        visc_200 = water_viscosity(200.0)
        visc_240 = water_viscosity(240.0)

        # Change from 40-80°F should be larger than 200-240°F
        change_low_temp = visc_40 - visc_80
        change_high_temp = visc_200 - visc_240

        assert change_low_temp > change_high_temp

    def test_below_freezing_error(self):
        """Test error for temperature below freezing."""
        with pytest.raises(ValueError, match="below freezing point"):
            water_viscosity(20.0)

    def test_above_range_error(self):
        """Test error for temperature above valid range."""
        with pytest.raises(ValueError, match="exceeds valid range"):
            water_viscosity(450.0)


class TestDensityViscosityVBACompatibility:
    """Tests for VBA-compatible wrapper functions."""

    def test_water_density_wrapper(self):
        """Test WaterDensity() VBA wrapper."""
        result_vba = WaterDensity(60.0)
        result_python = water_density(60.0)

        assert result_vba == result_python
        assert abs(result_vba - 62.366) < 0.1

    def test_water_density_wrapper_with_pressure(self):
        """Test WaterDensity() with pressure parameter."""
        result_vba = WaterDensity(100.0, 100.0)
        result_python = water_density(100.0, 100.0)

        assert result_vba == result_python

    def test_water_viscosity_wrapper(self):
        """Test WaterViscosity() VBA wrapper."""
        result_vba = WaterViscosity(68.0)
        result_python = water_viscosity(68.0)

        assert result_vba == result_python
        assert abs(result_vba - 0.000668) < 0.00005


class TestDensityViscosityIntegration:
    """Integration tests for density and viscosity with realistic scenarios."""

    def test_reynolds_number_calculation(self):
        """Test using density and viscosity to calculate Reynolds number."""
        # Typical pipe flow scenario
        temp = 100.0
        velocity = 5.0  # ft/s
        diameter = 0.5  # ft (6 inch pipe)

        density = water_density(temp)
        viscosity = water_viscosity(temp)

        # Reynolds number: Re = ρ * V * D / μ
        reynolds = density * velocity * diameter / viscosity

        # Should be in turbulent range (> 4000) for water at these conditions
        assert reynolds > 100000
        assert reynolds < 500000

    def test_kinematic_viscosity(self):
        """Test calculation of kinematic viscosity (ν = μ/ρ)."""
        temp = 68.0

        density = water_density(temp)
        dynamic_visc = water_viscosity(temp)

        kinematic_visc = dynamic_visc / density

        # At 68°F: ν ≈ 1.08e-5 ft²/s
        assert abs(kinematic_visc - 1.08e-5) < 1.0e-6

    def test_pressure_drop_parameters(self):
        """Test parameters needed for Darcy-Weisbach pressure drop."""
        # Hot water heating system
        temp = 180.0
        pressure = 30.0  # psig = 44.7 psia

        density = water_density(temp, 44.7)
        viscosity = water_viscosity(temp)

        # Verify reasonable values for hot water
        assert density > 59 and density < 61  # Hot water less dense
        assert viscosity < 0.0003  # Hot water less viscous

    def test_temperature_effect_on_flow(self):
        """Test how temperature affects flow properties."""
        cold_temp = 50.0
        hot_temp = 200.0

        # Cold water properties
        rho_cold = water_density(cold_temp)
        mu_cold = water_viscosity(cold_temp)
        nu_cold = mu_cold / rho_cold

        # Hot water properties
        rho_hot = water_density(hot_temp)
        mu_hot = water_viscosity(hot_temp)
        nu_hot = mu_hot / rho_hot

        # Hot water: lower density, lower viscosity, lower kinematic viscosity
        assert rho_hot < rho_cold
        assert mu_hot < mu_cold
        assert nu_hot < nu_cold

        # Kinematic viscosity should decrease significantly
        ratio = nu_cold / nu_hot
        assert ratio > 3.0  # At least 3x decrease

    def test_high_pressure_boiler_feedwater(self):
        """Test properties for high-pressure boiler feedwater."""
        # Typical deaerator outlet conditions
        temp = 227.0  # Saturated at ~15 psia
        pressure = 200.0  # Pump discharge

        density = water_density(temp, pressure)
        viscosity = water_viscosity(temp)

        # Verify reasonable values
        assert density > 58 and density < 60
        assert viscosity > 0.00015 and viscosity < 0.0002


class TestWaterSpecificHeat:
    """Tests for water_specific_heat() function."""

    def test_standard_conditions(self):
        """Test at standard conditions (60°F)."""
        result = water_specific_heat(60.0)
        # Reference: ~0.9988-1.005 BTU/(lb·°F) at 60°F (small variations in literature)
        assert result > 0.998 and result < 1.010

    def test_freezing_point(self):
        """Test at freezing point (32°F)."""
        result = water_specific_heat(32.0)
        # Reference: ~1.0074 BTU/(lb·°F) at 32°F
        assert abs(result - 1.0074) < 0.005

    def test_room_temperature(self):
        """Test at room temperature (68°F)."""
        result = water_specific_heat(68.0)
        # Should be very close to 1.0 BTU/(lb·°F)
        assert abs(result - 1.0) < 0.005

    def test_boiling_point(self):
        """Test at atmospheric boiling point (212°F)."""
        result = water_specific_heat(212.0)
        # Reference: ~1.007 BTU/(lb·°F) at 212°F
        assert abs(result - 1.007) < 0.005

    def test_hot_water_200f(self):
        """Test at 200°F."""
        result = water_specific_heat(200.0)
        # Should be around 1.006-1.008
        assert result > 1.002 and result < 1.012

    def test_minimum_cp(self):
        """Test that cp variation is small over typical range."""
        # cp varies by less than 1% over typical water temperature range
        cp_60 = water_specific_heat(60.0)
        cp_95 = water_specific_heat(95.0)
        cp_150 = water_specific_heat(150.0)

        # All values should be close to 1.0
        assert cp_60 > 0.998 and cp_60 < 1.012
        assert cp_95 > 0.998 and cp_95 < 1.012
        assert cp_150 > 0.998 and cp_150 < 1.012

    def test_cp_relatively_constant(self):
        """Test that cp variation is small over typical range."""
        cp_50 = water_specific_heat(50.0)
        cp_150 = water_specific_heat(150.0)

        # Variation should be less than 2%
        variation = abs(cp_150 - cp_50) / cp_50
        assert variation < 0.02

    def test_below_freezing_error(self):
        """Test error for temperature below freezing."""
        with pytest.raises(ValueError, match="below freezing point"):
            water_specific_heat(20.0)

    def test_above_range_error(self):
        """Test error for temperature above valid range."""
        with pytest.raises(ValueError, match="exceeds valid range"):
            water_specific_heat(450.0)


class TestWaterThermalConductivity:
    """Tests for water_thermal_conductivity() function."""

    def test_standard_conditions(self):
        """Test at standard conditions (68°F)."""
        result = water_thermal_conductivity(68.0)
        # Reference: ~0.345 BTU/(hr·ft·°F) at 68°F
        assert abs(result - 0.345) < 0.005

    def test_freezing_point(self):
        """Test at freezing point (32°F)."""
        result = water_thermal_conductivity(32.0)
        # Reference: ~0.319-0.327 BTU/(hr·ft·°F) at 32°F
        assert result > 0.315 and result < 0.330

    def test_boiling_point(self):
        """Test at atmospheric boiling point (212°F)."""
        result = water_thermal_conductivity(212.0)
        # Reference: ~0.390-0.395 BTU/(hr·ft·°F) at 212°F
        assert result > 0.385 and result < 0.400

    def test_hot_water_200f(self):
        """Test at 200°F."""
        result = water_thermal_conductivity(200.0)
        # Should be around 0.390-0.395
        assert result > 0.385 and result < 0.400

    def test_cold_water_40f(self):
        """Test at 40°F."""
        result = water_thermal_conductivity(40.0)
        # Should be around 0.325
        assert abs(result - 0.325) < 0.01

    def test_conductivity_increases_with_temperature(self):
        """Test that thermal conductivity increases with temperature (unusual for liquids)."""
        temps = [40, 80, 120, 160, 200, 300]
        conductivities = [water_thermal_conductivity(t) for t in temps]

        # Conductivity should increase monotonically
        for i in range(1, len(conductivities)):
            assert conductivities[i] > conductivities[i-1]

    def test_relative_increase(self):
        """Test magnitude of conductivity increase."""
        k_40 = water_thermal_conductivity(40.0)
        k_200 = water_thermal_conductivity(200.0)

        # Should increase by about 20% from 40°F to 200°F
        increase_percent = (k_200 - k_40) / k_40 * 100
        assert increase_percent > 15 and increase_percent < 25

    def test_below_freezing_error(self):
        """Test error for temperature below freezing."""
        with pytest.raises(ValueError, match="below freezing point"):
            water_thermal_conductivity(20.0)

    def test_above_range_error(self):
        """Test error for temperature above valid range."""
        with pytest.raises(ValueError, match="exceeds valid range"):
            water_thermal_conductivity(450.0)


class TestThermalPropertiesVBACompatibility:
    """Tests for VBA-compatible wrapper functions."""

    def test_specific_heat_wrapper(self):
        """Test WaterSpecificHeat() VBA wrapper."""
        result_vba = WaterSpecificHeat(60.0)
        result_python = water_specific_heat(60.0)

        assert result_vba == result_python
        assert result_vba > 0.998 and result_vba < 1.010

    def test_thermal_conductivity_wrapper(self):
        """Test WaterThermalConductivity() VBA wrapper."""
        result_vba = WaterThermalConductivity(68.0)
        result_python = water_thermal_conductivity(68.0)

        assert result_vba == result_python
        assert abs(result_vba - 0.345) < 0.005


class TestThermalPropertiesIntegration:
    """Integration tests for thermal properties with realistic scenarios."""

    def test_prandtl_number_calculation(self):
        """Test calculation of Prandtl number (Pr = cp·μ/k)."""
        temp = 100.0

        cp = water_specific_heat(temp)
        mu = water_viscosity(temp)
        k = water_thermal_conductivity(temp)

        # Prandtl number for water
        # Note: Need consistent units
        # cp is in BTU/(lb·°F) = BTU/(lb·°F)
        # mu is in lb/(ft·s)
        # k is in BTU/(hr·ft·°F)

        # Convert mu to lb/(ft·hr) for consistency
        mu_lb_ft_hr = mu * 3600.0  # Convert ft·s to ft·hr

        # Pr = cp·μ/k (dimensionless)
        prandtl = (cp * mu_lb_ft_hr) / k

        # Water Prandtl number at 100°F should be around 4-6
        assert prandtl > 3.5 and prandtl < 7.0

    def test_thermal_diffusivity(self):
        """Test calculation of thermal diffusivity (α = k/(ρ·cp))."""
        temp = 68.0

        k = water_thermal_conductivity(temp)
        rho = water_density(temp)
        cp = water_specific_heat(temp)

        # Thermal diffusivity α = k/(ρ·cp)
        # k is in BTU/(hr·ft·°F)
        # ρ is in lb/ft³
        # cp is in BTU/(lb·°F)
        # α should be in ft²/hr

        alpha = k / (rho * cp)

        # Water thermal diffusivity at 68°F ~ 0.0058 ft²/hr
        assert alpha > 0.005 and alpha < 0.007

    def test_heat_duty_calculation(self):
        """Test heat duty calculation with specific heat."""
        # Heating 1000 lb/hr of water from 60°F to 180°F
        mass_flow = 1000.0  # lb/hr
        temp_in = 60.0
        temp_out = 180.0

        # Use average temperature for cp
        temp_avg = (temp_in + temp_out) / 2.0
        cp = water_specific_heat(temp_avg)

        # Heat duty: Q = m·cp·ΔT
        delta_t = temp_out - temp_in
        heat_duty = mass_flow * cp * delta_t

        # Should be around 120,000 BTU/hr
        # (1000 lb/hr * 1.0 BTU/(lb·°F) * 120 °F = 120,000 BTU/hr)
        assert heat_duty > 118000 and heat_duty < 122000

    def test_heat_exchanger_ua(self):
        """Test overall heat transfer coefficient calculation."""
        # Simplified heat transfer scenario
        temp = 150.0

        k = water_thermal_conductivity(temp)

        # Typical film coefficient calculation uses k
        # For natural convection: h ~ k / L
        # Characteristic length L = 1 ft
        L = 1.0
        h_natural = k / L  # Very simplified

        # Should be in reasonable range for natural convection
        # Actual h would be higher with forced convection
        assert h_natural > 0.3 and h_natural < 0.5

    def test_properties_variation_over_range(self):
        """Test how all thermal properties vary together."""
        temp_cold = 50.0
        temp_hot = 200.0

        # Cold water properties
        cp_cold = water_specific_heat(temp_cold)
        k_cold = water_thermal_conductivity(temp_cold)

        # Hot water properties
        cp_hot = water_specific_heat(temp_hot)
        k_hot = water_thermal_conductivity(temp_hot)

        # Both cp and k increase with temperature
        assert cp_hot > cp_cold  # cp increases (slightly)
        assert k_hot > k_cold    # k increases significantly

        # But cp increases less than k
        cp_ratio = cp_hot / cp_cold
        k_ratio = k_hot / k_cold

        assert k_ratio > cp_ratio  # k increases more than cp

    def test_boiler_feedwater_thermal_properties(self):
        """Test thermal properties for boiler feedwater."""
        # Typical deaerator outlet: 227°F
        temp = 227.0

        cp = water_specific_heat(temp)
        k = water_thermal_conductivity(temp)

        # Verify reasonable values for high-temperature water
        assert cp > 1.002 and cp < 1.015  # Slightly higher than 1.0
        assert k > 0.390 and k < 0.410    # Higher than room temp


# ============================================================================
# STEAM ENTHALPY AND QUALITY TESTS
# ============================================================================


class TestSteamEnthalpy:
    """Tests for steam_enthalpy() function."""

    def test_saturated_liquid_low_pressure(self):
        """Test hf at 14.7 psia (212°F)."""
        result = steam_enthalpy(212.0, 14.7, 0.0)
        assert abs(result - 180.0) < 2.0  # Within 1% of 180 BTU/lb

    def test_saturated_vapor_low_pressure(self):
        """Test hg at 14.7 psia (212°F)."""
        result = steam_enthalpy(212.0, 14.7, 1.0)
        assert abs(result - 1150.0) < 10.0  # Within 1% of 1150 BTU/lb

    def test_two_phase_mixture_50_percent(self):
        """Test enthalpy of 50% quality mixture."""
        result = steam_enthalpy(212.0, 14.7, 0.5)
        # Should be average of hf and hg
        expected = (180.0 + 1150.0) / 2.0
        assert abs(result - expected) < 10.0

    def test_saturated_liquid_100_psia(self):
        """Test hf at 100 psia (328°F)."""
        result = steam_enthalpy(328.0, 100.0, 0.0)
        assert abs(result - 298.0) < 3.0  # Within 1% of 298 BTU/lb

    def test_saturated_vapor_100_psia(self):
        """Test hg at 100 psia (328°F)."""
        result = steam_enthalpy(328.0, 100.0, 1.0)
        assert abs(result - 1187.0) < 12.0  # Within 1% of 1187 BTU/lb

    def test_saturated_liquid_200_psia(self):
        """Test hf at 200 psia (382°F)."""
        result = steam_enthalpy(382.0, 200.0, 0.0)
        assert abs(result - 355.0) < 5.0  # Within 1.5% of 355 BTU/lb

    def test_saturated_vapor_200_psia(self):
        """Test hg at 200 psia (382°F)."""
        result = steam_enthalpy(382.0, 200.0, 1.0)
        assert abs(result - 1198.0) < 12.0  # Within 1% of 1198 BTU/lb

    def test_quality_variation(self):
        """Test that enthalpy increases with quality."""
        pressure = 50.0
        temp = saturation_temperature(pressure)

        qualities = [0.0, 0.25, 0.5, 0.75, 1.0]
        enthalpies = [steam_enthalpy(temp, pressure, q) for q in qualities]

        # Verify monotonically increasing
        for i in range(1, len(enthalpies)):
            assert enthalpies[i] > enthalpies[i-1]

    def test_subcooled_liquid(self):
        """Test subcooled liquid (T < Tsat)."""
        # At 14.7 psia, Tsat = 212°F
        # Test at 150°F (62°F subcooling)
        result = steam_enthalpy(150.0, 14.7, 0.0)

        # Should be less than hf at saturation
        hf_sat = steam_enthalpy(212.0, 14.7, 0.0)
        assert result < hf_sat

        # Typical value around 118 BTU/lb
        assert result > 100.0 and result < 130.0

    def test_superheated_vapor(self):
        """Test superheated vapor (T > Tsat)."""
        # At 14.7 psia, Tsat = 212°F
        # Test at 300°F (88°F superheat)
        result = steam_enthalpy(300.0, 14.7, 1.0)

        # Should be greater than hg at saturation
        hg_sat = steam_enthalpy(212.0, 14.7, 1.0)
        assert result > hg_sat

        # Typical value around 1195 BTU/lb
        assert result > 1190.0 and result < 1210.0

    def test_error_temperature_too_low(self):
        """Test error for temperature below 32°F."""
        with pytest.raises(ValueError, match="below freezing"):
            steam_enthalpy(20.0, 14.7, 1.0)

    def test_error_temperature_too_high(self):
        """Test error for temperature above 700°F."""
        with pytest.raises(ValueError, match="above valid range"):
            steam_enthalpy(750.0, 100.0, 1.0)

    def test_error_pressure_zero(self):
        """Test error for zero pressure."""
        with pytest.raises(ValueError, match="Pressure"):
            steam_enthalpy(212.0, 0.0, 1.0)

    def test_error_pressure_negative(self):
        """Test error for negative pressure."""
        with pytest.raises(ValueError, match="Pressure"):
            steam_enthalpy(212.0, -10.0, 1.0)

    def test_error_pressure_too_high(self):
        """Test error for pressure above 3000 psia."""
        with pytest.raises(ValueError, match="Pressure"):
            steam_enthalpy(500.0, 3500.0, 1.0)

    def test_error_quality_negative(self):
        """Test error for quality below 0."""
        with pytest.raises(ValueError, match="Quality"):
            steam_enthalpy(212.0, 14.7, -0.1)

    def test_error_quality_above_one(self):
        """Test error for quality above 1."""
        with pytest.raises(ValueError, match="Quality"):
            steam_enthalpy(212.0, 14.7, 1.5)


class TestSteamQuality:
    """Tests for steam_quality() function."""

    def test_saturated_liquid(self):
        """Test that hf gives quality = 0."""
        hf = steam_enthalpy(212.0, 14.7, 0.0)
        quality = steam_quality(hf, 14.7)
        assert abs(quality - 0.0) < 0.01

    def test_saturated_vapor(self):
        """Test that hg gives quality = 1."""
        hg = steam_enthalpy(212.0, 14.7, 1.0)
        quality = steam_quality(hg, 14.7)
        assert abs(quality - 1.0) < 0.01

    def test_50_percent_quality(self):
        """Test 50% quality mixture."""
        h_50 = steam_enthalpy(212.0, 14.7, 0.5)
        quality = steam_quality(h_50, 14.7)
        assert abs(quality - 0.5) < 0.02

    def test_25_percent_quality(self):
        """Test 25% quality mixture."""
        h_25 = steam_enthalpy(212.0, 14.7, 0.25)
        quality = steam_quality(h_25, 14.7)
        assert abs(quality - 0.25) < 0.02

    def test_75_percent_quality(self):
        """Test 75% quality mixture."""
        h_75 = steam_enthalpy(212.0, 14.7, 0.75)
        quality = steam_quality(h_75, 14.7)
        assert abs(quality - 0.75) < 0.02

    def test_quality_at_100_psia(self):
        """Test quality calculation at higher pressure."""
        h_60 = steam_enthalpy(328.0, 100.0, 0.6)
        quality = steam_quality(h_60, 100.0)
        assert abs(quality - 0.6) < 0.02

    def test_subcooled_returns_negative(self):
        """Test that subcooled liquid returns quality < 0."""
        # Enthalpy of subcooled water at 150°F
        h_subcool = 118.0  # Less than hf at 14.7 psia
        quality = steam_quality(h_subcool, 14.7)
        assert quality < 0.0

    def test_superheated_returns_above_one(self):
        """Test that superheated vapor returns quality > 1."""
        # Enthalpy of superheated steam at 300°F
        h_superheat = 1200.0  # More than hg at 14.7 psia
        quality = steam_quality(h_superheat, 14.7)
        assert quality > 1.0

    def test_roundtrip_conversion(self):
        """Test h -> quality -> h roundtrip."""
        pressure = 50.0
        temp = saturation_temperature(pressure)
        quality_original = 0.65

        # Calculate enthalpy from quality
        h = steam_enthalpy(temp, pressure, quality_original)

        # Calculate quality from enthalpy
        quality_calc = steam_quality(h, pressure)

        # Should match original
        assert abs(quality_calc - quality_original) < 0.01

    def test_error_pressure_zero(self):
        """Test error for zero pressure."""
        with pytest.raises(ValueError, match="Pressure"):
            steam_quality(665.0, 0.0)

    def test_error_pressure_negative(self):
        """Test error for negative pressure."""
        with pytest.raises(ValueError, match="Pressure"):
            steam_quality(665.0, -10.0)

    def test_error_pressure_too_high(self):
        """Test error for pressure above 3000 psia."""
        with pytest.raises(ValueError, match="Pressure"):
            steam_quality(665.0, 3500.0)


class TestSteamPropertiesIntegration:
    """Integration tests for steam properties."""

    def test_flash_steam_calculation(self):
        """Test flash steam scenario: high-pressure condensate flashing."""
        # High-pressure condensate at 200 psia
        h_condensate = steam_enthalpy(382.0, 200.0, 0.0)

        # Flashes to 14.7 psia (enthalpy remains constant)
        quality_flash = steam_quality(h_condensate, 14.7)

        # Should produce significant flash steam
        assert quality_flash > 0.15 and quality_flash < 0.25

    def test_turbine_expansion(self):
        """Test steam turbine expansion calculation."""
        # Inlet: superheated steam at 200 psia, 500°F
        h_inlet = steam_enthalpy(500.0, 200.0, 1.0)

        # Exhaust at 14.7 psia (isentropic expansion ~85% efficient)
        # Simplified: assume some enthalpy drop
        h_exhaust = h_inlet - 150.0  # BTU/lb

        quality_exhaust = steam_quality(h_exhaust, 14.7)

        # Should be wet steam (quality < 1)
        assert quality_exhaust > 0.85 and quality_exhaust < 1.0

    def test_boiler_heat_duty(self):
        """Test boiler heat duty calculation."""
        # Feedwater at 227°F, 200 psia
        h_feedwater = steam_enthalpy(227.0, 200.0, 0.0)

        # Steam output at 200 psia, saturated
        h_steam = steam_enthalpy(382.0, 200.0, 1.0)

        # Heat duty per lb
        q = h_steam - h_feedwater

        # Should be reasonable for feedwater to steam conversion
        # hfg at 200 psia is ~843 BTU/lb, plus sensible heat
        assert q > 950.0 and q < 1000.0  # BTU/lb

    def test_vba_compatibility_steam_enthalpy(self):
        """Test VBA wrapper for steam_enthalpy."""
        result = SteamEnthalpy(212.0, 14.7, 1.0)
        expected = steam_enthalpy(212.0, 14.7, 1.0)
        assert abs(result - expected) < 0.01

    def test_vba_compatibility_steam_quality(self):
        """Test VBA wrapper for steam_quality."""
        result = SteamQuality(665.0, 14.7)
        expected = steam_quality(665.0, 14.7)
        assert abs(result - expected) < 0.001
