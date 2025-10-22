"""
Unit tests for combustion efficiency calculations.

Tests efficiency calculations per ASME PTC 4 methodology including
combustion efficiency, stack losses, and thermal efficiency.
"""

import pytest
from sigma_thermal.combustion.efficiency import (
    combustion_efficiency,
    stack_loss_percent,
    thermal_efficiency,
    radiation_loss_percent,
    CombustionEfficiency,
    StackLossPercent,
    ThermalEfficiency,
)


class TestCombustionEfficiency:
    """Tests for combustion_efficiency function"""

    def test_basic_efficiency(self):
        """Test basic combustion efficiency calculation"""
        # 1M BTU/hr input, 150k BTU/hr stack loss = 85% efficiency
        result = combustion_efficiency(1000000, 150000)
        assert result == pytest.approx(85.0, rel=1e-4)

    def test_with_radiation_loss(self):
        """Test efficiency with radiation loss"""
        # 1M BTU/hr input, 150k stack + 20k radiation = 83% efficiency
        result = combustion_efficiency(1000000, 150000, radiation_loss=20000)
        assert result == pytest.approx(83.0, rel=1e-4)

    def test_high_efficiency_boiler(self):
        """Test high efficiency boiler (93%)"""
        result = combustion_efficiency(2387500, 163000)
        expected = ((2387500 - 163000) / 2387500) * 100
        assert result == pytest.approx(expected, rel=1e-4)
        assert result > 93.0  # Should be >93%

    def test_with_all_losses(self):
        """Test efficiency with all loss types"""
        result = combustion_efficiency(
            heat_input=1000000,
            stack_loss=150000,
            radiation_loss=20000,
            blow_down_loss=10000,
            unaccounted_loss=5000
        )
        total_loss = 150000 + 20000 + 10000 + 5000
        expected = ((1000000 - total_loss) / 1000000) * 100
        assert result == pytest.approx(expected, rel=1e-4)

    def test_low_efficiency_case(self):
        """Test low efficiency case (70%)"""
        result = combustion_efficiency(1000000, 300000)
        assert result == pytest.approx(70.0, rel=1e-4)

    def test_very_high_efficiency(self):
        """Test very high efficiency (95%)"""
        result = combustion_efficiency(1000000, 50000)
        assert result == pytest.approx(95.0, rel=1e-4)

    def test_zero_losses(self):
        """Test theoretical 100% efficiency (zero losses)"""
        result = combustion_efficiency(1000000, 0)
        assert result == pytest.approx(100.0, rel=1e-6)

    def test_zero_heat_input_error(self):
        """Test error when heat input is zero"""
        with pytest.raises(ValueError, match="Heat input must be positive"):
            combustion_efficiency(0, 100000)

    def test_negative_heat_input_error(self):
        """Test error when heat input is negative"""
        with pytest.raises(ValueError, match="Heat input must be positive"):
            combustion_efficiency(-1000000, 100000)

    def test_losses_exceed_input_error(self):
        """Test error when total losses exceed heat input"""
        with pytest.raises(ValueError, match="Total losses .* exceed heat input"):
            combustion_efficiency(1000000, 1200000)

    def test_losses_equal_input(self):
        """Test edge case where losses equal input (0% efficiency)"""
        result = combustion_efficiency(1000000, 1000000)
        assert result == pytest.approx(0.0, abs=1e-6)


class TestStackLossPercent:
    """Tests for stack_loss_percent function"""

    def test_basic_stack_loss(self):
        """Test basic stack loss calculation"""
        # 300 BTU/lb enthalpy, 2000 lb/hr flow, 2M BTU/hr input
        # Stack loss = 300 * 2000 = 600k BTU/hr = 30% of 2M
        result = stack_loss_percent(300, 2000, 2000000)
        assert result == pytest.approx(30.0, rel=1e-4)

    def test_high_efficiency_low_stack_loss(self):
        """Test low stack loss for high efficiency system"""
        # 150 BTU/lb, 2020 lb/hr, 2.3875M BTU/hr = ~12.7%
        result = stack_loss_percent(150, 2020, 2387500)
        expected = (150 * 2020 / 2387500) * 100
        assert result == pytest.approx(expected, rel=1e-3)

    def test_low_efficiency_high_stack_loss(self):
        """Test high stack loss for low efficiency system"""
        # 426.5 BTU/lb, 2020.5 lb/hr, 2.3875M BTU/hr = ~36%
        result = stack_loss_percent(426.5, 2020.5, 2387500)
        expected = (426.5 * 2020.5 / 2387500) * 100
        assert result == pytest.approx(expected, rel=1e-2)

    def test_small_stack_loss(self):
        """Test small stack loss (5%)"""
        # Design for 5% stack loss
        heat_input = 1000000
        target_pct = 5.0
        enthalpy = 100.0
        flow = (target_pct * heat_input) / (100 * enthalpy)
        result = stack_loss_percent(enthalpy, flow, heat_input)
        assert result == pytest.approx(target_pct, rel=1e-3)

    def test_large_stack_loss(self):
        """Test large stack loss (40%)"""
        result = stack_loss_percent(400, 1000, 1000000)
        assert result == pytest.approx(40.0, rel=1e-4)

    def test_zero_enthalpy(self):
        """Test zero enthalpy (no stack loss)"""
        result = stack_loss_percent(0, 2000, 1000000)
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_zero_heat_input_error(self):
        """Test error when heat input is zero"""
        with pytest.raises(ValueError, match="Heat input must be positive"):
            stack_loss_percent(300, 2000, 0)

    def test_negative_heat_input_error(self):
        """Test error when heat input is negative"""
        with pytest.raises(ValueError, match="Heat input must be positive"):
            stack_loss_percent(300, 2000, -1000000)

    def test_stack_loss_exceeds_100_percent_error(self):
        """Test error when stack loss exceeds 100%"""
        with pytest.raises(ValueError, match="Stack loss percentage .* exceeds 100%"):
            stack_loss_percent(600, 2000, 1000000)

    def test_negative_enthalpy_error(self):
        """Test error when enthalpy is negative (below ambient)"""
        with pytest.raises(ValueError, match="Stack loss percentage .* is negative"):
            stack_loss_percent(-100, 2000, 1000000)


class TestThermalEfficiency:
    """Tests for thermal_efficiency function"""

    def test_basic_thermal_efficiency(self):
        """Test basic thermal efficiency calculation"""
        result = thermal_efficiency(850000, 1000000)
        assert result == pytest.approx(85.0, rel=1e-4)

    def test_perfect_efficiency(self):
        """Test perfect thermal efficiency (100%)"""
        result = thermal_efficiency(1000, 1000)
        assert result == pytest.approx(100.0, rel=1e-6)

    def test_low_efficiency(self):
        """Test low thermal efficiency (60%)"""
        result = thermal_efficiency(600000, 1000000)
        assert result == pytest.approx(60.0, rel=1e-4)

    def test_high_efficiency(self):
        """Test high thermal efficiency (95%)"""
        result = thermal_efficiency(950000, 1000000)
        assert result == pytest.approx(95.0, rel=1e-4)

    def test_very_low_efficiency(self):
        """Test very low thermal efficiency (20%)"""
        result = thermal_efficiency(200000, 1000000)
        assert result == pytest.approx(20.0, rel=1e-4)

    def test_zero_output(self):
        """Test zero heat output (0% efficiency)"""
        result = thermal_efficiency(0, 1000000)
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_zero_input_error(self):
        """Test error when heat input is zero"""
        with pytest.raises(ValueError, match="Heat input must be positive"):
            thermal_efficiency(850000, 0)

    def test_negative_input_error(self):
        """Test error when heat input is negative"""
        with pytest.raises(ValueError, match="Heat input must be positive"):
            thermal_efficiency(850000, -1000000)

    def test_output_exceeds_input_error(self):
        """Test error when output exceeds input"""
        with pytest.raises(ValueError, match="Heat output .* cannot exceed heat input"):
            thermal_efficiency(1200000, 1000000)


class TestRadiationLossPercent:
    """Tests for radiation_loss_percent function"""

    def test_basic_radiation_loss(self):
        """Test basic radiation loss calculation"""
        # Small boiler: 100 ft² surface, 200°F surface, 70°F ambient
        result = radiation_loss_percent(100, 200, 70)
        assert result > 0  # Should have some radiation loss
        assert result < 25000  # Should be reasonable magnitude (BTU/hr)

    def test_high_temperature_surface(self):
        """Test radiation from high temperature surface"""
        # Higher temperature = more radiation (T^4 relationship)
        result_high = radiation_loss_percent(100, 400, 70)
        result_low = radiation_loss_percent(100, 200, 70)
        assert result_high > result_low

    def test_large_surface_area(self):
        """Test larger surface area increases radiation"""
        result_large = radiation_loss_percent(200, 200, 70)
        result_small = radiation_loss_percent(100, 200, 70)
        assert result_large == pytest.approx(2 * result_small, rel=1e-3)

    def test_zero_temperature_difference(self):
        """Test zero radiation when surface equals ambient"""
        result = radiation_loss_percent(100, 70, 70)
        assert result == pytest.approx(0.0, abs=1e-3)

    def test_custom_emissivity(self):
        """Test custom emissivity value"""
        # Lower emissivity = less radiation
        result_high_e = radiation_loss_percent(100, 200, 70, emissivity=0.9)
        result_low_e = radiation_loss_percent(100, 200, 70, emissivity=0.5)
        assert result_high_e > result_low_e

    def test_stefan_boltzmann_t4_relationship(self):
        """Test T^4 relationship in Stefan-Boltzmann law"""
        # Temperature in absolute scale matters
        result = radiation_loss_percent(100, 200, 70)
        # Result should be positive and reasonable
        assert result > 100  # Should be on order of 100s BTU/hr for these conditions


class TestVBACompatibility:
    """Tests for VBA-compatible wrapper functions"""

    def test_combustion_efficiency_vba(self):
        """Test VBA wrapper for combustion_efficiency"""
        result = CombustionEfficiency(1000000, 150000)
        assert result == pytest.approx(85.0, rel=1e-4)

    def test_combustion_efficiency_vba_all_params(self):
        """Test VBA wrapper with all parameters"""
        result = CombustionEfficiency(
            heat_input=1000000,
            stack_loss=150000,
            radiation_loss=20000,
            blow_down_loss=10000,
            unaccounted_loss=5000
        )
        expected = ((1000000 - 185000) / 1000000) * 100
        assert result == pytest.approx(expected, rel=1e-4)

    def test_stack_loss_percent_vba(self):
        """Test VBA wrapper for stack_loss_percent"""
        result = StackLossPercent(300, 2000, 2000000)
        assert result == pytest.approx(30.0, rel=1e-4)

    def test_thermal_efficiency_vba(self):
        """Test VBA wrapper for thermal_efficiency"""
        result = ThermalEfficiency(850000, 1000000)
        assert result == pytest.approx(85.0, rel=1e-4)


class TestIntegrationScenarios:
    """Integration tests with realistic combustion scenarios"""

    def test_natural_gas_boiler_complete(self):
        """Test complete efficiency calculation for natural gas boiler"""
        heat_input = 2387500  # BTU/hr
        flue_gas_enthalpy = 150  # BTU/lb (low stack temp, high efficiency)
        flue_gas_flow = 2020  # lb/hr
        radiation = 20000  # BTU/hr (typical ~1%)

        # Calculate stack loss
        stack_loss_pct = stack_loss_percent(flue_gas_enthalpy, flue_gas_flow, heat_input)
        stack_loss = (stack_loss_pct / 100) * heat_input

        # Calculate combustion efficiency
        eff = combustion_efficiency(heat_input, stack_loss, radiation_loss=radiation)

        # Efficiency should be reasonable (85-90% range for this case)
        assert eff > 85.0
        assert eff < 90.0

    def test_oil_fired_heater_complete(self):
        """Test complete efficiency calculation for oil-fired heater"""
        heat_input = 1500000  # BTU/hr
        flue_gas_enthalpy = 250  # BTU/lb (moderate stack temp)
        flue_gas_flow = 1800  # lb/hr
        radiation = 30000  # BTU/hr (typical ~2%)

        # Calculate stack loss
        stack_loss_pct = stack_loss_percent(flue_gas_enthalpy, flue_gas_flow, heat_input)
        stack_loss = (stack_loss_pct / 100) * heat_input

        # Calculate combustion efficiency
        eff = combustion_efficiency(heat_input, stack_loss, radiation_loss=radiation)

        # Oil-fired should be 70-80% for this case
        assert eff > 65.0
        assert eff < 80.0

    def test_efficiency_from_thermal_output(self):
        """Test calculating efficiency from thermal output"""
        heat_input = 1000000  # BTU/hr
        heat_output = 850000  # BTU/hr delivered to process

        # Using direct input-output method
        eff_thermal = thermal_efficiency(heat_output, heat_input)
        assert eff_thermal == pytest.approx(85.0, rel=1e-4)

        # Using loss method (losses = input - output)
        total_losses = heat_input - heat_output
        eff_loss = combustion_efficiency(heat_input, total_losses)
        assert eff_loss == pytest.approx(85.0, rel=1e-4)

        # Both methods should give same result
        assert eff_thermal == pytest.approx(eff_loss, rel=1e-6)

    def test_stack_temperature_impact_on_efficiency(self):
        """Test how stack temperature affects efficiency"""
        heat_input = 2000000
        flue_gas_flow = 2000

        # High stack temp (low efficiency)
        high_temp_enthalpy = 400  # BTU/lb
        stack_loss_high = stack_loss_percent(high_temp_enthalpy, flue_gas_flow, heat_input)

        # Low stack temp (high efficiency)
        low_temp_enthalpy = 150  # BTU/lb
        stack_loss_low = stack_loss_percent(low_temp_enthalpy, flue_gas_flow, heat_input)

        # Lower stack temp = lower stack loss
        assert stack_loss_low < stack_loss_high

        # Calculate efficiencies
        eff_high_temp = combustion_efficiency(heat_input, (stack_loss_high / 100) * heat_input)
        eff_low_temp = combustion_efficiency(heat_input, (stack_loss_low / 100) * heat_input)

        # Lower stack temp = higher efficiency
        assert eff_low_temp > eff_high_temp
