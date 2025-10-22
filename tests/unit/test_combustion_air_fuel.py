"""
Unit tests for air-fuel ratio calculations.

Tests stoichiometric air requirements and excess air calculations
for both gaseous and liquid fuels.
"""

import pytest
from sigma_thermal.combustion.air_fuel import (
    stoich_air_mass_gas,
    stoich_air_vol_gas,
    stoich_air_mass_liquid,
    excess_air_percent,
    GasCompositionMass,
    GasCompositionVolume,
    StoichAirMassGas,
    StoichAirMassLiquid,
    ExcessAirPercent,
)


class TestStoichAirMassGas:
    """Tests for stoich_air_mass_gas function"""

    def test_pure_methane(self):
        """Test pure methane stoichiometric air"""
        comp = GasCompositionMass(methane_mass=100.0)
        result = stoich_air_mass_gas(comp)
        assert result == pytest.approx(17.24, rel=1e-4)

    def test_pure_ethane(self):
        """Test pure ethane stoichiometric air"""
        comp = GasCompositionMass(ethane_mass=100.0)
        result = stoich_air_mass_gas(comp)
        assert result == pytest.approx(16.12, rel=1e-4)

    def test_pure_propane(self):
        """Test pure propane stoichiometric air"""
        comp = GasCompositionMass(propane_mass=100.0)
        result = stoich_air_mass_gas(comp)
        assert result == pytest.approx(15.69, rel=1e-4)

    def test_pure_hydrogen(self):
        """Test pure hydrogen stoichiometric air"""
        comp = GasCompositionMass(hydrogen_mass=100.0)
        result = stoich_air_mass_gas(comp)
        assert result == pytest.approx(34.28, rel=1e-4)

    def test_natural_gas_mixture(self):
        """Test typical natural gas mixture"""
        # 90% CH4, 5% C2H6, 3% C3H8, 2% N2
        comp = GasCompositionMass(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )
        result = stoich_air_mass_gas(comp)
        # Weighted average: 0.90*17.24 + 0.05*16.12 + 0.03*15.69 + 0.02*0
        expected = 0.90 * 17.24 + 0.05 * 16.12 + 0.03 * 15.69
        assert result == pytest.approx(expected, rel=1e-4)

    def test_gas_with_co2_inert(self):
        """Test gas with CO2 inert component"""
        comp = GasCompositionMass(
            methane_mass=95.0,
            co2_mass=5.0
        )
        result = stoich_air_mass_gas(comp)
        # Only CH4 contributes to air requirement
        expected = 0.95 * 17.24
        assert result == pytest.approx(expected, rel=1e-4)

    def test_gas_with_oxygen(self):
        """Test gas with oxygen reduces air requirement"""
        comp = GasCompositionMass(
            methane_mass=98.0,
            o2_mass=2.0
        )
        result = stoich_air_mass_gas(comp)
        # CH4 air requirement minus O2 credit
        expected = 0.98 * 17.24 - 0.02 * 4.32
        assert result == pytest.approx(expected, rel=1e-4)

    def test_complex_mixture(self):
        """Test complex 8-component mixture"""
        comp = GasCompositionMass(
            methane_mass=85.0,
            ethane_mass=6.0,
            propane_mass=4.0,
            butane_mass=2.0,
            pentane_mass=1.0,
            n2_mass=1.5,
            co2_mass=0.5
        )
        result = stoich_air_mass_gas(comp)
        expected = (
            0.85 * 17.24 +
            0.06 * 16.12 +
            0.04 * 15.69 +
            0.02 * 15.47 +
            0.01 * 15.35
        )
        assert result == pytest.approx(expected, rel=1e-4)

    def test_composition_not_100_percent_low(self):
        """Test error when composition sums below 99%"""
        comp = GasCompositionMass(methane_mass=90.0)
        with pytest.raises(ValueError, match="Composition must sum to 100%"):
            stoich_air_mass_gas(comp)

    def test_composition_not_100_percent_high(self):
        """Test error when composition sums above 101%"""
        comp = GasCompositionMass(
            methane_mass=60.0,
            ethane_mass=50.0
        )
        with pytest.raises(ValueError, match="Composition must sum to 100%"):
            stoich_air_mass_gas(comp)

    def test_composition_exactly_100_percent(self):
        """Test valid composition at exactly 100%"""
        comp = GasCompositionMass(
            methane_mass=95.0,
            ethane_mass=5.0
        )
        result = stoich_air_mass_gas(comp)
        assert result > 0  # Should not raise


class TestStoichAirVolGas:
    """Tests for stoich_air_vol_gas function"""

    def test_pure_methane_volume(self):
        """Test pure methane volumetric stoichiometric air"""
        comp = GasCompositionVolume(methane_vol=100.0)
        result = stoich_air_vol_gas(comp)
        assert result == pytest.approx(9.53, rel=1e-4)

    def test_pure_ethane_volume(self):
        """Test pure ethane volumetric stoichiometric air"""
        comp = GasCompositionVolume(ethane_vol=100.0)
        result = stoich_air_vol_gas(comp)
        assert result == pytest.approx(16.68, rel=1e-4)

    def test_pure_propane_volume(self):
        """Test pure propane volumetric stoichiometric air"""
        comp = GasCompositionVolume(propane_vol=100.0)
        result = stoich_air_vol_gas(comp)
        assert result == pytest.approx(23.82, rel=1e-4)

    def test_pure_hydrogen_volume(self):
        """Test pure hydrogen volumetric stoichiometric air"""
        comp = GasCompositionVolume(hydrogen_vol=100.0)
        result = stoich_air_vol_gas(comp)
        assert result == pytest.approx(2.38, rel=1e-4)

    def test_natural_gas_volume_mixture(self):
        """Test typical natural gas volumetric mixture"""
        # 95% CH4, 3% C2H6, 2% N2
        comp = GasCompositionVolume(
            methane_vol=95.0,
            ethane_vol=3.0,
            n2_vol=2.0
        )
        result = stoich_air_vol_gas(comp)
        expected = 0.95 * 9.53 + 0.03 * 16.68
        assert result == pytest.approx(expected, rel=1e-4)

    def test_volume_with_oxygen(self):
        """Test volumetric with oxygen reduces air requirement"""
        comp = GasCompositionVolume(
            methane_vol=97.0,
            o2_vol=3.0
        )
        result = stoich_air_vol_gas(comp)
        # CH4 air requirement minus O2 credit
        expected = 0.97 * 9.53 - 0.03 * 4.76
        assert result == pytest.approx(expected, rel=1e-4)

    def test_volume_composition_validation(self):
        """Test composition must sum to 100%"""
        comp = GasCompositionVolume(methane_vol=85.0)
        with pytest.raises(ValueError, match="Composition must sum to 100%"):
            stoich_air_vol_gas(comp)


class TestStoichAirMassLiquid:
    """Tests for stoich_air_mass_liquid function"""

    def test_number_1_oil(self):
        """Test #1 fuel oil stoichiometric air"""
        result = stoich_air_mass_liquid('#1 oil')
        assert result == pytest.approx(14.7, rel=1e-4)

    def test_number_2_oil(self):
        """Test #2 fuel oil stoichiometric air"""
        result = stoich_air_mass_liquid('#2 oil')
        assert result == pytest.approx(14.5, rel=1e-4)

    def test_number_4_oil(self):
        """Test #4 fuel oil stoichiometric air"""
        result = stoich_air_mass_liquid('#4 oil')
        assert result == pytest.approx(14.0, rel=1e-4)

    def test_number_6_oil(self):
        """Test #6 fuel oil stoichiometric air"""
        result = stoich_air_mass_liquid('#6 oil')
        assert result == pytest.approx(13.5, rel=1e-4)

    def test_gasoline(self):
        """Test gasoline stoichiometric air"""
        result = stoich_air_mass_liquid('gasoline')
        assert result == pytest.approx(14.7, rel=1e-4)

    def test_diesel(self):
        """Test diesel stoichiometric air"""
        result = stoich_air_mass_liquid('diesel')
        assert result == pytest.approx(14.5, rel=1e-4)

    def test_kerosene(self):
        """Test kerosene stoichiometric air"""
        result = stoich_air_mass_liquid('kerosene')
        assert result == pytest.approx(14.7, rel=1e-4)

    def test_methanol(self):
        """Test methanol stoichiometric air"""
        result = stoich_air_mass_liquid('methanol')
        assert result == pytest.approx(6.47, rel=1e-4)

    def test_ethanol(self):
        """Test ethanol stoichiometric air"""
        result = stoich_air_mass_liquid('ethanol')
        assert result == pytest.approx(9.0, rel=1e-4)

    def test_case_insensitive(self):
        """Test case-insensitive fuel type lookup"""
        result1 = stoich_air_mass_liquid('DIESEL')
        result2 = stoich_air_mass_liquid('diesel')
        result3 = stoich_air_mass_liquid('DiEsEl')
        assert result1 == result2 == result3

    def test_whitespace_handling(self):
        """Test whitespace handling in fuel type"""
        result1 = stoich_air_mass_liquid('#2 oil')
        result2 = stoich_air_mass_liquid('  #2 oil  ')
        assert result1 == result2

    def test_unknown_fuel_type(self):
        """Test error for unknown fuel type"""
        with pytest.raises(ValueError, match="Unknown fuel type"):
            stoich_air_mass_liquid('unknown_fuel')

    def test_empty_fuel_type(self):
        """Test error for empty fuel type"""
        with pytest.raises(ValueError):
            stoich_air_mass_liquid('')


class TestExcessAirPercent:
    """Tests for excess_air_percent function"""

    def test_10_percent_excess(self):
        """Test 10% excess air calculation"""
        result = excess_air_percent(1896.4, 1724.0)
        assert result == pytest.approx(10.0, rel=1e-3)

    def test_20_percent_excess(self):
        """Test 20% excess air calculation"""
        result = excess_air_percent(1000.0, 833.33)
        assert result == pytest.approx(20.0, rel=1e-2)

    def test_zero_excess_air(self):
        """Test stoichiometric combustion (zero excess air)"""
        result = excess_air_percent(100.0, 100.0)
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_5_percent_excess(self):
        """Test 5% excess air (typical for gas)"""
        result = excess_air_percent(105.0, 100.0)
        assert result == pytest.approx(5.0, rel=1e-4)

    def test_25_percent_excess(self):
        """Test 25% excess air (typical for oil)"""
        result = excess_air_percent(125.0, 100.0)
        assert result == pytest.approx(25.0, rel=1e-4)

    def test_large_excess_air(self):
        """Test large excess air (50%)"""
        result = excess_air_percent(150.0, 100.0)
        assert result == pytest.approx(50.0, rel=1e-4)

    def test_very_low_excess(self):
        """Test very low excess air (1%)"""
        result = excess_air_percent(101.0, 100.0)
        assert result == pytest.approx(1.0, rel=1e-4)

    def test_negative_excess_air(self):
        """Test sub-stoichiometric (negative excess air)"""
        result = excess_air_percent(90.0, 100.0)
        assert result == pytest.approx(-10.0, rel=1e-4)

    def test_zero_stoich_air_error(self):
        """Test error when stoichiometric air is zero"""
        with pytest.raises(ValueError, match="Stoichiometric air must be positive"):
            excess_air_percent(100.0, 0.0)

    def test_negative_stoich_air_error(self):
        """Test error when stoichiometric air is negative"""
        with pytest.raises(ValueError, match="Stoichiometric air must be positive"):
            excess_air_percent(100.0, -50.0)

    def test_realistic_natural_gas_case(self):
        """Test realistic natural gas combustion case"""
        # Natural gas: 17.24 lb air/lb fuel, 10% excess
        stoich = 17.24 * 100.0  # lb/hr for 100 lb/hr fuel
        actual = stoich * 1.10  # 10% excess
        result = excess_air_percent(actual, stoich)
        assert result == pytest.approx(10.0, rel=1e-3)


class TestVBACompatibility:
    """Tests for VBA-compatible wrapper functions"""

    def test_stoich_air_mass_gas_vba(self):
        """Test VBA wrapper for stoich_air_mass_gas"""
        result = StoichAirMassGas(
            methane_mass=90.0,
            ethane_mass=5.0,
            propane_mass=3.0,
            n2_mass=2.0
        )
        expected = 0.90 * 17.24 + 0.05 * 16.12 + 0.03 * 15.69
        assert result == pytest.approx(expected, rel=1e-4)

    def test_stoich_air_mass_gas_vba_defaults(self):
        """Test VBA wrapper with default values"""
        result = StoichAirMassGas(methane_mass=100.0)
        assert result == pytest.approx(17.24, rel=1e-4)

    def test_stoich_air_mass_liquid_vba(self):
        """Test VBA wrapper for stoich_air_mass_liquid"""
        result = StoichAirMassLiquid('#2 oil')
        assert result == pytest.approx(14.5, rel=1e-4)

    def test_excess_air_percent_vba(self):
        """Test VBA wrapper for excess_air_percent"""
        result = ExcessAirPercent(110.0, 100.0)
        assert result == pytest.approx(10.0, rel=1e-4)


class TestIntegrationScenarios:
    """Integration tests with realistic combustion scenarios"""

    def test_natural_gas_boiler(self):
        """Test complete air requirement calculation for natural gas boiler"""
        # Fuel composition
        comp = GasCompositionMass(
            methane_mass=93.0,
            ethane_mass=4.0,
            propane_mass=2.0,
            n2_mass=1.0
        )

        # Calculate stoichiometric air
        stoich_air = stoich_air_mass_gas(comp)
        expected_stoich = 0.93 * 17.24 + 0.04 * 16.12 + 0.02 * 15.69
        assert stoich_air == pytest.approx(expected_stoich, rel=1e-4)

        # 10% excess air typical for gas
        fuel_rate = 100.0  # lb/hr
        stoich_air_rate = stoich_air * fuel_rate
        actual_air_rate = stoich_air_rate * 1.10

        excess = excess_air_percent(actual_air_rate, stoich_air_rate)
        assert excess == pytest.approx(10.0, rel=1e-3)

    def test_oil_fired_heater(self):
        """Test complete air requirement for oil-fired heater"""
        # #2 oil stoichiometric air
        stoich_air = stoich_air_mass_liquid('#2 oil')
        assert stoich_air == 14.5

        # 20% excess air typical for oil
        fuel_rate = 50.0  # lb/hr
        stoich_air_rate = stoich_air * fuel_rate
        actual_air_rate = stoich_air_rate * 1.20

        excess = excess_air_percent(actual_air_rate, stoich_air_rate)
        assert excess == pytest.approx(20.0, rel=1e-3)

    def test_hydrogen_combustion(self):
        """Test hydrogen combustion air requirements"""
        comp = GasCompositionMass(hydrogen_mass=100.0)
        stoich_air = stoich_air_mass_gas(comp)

        # Hydrogen requires much more air per unit mass
        assert stoich_air == pytest.approx(34.28, rel=1e-4)
        assert stoich_air > 17.24  # More than methane
