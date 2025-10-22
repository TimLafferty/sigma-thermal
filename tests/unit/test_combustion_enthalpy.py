"""
Unit tests for combustion enthalpy functions.

These tests validate the Python implementation against the VBA functions
from Engineering-Functions.xlam CombustionFunctions.bas module.
"""

import pytest
import numpy as np
from sigma_thermal.combustion.enthalpy import (
    enthalpy_co2,
    enthalpy_h2o,
    enthalpy_n2,
    enthalpy_o2,
    flue_gas_enthalpy,
    EnthalpyCO2,  # VBA alias
    EnthalpyH2O,
    EnthalpyN2,
    EnthalpyO2,
)
from sigma_thermal.engineering.units import Q_


class TestEnthalpyCO2:
    """Tests for enthalpy_co2 function"""

    def test_reference_temperature_zero(self):
        """Test that enthalpy is zero at reference temperature"""
        result = enthalpy_co2(77, 77)
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_typical_stack_temperature(self):
        """Test typical stack temperature (1500°F)"""
        result = enthalpy_co2(1500, 77)
        # Calculate expected value using VBA coefficients
        a, b, c = 1.08941e-05, 0.262597665, 176.9479842
        h_1500 = a * 1500**2 + b * 1500 + c
        h_77 = a * 77**2 + b * 77 + c
        expected = h_1500 - h_77
        assert result == pytest.approx(expected, rel=1e-6)

    def test_high_temperature(self):
        """Test high temperature (2500°F)"""
        result = enthalpy_co2(2500, 77)
        assert result > 0
        assert result > enthalpy_co2(1500, 77)  # Should be higher

    def test_below_ambient(self):
        """Test temperature below ambient"""
        result = enthalpy_co2(32, 77)  # Freezing point
        assert result < 0  # Should be negative

    def test_different_ambient(self):
        """Test with different ambient temperature"""
        h1 = enthalpy_co2(1500, 60)
        h2 = enthalpy_co2(1500, 77)
        assert h1 > h2  # Lower ambient means more enthalpy difference

    def test_vba_alias(self):
        """Test VBA compatibility alias"""
        result1 = enthalpy_co2(1500, 77)
        result2 = EnthalpyCO2(1500, 77)
        assert result1 == result2

    def test_with_quantity(self):
        """Test with pint Quantity input"""
        T_gas = Q_(1500, 'degF')
        T_amb = Q_(77, 'degF')
        result = enthalpy_co2(T_gas, T_amb, return_quantity=True)

        assert result.units == Q_(1, 'Btu/lb').units
        assert result.magnitude == pytest.approx(enthalpy_co2(1500, 77), rel=1e-6)

    def test_celsius_input(self):
        """Test with Celsius input (via Quantity)"""
        T_gas = Q_(800, 'degC')  # ~1472°F
        T_amb = Q_(25, 'degC')   # ~77°F
        result = enthalpy_co2(T_gas, T_amb)

        # Should be similar to 1500°F vs 77°F
        assert result == pytest.approx(enthalpy_co2(1472, 77), rel=0.01)


class TestEnthalpyH2O:
    """Tests for enthalpy_h2o function"""

    def test_reference_temperature_zero(self):
        """Test that enthalpy is zero at reference temperature"""
        result = enthalpy_h2o(77, 77)
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_typical_stack_temperature(self):
        """Test typical stack temperature (1500°F)"""
        result = enthalpy_h2o(1500, 77)
        # Calculate expected value using VBA coefficients
        a, b, c = 3.65285e-05, 0.452215911, 1049.366151
        h_1500 = a * 1500**2 + b * 1500 + c
        h_77 = a * 77**2 + b * 77 + c
        expected = h_1500 - h_77
        assert result == pytest.approx(expected, rel=1e-6)

    def test_higher_than_co2(self):
        """Test that H2O enthalpy is higher than CO2 (higher heat capacity)"""
        h_h2o = enthalpy_h2o(1500, 77)
        h_co2 = enthalpy_co2(1500, 77)
        assert h_h2o > h_co2  # Water vapor has higher specific heat

    def test_vba_alias(self):
        """Test VBA compatibility alias"""
        result1 = enthalpy_h2o(1500, 77)
        result2 = EnthalpyH2O(1500, 77)
        assert result1 == result2


class TestEnthalpyN2:
    """Tests for enthalpy_n2 function"""

    def test_reference_temperature_zero(self):
        """Test that enthalpy is zero at reference temperature"""
        result = enthalpy_n2(77, 77)
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_typical_stack_temperature(self):
        """Test typical stack temperature (1500°F)"""
        result = enthalpy_n2(1500, 77)
        # Calculate expected value using VBA coefficients
        a, b, c = 8.46332e-06, 0.255630011, 107.2712456
        h_1500 = a * 1500**2 + b * 1500 + c
        h_77 = a * 77**2 + b * 77 + c
        expected = h_1500 - h_77
        assert result == pytest.approx(expected, rel=1e-6)

    def test_vba_alias(self):
        """Test VBA compatibility alias"""
        result1 = enthalpy_n2(1500, 77)
        result2 = EnthalpyN2(1500, 77)
        assert result1 == result2


class TestEnthalpyO2:
    """Tests for enthalpy_o2 function"""

    def test_reference_temperature_zero(self):
        """Test that enthalpy is zero at reference temperature"""
        result = enthalpy_o2(77, 77)
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_typical_stack_temperature(self):
        """Test typical stack temperature (1500°F)"""
        result = enthalpy_o2(1500, 77)
        # Calculate expected value using VBA coefficients
        a, b, c = 7.53536e-06, 0.23706691, 92.56930357
        h_1500 = a * 1500**2 + b * 1500 + c
        h_77 = a * 77**2 + b * 77 + c
        expected = h_1500 - h_77
        assert result == pytest.approx(expected, rel=1e-6)

    def test_vba_alias(self):
        """Test VBA compatibility alias"""
        result1 = enthalpy_o2(1500, 77)
        result2 = EnthalpyO2(1500, 77)
        assert result1 == result2


class TestRelativeEnthalpies:
    """Test relative magnitudes of different gas enthalpies"""

    def test_enthalpy_ordering(self):
        """Test that enthalpies are in expected order"""
        temp = 1500
        amb = 77

        h_h2o = enthalpy_h2o(temp, amb)
        h_n2 = enthalpy_n2(temp, amb)
        h_o2 = enthalpy_o2(temp, amb)
        h_co2 = enthalpy_co2(temp, amb)

        # H2O should be highest (highest specific heat)
        assert h_h2o > h_n2
        assert h_h2o > h_o2
        assert h_h2o > h_co2

        # All should be positive for temp > amb
        assert h_h2o > 0
        assert h_n2 > 0
        assert h_o2 > 0
        assert h_co2 > 0

    def test_temperature_dependence(self):
        """Test that enthalpy increases with temperature"""
        temps = [500, 1000, 1500, 2000, 2500]
        amb = 77

        for gas_func in [enthalpy_co2, enthalpy_h2o, enthalpy_n2, enthalpy_o2]:
            enthalpies = [gas_func(T, amb) for T in temps]
            # Check monotonically increasing
            assert all(enthalpies[i] < enthalpies[i+1] for i in range(len(enthalpies)-1))


class TestFlueGasEnthalpy:
    """Tests for flue_gas_enthalpy function"""

    def test_pure_n2(self):
        """Test with pure nitrogen"""
        result = flue_gas_enthalpy(
            h2o_fraction=0.0,
            co2_fraction=0.0,
            n2_fraction=1.0,
            o2_fraction=0.0,
            gas_temp=1500,
            ambient_temp=77
        )
        expected = enthalpy_n2(1500, 77)
        assert result == pytest.approx(expected, rel=1e-6)

    def test_pure_co2(self):
        """Test with pure CO2"""
        result = flue_gas_enthalpy(
            h2o_fraction=0.0,
            co2_fraction=1.0,
            n2_fraction=0.0,
            o2_fraction=0.0,
            gas_temp=1500,
            ambient_temp=77
        )
        expected = enthalpy_co2(1500, 77)
        assert result == pytest.approx(expected, rel=1e-6)

    def test_typical_natural_gas_flue_gas(self):
        """Test with typical natural gas flue gas composition"""
        # Typical mass fractions for natural gas combustion products
        result = flue_gas_enthalpy(
            h2o_fraction=0.12,  # 12% water vapor
            co2_fraction=0.15,  # 15% CO2
            n2_fraction=0.70,   # 70% N2
            o2_fraction=0.03,   # 3% excess O2
            gas_temp=1500,
            ambient_temp=77
        )

        # Calculate expected manually
        h_h2o = enthalpy_h2o(1500, 77)
        h_co2 = enthalpy_co2(1500, 77)
        h_n2 = enthalpy_n2(1500, 77)
        h_o2 = enthalpy_o2(1500, 77)

        expected = 0.12 * h_h2o + 0.15 * h_co2 + 0.70 * h_n2 + 0.03 * h_o2

        assert result == pytest.approx(expected, rel=1e-6)

    def test_fractions_must_sum_to_one(self):
        """Test that fractions must sum to unity"""
        # Fractions sum to 0.5 - should fail
        with pytest.raises(ValueError, match="sum to unity"):
            flue_gas_enthalpy(
                h2o_fraction=0.1,
                co2_fraction=0.1,
                n2_fraction=0.2,
                o2_fraction=0.1,
                gas_temp=1500,
                ambient_temp=77
            )

    def test_fractions_within_tolerance(self):
        """Test that fractions within 1% tolerance are accepted"""
        # Fractions sum to 1.005 - should pass
        result = flue_gas_enthalpy(
            h2o_fraction=0.1205,
            co2_fraction=0.1500,
            n2_fraction=0.7000,
            o2_fraction=0.0300,  # Sum = 1.0005
            gas_temp=1500,
            ambient_temp=77
        )
        assert result > 0  # Should calculate successfully

    def test_reference_temperature_zero(self):
        """Test that enthalpy is zero at reference temperature"""
        result = flue_gas_enthalpy(
            h2o_fraction=0.12,
            co2_fraction=0.15,
            n2_fraction=0.70,
            o2_fraction=0.03,
            gas_temp=77,
            ambient_temp=77
        )
        assert result == pytest.approx(0.0, abs=1e-6)

    def test_with_quantity(self):
        """Test with pint Quantity inputs"""
        T_gas = Q_(1500, 'degF')
        T_amb = Q_(77, 'degF')

        result = flue_gas_enthalpy(
            h2o_fraction=0.12,
            co2_fraction=0.15,
            n2_fraction=0.70,
            o2_fraction=0.03,
            gas_temp=T_gas,
            ambient_temp=T_amb,
            return_quantity=True
        )

        assert result.units == Q_(1, 'Btu/lb').units
        assert result.magnitude > 0


class TestEnthalpyEdgeCases:
    """Test edge cases and boundary conditions"""

    def test_very_high_temperature(self):
        """Test with very high temperature (near limit of correlation)"""
        temp = 3000  # °F
        # All functions should still work
        assert enthalpy_co2(temp, 77) > 0
        assert enthalpy_h2o(temp, 77) > 0
        assert enthalpy_n2(temp, 77) > 0
        assert enthalpy_o2(temp, 77) > 0

    def test_array_inputs(self):
        """Test that functions work with numpy arrays (via loop/vectorization)"""
        temps = np.array([500, 1000, 1500, 2000])
        results = [enthalpy_co2(T, 77) for T in temps]

        assert len(results) == len(temps)
        assert all(r > 0 for r in results)
        # Check monotonically increasing
        assert all(results[i] < results[i+1] for i in range(len(results)-1))

    def test_negative_delta_t(self):
        """Test with gas temperature below ambient"""
        result = enthalpy_co2(32, 77)
        assert result < 0  # Should be negative

    def test_zero_ambient(self):
        """Test with zero absolute as ambient (extreme case)"""
        result = enthalpy_co2(500, -459.67)  # -459.67°F = 0 Rankine
        assert result > 0  # Should still work


class TestEnthalpyPhysicalMeaning:
    """Test physical meaning and relationships"""

    def test_energy_conservation(self):
        """Test that enthalpy change is path-independent"""
        # H(T1 -> T3) should equal H(T1 -> T2) + H(T2 -> T3)
        T1, T2, T3 = 77, 500, 1500

        h_direct = enthalpy_co2(T3, T1)
        h_step1 = enthalpy_co2(T2, T1)
        h_step2 = enthalpy_co2(T3, T2)
        h_stepwise = h_step1 + h_step2

        assert h_direct == pytest.approx(h_stepwise, rel=1e-6)

    def test_symmetry(self):
        """Test that H(T1, T2) = -H(T2, T1)"""
        T1, T2 = 77, 1500

        h_forward = enthalpy_co2(T2, T1)
        h_reverse = enthalpy_co2(T1, T2)

        assert h_forward == pytest.approx(-h_reverse, rel=1e-6)
