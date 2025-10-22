"""
Unit tests for unit conversion utilities.
"""

import pytest
import numpy as np
from sigma_thermal.engineering.units import (
    convert,
    Q_,
    ensure_units,
    btu_hr_to_kw,
    kw_to_btu_hr,
    degf_to_degc,
    degc_to_degf,
    psi_to_pa,
    pa_to_psi,
    ureg,
)


class TestConvert:
    """Tests for convert function"""

    def test_power_conversion(self):
        """Test power unit conversions"""
        result = convert(1000, 'Btu/hr', 'kW')
        assert result == pytest.approx(0.293071, rel=1e-4)

    def test_temperature_conversion(self):
        """Test temperature conversions"""
        result = convert(100, 'degF', 'degC')
        assert result == pytest.approx(37.7778, rel=1e-4)

    def test_length_conversion(self):
        """Test length conversions"""
        result = convert(1, 'ft', 'meter')
        assert result == pytest.approx(0.3048, rel=1e-4)

    def test_pressure_conversion(self):
        """Test pressure conversions"""
        result = convert(14.7, 'psi', 'kPa')
        assert result == pytest.approx(101.353, rel=1e-3)

    def test_array_conversion(self):
        """Test conversion of numpy arrays"""
        values = np.array([1, 2, 3, 4, 5])
        result = convert(values, 'ft', 'meter')
        expected = values * 0.3048
        np.testing.assert_allclose(result, expected, rtol=1e-4)


class TestQuantityCreation:
    """Tests for Q_ quantity creation"""

    def test_create_quantity(self):
        """Test creating a quantity"""
        duty = Q_(1000000, 'Btu/hr')
        assert duty.magnitude == 1000000
        # Pint uses full unit names in string representation
        assert str(duty.units) == 'british_thermal_unit / hour'

    def test_quantity_conversion(self):
        """Test converting quantities"""
        duty = Q_(1000000, 'Btu/hr')
        duty_kw = duty.to('kW')
        assert duty_kw.magnitude == pytest.approx(293.071, rel=1e-4)

    def test_quantity_arithmetic(self):
        """Test arithmetic with quantities"""
        q1 = Q_(1000, 'Btu/hr')
        q2 = Q_(500, 'Btu/hr')
        result = q1 + q2
        assert result.magnitude == 1500

    def test_quantity_unit_mismatch_error(self):
        """Test error on incompatible units"""
        q1 = Q_(100, 'degF')
        q2 = Q_(50, 'ft')
        with pytest.raises(Exception):  # pint raises DimensionalityError
            result = q1 + q2


class TestEnsureUnits:
    """Tests for ensure_units function"""

    def test_attach_units_to_float(self):
        """Test attaching units to a plain float"""
        result = ensure_units(100, 'degF')
        assert result.magnitude == 100
        assert str(result.units) == 'degree_Fahrenheit'

    def test_convert_existing_units(self):
        """Test converting quantity to different units"""
        temp_c = Q_(37.7778, 'degC')
        result = ensure_units(temp_c, 'degF')
        assert result.magnitude == pytest.approx(100, rel=1e-3)


class TestCommonConversions:
    """Tests for common conversion functions"""

    def test_btu_hr_to_kw(self):
        """Test BTU/hr to kW conversion"""
        result = btu_hr_to_kw(1000)
        assert result == pytest.approx(0.293071, rel=1e-4)

    def test_kw_to_btu_hr(self):
        """Test kW to BTU/hr conversion"""
        result = kw_to_btu_hr(1)
        assert result == pytest.approx(3412.14, rel=1e-4)

    def test_roundtrip_power(self):
        """Test roundtrip power conversion"""
        original = 5000
        converted = btu_hr_to_kw(original)
        back = kw_to_btu_hr(converted)
        assert back == pytest.approx(original, rel=1e-6)

    def test_degf_to_degc(self):
        """Test Fahrenheit to Celsius"""
        assert degf_to_degc(32) == pytest.approx(0, abs=1e-6)
        assert degf_to_degc(212) == pytest.approx(100, abs=1e-6)
        assert degf_to_degc(100) == pytest.approx(37.7778, rel=1e-4)

    def test_degc_to_degf(self):
        """Test Celsius to Fahrenheit"""
        assert degc_to_degf(0) == pytest.approx(32, abs=1e-6)
        assert degc_to_degf(100) == pytest.approx(212, abs=1e-6)
        assert degc_to_degf(37.7778) == pytest.approx(100, rel=1e-4)

    def test_roundtrip_temperature(self):
        """Test roundtrip temperature conversion"""
        original = 150
        celsius = degf_to_degc(original)
        back = degc_to_degf(celsius)
        assert back == pytest.approx(original, rel=1e-6)

    def test_psi_to_pa(self):
        """Test PSI to Pascal conversion"""
        result = psi_to_pa(14.7)
        assert result == pytest.approx(101353, rel=1e-3)

    def test_pa_to_psi(self):
        """Test Pascal to PSI conversion"""
        result = pa_to_psi(101353)
        assert result == pytest.approx(14.7, rel=1e-3)


class TestCustomUnits:
    """Tests for custom unit definitions"""

    def test_scfh_unit(self):
        """Test SCFH custom unit"""
        flow = Q_(1000, 'scfh')
        flow_cfm = flow.to('ft**3/min')
        assert flow_cfm.magnitude == pytest.approx(16.6667, rel=1e-4)

    def test_mmBtu_unit(self):
        """Test million BTU custom unit"""
        energy = Q_(1, 'mmBtu')
        energy_btu = energy.to('Btu')
        assert energy_btu.magnitude == pytest.approx(1e6, rel=1e-6)


class TestDimensionalAnalysis:
    """Tests for dimensional analysis capabilities"""

    def test_heat_duty_calculation(self):
        """Test dimensional analysis in heat duty calculation"""
        # Q = m * cp * ΔT
        mass_flow = Q_(1000, 'lb/hr')
        cp = Q_(1.0, 'Btu/(lb*degF)')
        delta_t = Q_(50, 'degF')

        duty = mass_flow * cp * delta_t
        duty_btu_hr = duty.to('Btu/hr')

        assert duty_btu_hr.magnitude == pytest.approx(50000, rel=1e-6)

    def test_velocity_calculation(self):
        """Test velocity from flow and area"""
        flow = Q_(1000, 'gal/min')
        area = Q_(10, 'inch**2')

        velocity = flow / area
        velocity_ft_s = velocity.to('ft/s')

        # Check that result is reasonable
        assert velocity_ft_s.magnitude > 0
