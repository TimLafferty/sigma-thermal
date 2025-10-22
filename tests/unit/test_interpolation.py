"""
Unit tests for interpolation utilities.
"""

import pytest
import numpy as np
from sigma_thermal.engineering.interpolation import (
    linear_interpolate,
    interpolate_from_table,
    bilinear_interpolate,
    Interpolate,
)


class TestLinearInterpolate:
    """Tests for linear_interpolate function"""

    def test_midpoint_interpolation(self):
        """Test interpolation at midpoint"""
        result = linear_interpolate(0, 5, 10, 0, 100)
        assert result == pytest.approx(50.0)

    def test_quarter_point_interpolation(self):
        """Test interpolation at 1/4 point"""
        result = linear_interpolate(0, 2.5, 10, 0, 100)
        assert result == pytest.approx(25.0)

    def test_three_quarter_point(self):
        """Test interpolation at 3/4 point"""
        result = linear_interpolate(100, 150, 200, 20, 30)
        assert result == pytest.approx(25.0)

    def test_at_lower_bound(self):
        """Test interpolation at lower bound"""
        result = linear_interpolate(0, 0, 10, 0, 100)
        assert result == pytest.approx(0.0)

    def test_at_upper_bound(self):
        """Test interpolation at upper bound"""
        result = linear_interpolate(0, 10, 10, 0, 100)
        assert result == pytest.approx(100.0)

    def test_extrapolation_below(self):
        """Test extrapolation below range"""
        result = linear_interpolate(0, -5, 10, 0, 100)
        assert result == pytest.approx(-50.0)

    def test_extrapolation_above(self):
        """Test extrapolation above range"""
        result = linear_interpolate(0, 15, 10, 0, 100)
        assert result == pytest.approx(150.0)

    def test_negative_slope(self):
        """Test with decreasing function"""
        result = linear_interpolate(0, 5, 10, 100, 0)
        assert result == pytest.approx(50.0)

    def test_zero_division_protection(self):
        """Test that equal x values don't cause division by zero"""
        result = linear_interpolate(5, 5, 5, 10, 20)
        assert result == pytest.approx(10.0)

    def test_vba_compatibility(self):
        """Test that Interpolate alias works"""
        result = Interpolate(0, 5, 10, 0, 100)
        assert result == pytest.approx(50.0)


class TestInterpolateFromTable:
    """Tests for interpolate_from_table function"""

    def test_simple_table(self):
        """Test interpolation from a simple table"""
        x_vals = [0, 10, 20, 30]
        y_vals = [0, 100, 150, 180]

        result = interpolate_from_table(15, x_vals, y_vals)
        assert result == pytest.approx(125.0)

    def test_exact_match(self):
        """Test when x matches a table value"""
        x_vals = [0, 10, 20, 30]
        y_vals = [0, 100, 150, 180]

        result = interpolate_from_table(20, x_vals, y_vals)
        assert result == pytest.approx(150.0)

    def test_numpy_arrays(self):
        """Test with numpy arrays"""
        x_vals = np.array([0, 10, 20, 30])
        y_vals = np.array([0, 100, 150, 180])

        result = interpolate_from_table(15, x_vals, y_vals)
        assert result == pytest.approx(125.0)

    def test_extrapolation_allowed(self):
        """Test extrapolation when allowed"""
        x_vals = [0, 10, 20]
        y_vals = [0, 100, 200]

        result = interpolate_from_table(25, x_vals, y_vals, extrapolate=True)
        assert result == pytest.approx(250.0)

    def test_extrapolation_disabled(self):
        """Test boundary clipping when extrapolation disabled"""
        x_vals = [0, 10, 20]
        y_vals = [0, 100, 200]

        # Above range
        result = interpolate_from_table(25, x_vals, y_vals, extrapolate=False)
        assert result == pytest.approx(200.0)

        # Below range
        result = interpolate_from_table(-5, x_vals, y_vals, extrapolate=False)
        assert result == pytest.approx(0.0)

    def test_mismatched_lengths(self):
        """Test error on mismatched array lengths"""
        with pytest.raises(ValueError, match="same length"):
            interpolate_from_table(5, [0, 10], [0, 100, 200])

    def test_insufficient_points(self):
        """Test error with too few points"""
        with pytest.raises(ValueError, match="at least 2"):
            interpolate_from_table(5, [10], [100])


class TestBilinearInterpolate:
    """Tests for bilinear_interpolate function"""

    def test_simple_grid(self):
        """Test bilinear interpolation on simple grid"""
        x_vals = [0, 10, 20]
        y_vals = [0, 5, 10]
        z_grid = np.array([
            [0, 5, 10],
            [10, 15, 20],
            [20, 25, 30]
        ])

        # Midpoint should be average
        result = bilinear_interpolate(10, 5, x_vals, y_vals, z_grid)
        assert result == pytest.approx(15.0)

    def test_corner_values(self):
        """Test that corner values are exact"""
        x_vals = [0, 10, 20]
        y_vals = [0, 5, 10]
        z_grid = np.array([
            [0, 5, 10],
            [10, 15, 20],
            [20, 25, 30]
        ])

        # Test all corners
        assert bilinear_interpolate(0, 0, x_vals, y_vals, z_grid) == pytest.approx(0.0)
        assert bilinear_interpolate(20, 0, x_vals, y_vals, z_grid) == pytest.approx(20.0)
        assert bilinear_interpolate(0, 10, x_vals, y_vals, z_grid) == pytest.approx(10.0)
        assert bilinear_interpolate(20, 10, x_vals, y_vals, z_grid) == pytest.approx(30.0)

    def test_property_table_use_case(self):
        """Test realistic use case with property table"""
        # Temperature in °F
        temps = [100, 200, 300]
        # Pressure in psi
        pressures = [10, 20, 30]
        # Density values
        densities = np.array([
            [60.0, 59.5, 59.0],
            [59.0, 58.5, 58.0],
            [58.0, 57.5, 57.0]
        ])

        # Interpolate at T=150°F, P=15 psi
        result = bilinear_interpolate(150, 15, temps, pressures, densities)
        assert result == pytest.approx(59.25, abs=0.01)
