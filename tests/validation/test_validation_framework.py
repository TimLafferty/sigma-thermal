"""
Validation test framework for comparing Python calculations against Excel.

This module provides the base framework for validation testing.
"""

import pytest
import pandas as pd
from pathlib import Path
from typing import Dict, Any
import json


class ValidationTestCase:
    """Base class for validation test cases"""

    def __init__(self, case_name: str, case_data: Dict[str, Any]):
        self.case_name = case_name
        self.inputs = case_data.get('inputs', {})
        self.excel_outputs = case_data.get('outputs', {})

    def compare_results(self, python_outputs: Dict[str, Any], tolerance: float = 0.01):
        """
        Compare Python outputs against Excel outputs.

        Parameters
        ----------
        python_outputs : dict
            Results from Python calculation
        tolerance : float
            Relative tolerance for comparison (default 1%)

        Returns
        -------
        dict
            Comparison results with pass/fail for each output
        """
        comparison = {}

        for key, excel_value in self.excel_outputs.items():
            if key not in python_outputs:
                comparison[key] = {
                    'status': 'MISSING',
                    'excel': excel_value,
                    'python': None,
                    'error': 'Output not found in Python results'
                }
                continue

            python_value = python_outputs[key]

            # Handle numeric comparisons
            if isinstance(excel_value, (int, float)) and isinstance(python_value, (int, float)):
                if excel_value == 0:
                    # Absolute comparison for zero values
                    diff = abs(python_value - excel_value)
                    passed = diff < 1e-6
                else:
                    # Relative comparison
                    rel_diff = abs(python_value - excel_value) / abs(excel_value)
                    passed = rel_diff <= tolerance

                comparison[key] = {
                    'status': 'PASS' if passed else 'FAIL',
                    'excel': excel_value,
                    'python': python_value,
                    'rel_diff': rel_diff if excel_value != 0 else diff,
                    'tolerance': tolerance
                }
            else:
                # Direct comparison for non-numeric values
                passed = str(excel_value) == str(python_value)
                comparison[key] = {
                    'status': 'PASS' if passed else 'FAIL',
                    'excel': excel_value,
                    'python': python_value,
                }

        return comparison

    def assert_all_pass(self, comparison: Dict[str, Any]):
        """Assert that all comparisons passed"""
        failures = [k for k, v in comparison.items() if v['status'] != 'PASS']

        if failures:
            error_msg = f"\nValidation failures for {self.case_name}:\n"
            for key in failures:
                result = comparison[key]
                error_msg += f"\n  {key}:"
                error_msg += f"\n    Excel:  {result.get('excel')}"
                error_msg += f"\n    Python: {result.get('python')}"
                if 'rel_diff' in result:
                    error_msg += f"\n    Rel Diff: {result['rel_diff']:.4%}"
                if 'error' in result:
                    error_msg += f"\n    Error: {result['error']}"

            pytest.fail(error_msg)


def load_validation_case(case_file: Path) -> ValidationTestCase:
    """Load a validation test case from JSON file"""
    with open(case_file) as f:
        case_data = json.load(f)

    case_name = case_file.stem
    return ValidationTestCase(case_name, case_data)


@pytest.fixture
def validation_case_loader(test_data_dir):
    """Fixture to load validation cases"""
    def loader(case_name: str) -> ValidationTestCase:
        case_file = test_data_dir / f"{case_name}.json"
        if not case_file.exists():
            pytest.skip(f"Validation case file not found: {case_file}")
        return load_validation_case(case_file)
    return loader


# Example validation test (placeholder until real cases are created)
class TestValidationFramework:
    """Test the validation framework itself"""

    def test_comparison_pass(self):
        """Test that identical values pass comparison"""
        case = ValidationTestCase('test', {
            'inputs': {},
            'outputs': {'value1': 100.0, 'value2': 200.0}
        })

        python_outputs = {'value1': 100.0, 'value2': 200.0}
        comparison = case.compare_results(python_outputs, tolerance=0.01)

        assert comparison['value1']['status'] == 'PASS'
        assert comparison['value2']['status'] == 'PASS'

    def test_comparison_within_tolerance(self):
        """Test that values within tolerance pass"""
        case = ValidationTestCase('test', {
            'inputs': {},
            'outputs': {'value1': 100.0}
        })

        python_outputs = {'value1': 100.5}  # 0.5% difference
        comparison = case.compare_results(python_outputs, tolerance=0.01)

        assert comparison['value1']['status'] == 'PASS'

    def test_comparison_exceeds_tolerance(self):
        """Test that values exceeding tolerance fail"""
        case = ValidationTestCase('test', {
            'inputs': {},
            'outputs': {'value1': 100.0}
        })

        python_outputs = {'value1': 102.0}  # 2% difference
        comparison = case.compare_results(python_outputs, tolerance=0.01)

        assert comparison['value1']['status'] == 'FAIL'

    def test_missing_output(self):
        """Test handling of missing outputs"""
        case = ValidationTestCase('test', {
            'inputs': {},
            'outputs': {'value1': 100.0, 'value2': 200.0}
        })

        python_outputs = {'value1': 100.0}  # Missing value2
        comparison = case.compare_results(python_outputs)

        assert comparison['value1']['status'] == 'PASS'
        assert comparison['value2']['status'] == 'MISSING'
