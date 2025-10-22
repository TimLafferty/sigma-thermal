# Fluids Module Excel Validation Test Plan

**Project:** Sigma Thermal Fluids Validation
**Date:** October 22, 2025
**Purpose:** Define test plan for validating Python fluids module against Excel VBA macros
**Status:** Planning - Implementation Week 3

---

## Objective

Validate all 8 fluids module functions against Excel VBA equivalents to ensure 1:1 parity with existing Engineering-Functions.xlam macros.

**Success Criteria:**
- All 8 functions validated with <1% deviation
- Minimum 20 test cases covering full operating range
- Discrepancies documented and explained
- Test suite integrated into CI/CD pipeline

---

## Functions to Validate

| # | Python Function | Excel VBA Function | Priority | Test Cases |
|---|----------------|-------------------|----------|------------|
| 1 | `saturation_pressure()` | `SaturationPressure(T)` | High | 8 |
| 2 | `saturation_temperature()` | `SaturationTemperature(P)` | High | 8 |
| 3 | `water_density()` | `WaterDensity(T, P)` | High | 6 |
| 4 | `water_viscosity()` | `WaterViscosity(T)` | High | 6 |
| 5 | `water_specific_heat()` | `WaterSpecificHeat(T)` | Medium | 6 |
| 6 | `water_thermal_conductivity()` | `WaterThermalConductivity(T)` | Medium | 6 |
| 7 | `steam_enthalpy()` | `SteamEnthalpy(T, P, x)` | High | 12 |
| 8 | `steam_quality()` | `SteamQuality(h, P)` | High | 8 |

**Total Test Cases:** 60 minimum

---

## Test Case Structure

### Template for Each Test Case

```python
{
    "case_name": "Saturation Pressure at Boiling Point",
    "function": "saturation_pressure",
    "inputs": {
        "temperature": 212.0  # °F
    },
    "excel_output": 14.696,  # psia (from Excel VBA)
    "python_output": None,  # To be calculated
    "tolerance": 0.01,  # 1% relative
    "reference": "ASME Steam Tables",
    "expected_status": "PASS"
}
```

---

## Saturation Properties Test Cases

### 1. Saturation Pressure Tests (8 cases)

| Case | Temperature (°F) | Expected P (psia) | Reference | Notes |
|------|-----------------|-------------------|-----------|-------|
| SP-1 | 32 | 0.089 | ASME | Freezing point |
| SP-2 | 100 | 0.949 | ASME | Low temp |
| SP-3 | 212 | 14.696 | ASME | Boiling point |
| SP-4 | 300 | 66.98 | ASME | Moderate |
| SP-5 | 400 | 247.1 | ASME | High temp |
| SP-6 | 500 | 680.0 | ASME | Very high temp |
| SP-7 | 600 | 1542.0 | ASME | Extreme temp |
| SP-8 | 327.8 | 100.0 | ASME | Round pressure |

**Tolerance:** 0.5% (saturation properties are critical)

---

### 2. Saturation Temperature Tests (8 cases)

| Case | Pressure (psia) | Expected T (°F) | Reference | Notes |
|------|----------------|-----------------|-----------|-------|
| ST-1 | 0.5 | 79.58 | ASME | Vacuum |
| ST-2 | 1.0 | 101.74 | ASME | Low pressure |
| ST-3 | 14.7 | 212.0 | ASME | Atmospheric |
| ST-4 | 50 | 281.0 | ASME | Moderate |
| ST-5 | 100 | 327.8 | ASME | Common boiler |
| ST-6 | 200 | 381.8 | ASME | High pressure |
| ST-7 | 500 | 467.0 | ASME | Very high |
| ST-8 | 1000 | 544.6 | ASME | Extreme |

**Tolerance:** 0.5%

---

## Transport Properties Test Cases

### 3. Water Density Tests (6 cases)

| Case | Temp (°F) | Pressure (psia) | Expected ρ (lb/ft³) | Notes |
|------|----------|-----------------|---------------------|-------|
| WD-1 | 32 | 14.7 | 62.42 | Freezing |
| WD-2 | 68 | 14.7 | 62.32 | Standard |
| WD-3 | 212 | 14.7 | 59.83 | Boiling |
| WD-4 | 300 | 100 | 57.31 | High temp |
| WD-5 | 68 | 1000 | 62.47 | High pressure effect |
| WD-6 | 180 | 14.7 | 60.59 | Hot water |

**Tolerance:** 0.5%

---

### 4. Water Viscosity Tests (6 cases)

| Case | Temp (°F) | Expected μ (cP) | Notes |
|------|----------|-----------------|-------|
| WV-1 | 32 | 1.787 | Cold water |
| WV-2 | 68 | 1.002 | Standard |
| WV-3 | 100 | 0.682 | Warm |
| WV-4 | 150 | 0.433 | Hot |
| WV-5 | 212 | 0.282 | Boiling |
| WV-6 | 300 | 0.146 | Very hot |

**Tolerance:** 1% (viscosity is more variable)

---

## Thermal Properties Test Cases

### 5. Specific Heat Tests (6 cases)

| Case | Temp (°F) | Expected Cp (BTU/lb·°F) | Notes |
|------|----------|------------------------|-------|
| WSH-1 | 32 | 1.0074 | Freezing |
| WSH-2 | 68 | 0.9988 | Standard |
| WSH-3 | 100 | 0.9980 | Minimum Cp |
| WSH-4 | 150 | 1.0000 | Warm |
| WSH-5 | 212 | 1.0070 | Boiling |
| WSH-6 | 300 | 1.0200 | Hot |

**Tolerance:** 0.5%

---

### 6. Thermal Conductivity Tests (6 cases)

| Case | Temp (°F) | Expected k (BTU/hr·ft·°F) | Notes |
|------|----------|--------------------------|-------|
| WTC-1 | 32 | 0.319 | Cold |
| WTC-2 | 68 | 0.345 | Standard |
| WTC-3 | 100 | 0.362 | Warm |
| WTC-4 | 150 | 0.384 | Hot |
| WTC-5 | 212 | 0.393 | Boiling |
| WTC-6 | 300 | 0.395 | Very hot |

**Tolerance:** 1%

---

## Steam Properties Test Cases

### 7. Steam Enthalpy Tests (12 cases)

#### Saturated Liquid (quality = 0)

| Case | P (psia) | T (°F) | Quality | Expected h (BTU/lb) | Notes |
|------|---------|--------|---------|---------------------|-------|
| SE-1 | 14.7 | 212 | 0.0 | 180.1 | Atm saturated liquid |
| SE-2 | 100 | 327.8 | 0.0 | 298.4 | Moderate pressure |
| SE-3 | 200 | 381.8 | 0.0 | 355.4 | High pressure |

#### Saturated Vapor (quality = 1)

| Case | P (psia) | T (°F) | Quality | Expected h (BTU/lb) | Notes |
|------|---------|--------|---------|---------------------|-------|
| SE-4 | 14.7 | 212 | 1.0 | 1150.4 | Atm saturated vapor |
| SE-5 | 100 | 327.8 | 1.0 | 1187.2 | Moderate pressure |
| SE-6 | 200 | 381.8 | 1.0 | 1198.4 | High pressure |

#### Two-Phase Mixture

| Case | P (psia) | T (°F) | Quality | Expected h (BTU/lb) | Notes |
|------|---------|--------|---------|---------------------|-------|
| SE-7 | 14.7 | 212 | 0.5 | 665.2 | 50% quality |
| SE-8 | 100 | 327.8 | 0.25 | 520.6 | 25% quality |
| SE-9 | 200 | 381.8 | 0.75 | 987.7 | 75% quality |

#### Subcooled Liquid

| Case | P (psia) | T (°F) | Quality | Expected h (BTU/lb) | Notes |
|------|---------|--------|---------|---------------------|-------|
| SE-10 | 14.7 | 150 | 0.0 | 118.0 | 62°F subcooling |

#### Superheated Vapor

| Case | P (psia) | T (°F) | Quality | Expected h (BTU/lb) | Notes |
|------|---------|--------|---------|---------------------|-------|
| SE-11 | 14.7 | 300 | 1.0 | 1195.0 | 88°F superheat |
| SE-12 | 200 | 500 | 1.0 | 1268.8 | 118°F superheat |

**Tolerance:** 1% (steam enthalpy has more complexity)

---

### 8. Steam Quality Tests (8 cases)

| Case | h (BTU/lb) | P (psia) | Expected Quality | Notes |
|------|-----------|---------|------------------|-------|
| SQ-1 | 180.1 | 14.7 | 0.00 | Saturated liquid |
| SQ-2 | 1150.4 | 14.7 | 1.00 | Saturated vapor |
| SQ-3 | 665.2 | 14.7 | 0.50 | 50% quality |
| SQ-4 | 422.6 | 14.7 | 0.25 | 25% quality |
| SQ-5 | 907.8 | 14.7 | 0.75 | 75% quality |
| SQ-6 | 118.0 | 14.7 | <0 | Subcooled (negative quality) |
| SQ-7 | 1200.0 | 14.7 | >1 | Superheated (>1 quality) |
| SQ-8 | 520.6 | 100 | 0.25 | Higher pressure |

**Tolerance:** 2% (quality is calculated, compounds errors)

---

## Implementation Plan

### Week 3 Day 1: Excel Data Extraction

**Tasks:**
1. Open `Engineering-Functions.xlam` in Excel
2. Create test workbook with:
   - Input cells for each test case
   - Formula cells calling VBA functions
   - Output cells with results
3. Run all 60 test cases
4. Export results to CSV/JSON

**Deliverable:** `fluids_validation_data.json`

```json
{
  "test_cases": [
    {
      "id": "SP-1",
      "function": "SaturationPressure",
      "inputs": {"temperature": 32.0},
      "excel_output": 0.089,
      "date_tested": "2025-10-23",
      "excel_version": "Engineering-Functions.xlam v2.1"
    },
    ...
  ]
}
```

---

### Week 3 Day 2: Python Test Implementation

**Tasks:**
1. Create `tests/validation/test_validation_fluids.py`
2. Load test data from JSON
3. Implement test functions using validation framework
4. Run tests, capture initial results

**Code Structure:**

```python
"""
Validation tests for fluids module against Excel VBA.
"""

import pytest
import json
from pathlib import Path
from sigma_thermal.fluids import (
    saturation_pressure,
    saturation_temperature,
    water_density,
    water_viscosity,
    water_specific_heat,
    water_thermal_conductivity,
    steam_enthalpy,
    steam_quality
)

# Load test data
TEST_DATA_FILE = Path(__file__).parent / "data" / "fluids_validation_data.json"

with open(TEST_DATA_FILE) as f:
    TEST_DATA = json.load(f)


class TestSaturationPressureValidation:
    """Validate saturation_pressure() against Excel VBA."""

    @pytest.mark.parametrize("test_case",
        [tc for tc in TEST_DATA["test_cases"] if tc["function"] == "SaturationPressure"],
        ids=lambda tc: tc["id"]
    )
    def test_saturation_pressure_vs_excel(self, test_case):
        """Test saturation_pressure against Excel VBA."""
        # Get inputs
        temp = test_case["inputs"]["temperature"]

        # Python calculation
        python_result = saturation_pressure(temp)

        # Excel result
        excel_result = test_case["excel_output"]

        # Compare
        rel_diff = abs(python_result - excel_result) / excel_result
        tolerance = test_case.get("tolerance", 0.01)

        assert rel_diff <= tolerance, (
            f"Deviation {rel_diff:.4%} exceeds tolerance {tolerance:.2%}\n"
            f"  Python: {python_result:.4f}\n"
            f"  Excel:  {excel_result:.4f}"
        )


# Repeat for all 8 functions...
```

---

### Week 3 Day 3: Analysis & Documentation

**Tasks:**
1. Run full validation test suite
2. Analyze results:
   - Calculate statistics (mean, max, std deviation)
   - Identify any failures
   - Investigate root causes
3. Update `EXCEL_VBA_DISCREPANCIES.md` with findings
4. Create summary report

**Deliverable:** Validation Report

```markdown
## Fluids Module Validation Results

**Test Date:** October 25, 2025
**Tests Run:** 60
**Tests Passed:** 58 (96.7%)
**Tests Failed:** 2 (3.3%)

### Summary Statistics

| Function | Tests | Pass | Fail | Mean Dev | Max Dev |
|----------|-------|------|------|----------|---------|
| saturation_pressure | 8 | 8 | 0 | 0.02% | 0.05% |
| saturation_temperature | 8 | 8 | 0 | 0.03% | 0.08% |
| ... | ... | ... | ... | ... | ... |

### Failures Investigated

**F-1: steam_enthalpy at 200 psia superheated**
- Deviation: 2.3% (exceeds 1% tolerance)
- Root Cause: Excel uses outdated superheated steam correlation
- Resolution: Python is more accurate (matches ASME tables)
```

---

## Excel VBA Function Reference

### Expected VBA Function Signatures

Based on `Engineering-Functions.xlam`:

```vba
' Saturation properties
Function SaturationPressure(Temperature As Double) As Double
Function SaturationTemperature(Pressure As Double) As Double

' Water transport properties
Function WaterDensity(Temperature As Double, Optional Pressure As Double = 14.7) As Double
Function WaterViscosity(Temperature As Double) As Double

' Water thermal properties
Function WaterSpecificHeat(Temperature As Double) As Double
Function WaterThermalConductivity(Temperature As Double) As Double

' Steam properties
Function SteamEnthalpy(Temperature As Double, Pressure As Double, Optional Quality As Double = 1.0) As Double
Function SteamQuality(Enthalpy As Double, Pressure As Double) As Double
```

---

## Automated Test Execution

### Integration with CI/CD

Add to `pyproject.toml`:

```toml
[tool.pytest.ini_options]
markers = [
    "validation: Excel VBA validation tests (deselect with '-m \"not validation\"')",
    "fluids: Fluids module tests",
]
```

Run validation tests:

```bash
# Run all validation tests
pytest tests/validation/ -v

# Run only fluids validation
pytest tests/validation/test_validation_fluids.py -v

# Skip validation tests (for faster CI)
pytest -m "not validation"
```

---

## Success Criteria Checklist

### Must-Have (Required for Release)

- [ ] All 8 functions have Excel validation tests
- [ ] Minimum 60 test cases covering operating range
- [ ] All tests passing (<1% deviation) OR discrepancies documented
- [ ] Results documented in EXCEL_VBA_DISCREPANCIES.md
- [ ] Validation tests integrated into CI/CD

### Nice-to-Have (Future Enhancements)

- [ ] 100+ test cases for comprehensive coverage
- [ ] Automated Excel calling (via xlwings/COM)
- [ ] Real-time comparison in calculator UI
- [ ] Performance benchmarking Python vs Excel
- [ ] Cross-platform validation (Mac/Windows)

---

## Risk Management

### Potential Issues

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Excel VBA functions missing | Low | High | Document in requirements doc |
| Excel file corrupted | Low | Medium | Backup files, version control |
| Large deviations found | Medium | High | Investigate root cause, document |
| Test data extraction slow | Medium | Low | Automate with VBA macro |
| COM/xlwings compatibility | Medium | Medium | Manual extraction if needed |

---

## Timeline

| Day | Tasks | Deliverable | Hours |
|-----|-------|-------------|-------|
| Day 1 | Extract Excel test data | fluids_validation_data.json | 4 |
| Day 2 | Implement validation tests | test_validation_fluids.py | 6 |
| Day 3 | Run tests, analyze, document | Validation report | 4 |

**Total Estimated Time:** 14 hours (2 work days)

---

## Reference Materials

### ASME Steam Tables
- ASME Steam Tables (9th Edition, 2017)
- Online: https://www.asme.org/codes-standards/steam-tables
- Use for validating saturation properties, enthalpy

### Perry's Chemical Engineers' Handbook
- 9th Edition, Chapter 2: Physical Properties
- Water density, viscosity correlations

### NIST Chemistry WebBook
- https://webbook.nist.gov/chemistry/fluid/
- Water/steam properties calculator
- Use for cross-validation

---

*Document Version: 1.0*
*Last Updated: October 22, 2025*
*Next Review: After validation implementation (Week 3 Day 3)*
