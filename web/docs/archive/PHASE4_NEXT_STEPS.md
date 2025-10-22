# Phase 4: Detailed Next Steps & Action Plan

**Phase:** 4 - Fluids Module & Combustion Completion
**Status:** Ready to Start
**Duration:** 2-3 weeks (estimated)
**Start Date:** Week 9
**Priority:** High

---

## Executive Summary

Phase 4 will complete the combustion module (33% remaining) and begin the fluids module, applying the proven validation patterns from Phase 3. This phase focuses on production-ready implementation with comprehensive testing.

**Key Objectives:**
1. Complete remaining combustion functions (10 functions)
2. Implement core fluids module (15-20 functions)
3. Validate all new functions against VBA
4. Maintain 100% test pass rate
5. Expand documentation

---

## Table of Contents

1. [Pre-Phase 4 Checklist](#1-pre-phase-4-checklist)
2. [Combustion Module Completion](#2-combustion-module-completion)
3. [Fluids Module Development](#3-fluids-module-development)
4. [Testing & Validation](#4-testing--validation)
5. [Documentation](#5-documentation)
6. [CI/CD Enhancements](#6-cicd-enhancements)
7. [Timeline & Milestones](#7-timeline--milestones)
8. [Resource Requirements](#8-resource-requirements)
9. [Success Criteria](#9-success-criteria)
10. [Risk Management](#10-risk-management)

---

## 1. Pre-Phase 4 Checklist

### 1.1 Immediate Actions (Before Starting Phase 4)

| Task | Priority | Estimate | Status |
|------|----------|----------|--------|
| Update README.md with Phase 3 status | High | 15 min | 🔲 TODO |
| Fix 2 unit test failures | Medium | 1 hour | 🔲 TODO |
| Review Phase 3 lessons learned | High | 30 min | 🔲 TODO |
| Set up Phase 4 branch | High | 5 min | 🔲 TODO |
| Create Phase 4 project board | Medium | 15 min | 🔲 TODO |

### 1.2 Update README.md

**File:** `README.md`

**Changes Needed:**
```markdown
## Project Status

**Current Phase**: Phase 4 - Fluids Module & Combustion Completion ⏳ **IN PROGRESS**

**Phase 3 Complete:** ✅
- [x] 44 validation and integration tests (100% passing)
- [x] <0.01% accuracy vs Excel VBA
- [x] 5-7x performance improvement
- [x] Comprehensive HTML documentation
- [x] CI/CD pipeline operational

**Metrics:**
- 576 VBA functions inventoried
- 20 functions implemented (3.5%)
- 224 tests created (220 passing, 98.2%)
- 92% test coverage on combustion module
- Validation framework operational
```

### 1.3 Fix Unit Test Failures

**Files to Fix:**
- `tests/unit/test_units.py` (2 failures)

**Issues:**
1. `test_temperature_conversion` - Pint offset units
2. `test_create_quantity` - Pint offset units

**Solution Options:**
- Document as known limitation
- Implement workaround (use magnitude extraction)
- Or defer to future phase if not blocking

**Priority:** Medium (not blocking combustion work)

---

## 2. Combustion Module Completion

### 2.1 Remaining Functions (10 functions, ~33%)

| Function | Category | Priority | Complexity | Estimate |
|----------|----------|----------|------------|----------|
| **Air-Fuel Ratios** (4 functions) |
| stoich_air_mass_gas() | Air-Fuel | High | Low | 2 hours |
| stoich_air_vol_gas() | Air-Fuel | High | Low | 2 hours |
| stoich_air_mass_liquid() | Air-Fuel | High | Low | 2 hours |
| excess_air_percent() | Air-Fuel | High | Low | 2 hours |
| **Flame Temperature** (2 functions) |
| adiabatic_flame_temp_gas() | Temperature | Medium | Medium | 4 hours |
| adiabatic_flame_temp_liquid() | Temperature | Medium | Medium | 4 hours |
| **Efficiency Helpers** (2 functions) |
| combustion_efficiency() | Efficiency | High | Low | 3 hours |
| stack_loss_percent() | Efficiency | High | Low | 2 hours |
| **Emissions** (2 functions) |
| nox_emissions() | Emissions | Medium | Medium | 4 hours |
| co2_emissions_per_mmbtu() | Emissions | High | Low | 2 hours |

**Total Estimated Time:** 27 hours (~3-4 days)

### 2.2 Implementation Plan

#### Week 1: Air-Fuel Ratios & Efficiency

**Day 1-2: Air-Fuel Ratio Functions**

1. **stoich_air_mass_gas()**
   ```python
   def stoich_air_mass_gas(
       composition: GasCompositionMass
   ) -> float:
       """Calculate stoichiometric air required (lb air / lb fuel)"""
   ```
   - Implement mass-based stoichiometric calculation
   - Use lookup tables for component air requirements
   - Test with pure fuels (CH4, C2H6, C3H8)
   - Test with natural gas mixtures
   - Validate against VBA (tolerance <0.1%)

2. **stoich_air_vol_gas()**
   ```python
   def stoich_air_vol_gas(
       composition: GasCompositionVolume
   ) -> float:
       """Calculate stoichiometric air required (scf air / scf fuel)"""
   ```
   - Implement volume-based calculation
   - Test with typical gas compositions
   - Validate against VBA

3. **stoich_air_mass_liquid()**
   ```python
   def stoich_air_mass_liquid(
       fuel_type: str
   ) -> float:
       """Calculate stoichiometric air for liquid fuel"""
   ```
   - Use lookup tables for #1-#6 oil, gasoline
   - Test all fuel types
   - Validate against VBA

4. **excess_air_percent()**
   ```python
   def excess_air_percent(
       actual_air: float,
       stoich_air: float
   ) -> float:
       """Calculate excess air percentage"""
   ```
   - Simple calculation: (actual - stoich) / stoich * 100
   - Test edge cases (0%, negative, very high)
   - Validate with known scenarios

**Day 3-4: Efficiency Helpers**

5. **combustion_efficiency()**
   ```python
   def combustion_efficiency(
       heat_input: float,
       stack_loss: float,
       radiation_loss: float = 0.0,
       blow_down_loss: float = 0.0
   ) -> float:
       """Calculate combustion efficiency (%)"""
   ```
   - Implement ASME PTC 4 method
   - Test with realistic loss scenarios
   - Validate against manual calculations
   - Add integration tests with complete workflows

6. **stack_loss_percent()**
   ```python
   def stack_loss_percent(
       flue_gas_enthalpy: float,
       flue_gas_flow: float,
       heat_input: float
   ) -> float:
       """Calculate stack loss as % of heat input"""
   ```
   - Simple calculation: (enthalpy * flow) / heat_input * 100
   - Test with various stack temperatures
   - Validate with Phase 3 test cases

#### Week 2: Flame Temperature & Emissions

**Day 5-6: Flame Temperature**

7. **adiabatic_flame_temp_gas()**
   ```python
   def adiabatic_flame_temp_gas(
       composition: GasCompositionMass,
       air_temp: float,
       fuel_temp: float,
       excess_air_percent: float
   ) -> float:
       """Calculate adiabatic flame temperature (°F)"""
   ```
   - Implement iterative energy balance
   - Use enthalpy functions from Phase 3
   - Test with various excess air levels
   - Validate against published data (GPSA)

8. **adiabatic_flame_temp_liquid()**
   ```python
   def adiabatic_flame_temp_liquid(
       fuel_type: str,
       air_temp: float,
       fuel_temp: float,
       excess_air_percent: float
   ) -> float:
       """Calculate adiabatic flame temperature for liquid fuel"""
   ```
   - Similar to gas, but use liquid fuel properties
   - Test with #2 oil, gasoline
   - Validate against VBA

**Day 7: Emissions**

9. **nox_emissions()**
   ```python
   def nox_emissions(
       flame_temp: float,
       residence_time: float,
       oxygen_percent: float
   ) -> float:
       """Estimate NOx emissions (ppm or lb/MMBtu)"""
   ```
   - Implement Zeldovich mechanism approximation
   - Or use empirical correlation from AP-42
   - Test with typical industrial conditions
   - Document assumptions and limitations

10. **co2_emissions_per_mmbtu()**
    ```python
    def co2_emissions_per_mmbtu(
        composition: GasCompositionMass,
        hhv: float
    ) -> float:
        """Calculate CO2 emissions (lb CO2 / MMBtu)"""
    ```
    - Use POC functions from Phase 3
    - Test with various fuels
    - Validate against EPA factors

### 2.3 Testing Requirements

**For Each Function:**
- ✅ Unit tests (5-10 tests per function)
- ✅ Edge case tests
- ✅ Validation against VBA (tolerance <0.5%)
- ✅ Integration with existing functions
- ✅ Type hints and docstrings

**Total New Tests:** ~60 tests (10 functions × 6 tests average)

---

## 3. Fluids Module Development

### 3.1 Core Functions (15-20 functions)

| Function | Category | Priority | Complexity | Estimate |
|----------|----------|----------|------------|----------|
| **Water/Steam Properties** (8 functions) |
| water_density() | Properties | High | Low | 2 hours |
| water_viscosity() | Properties | High | Low | 2 hours |
| water_thermal_conductivity() | Properties | Medium | Low | 2 hours |
| water_specific_heat() | Properties | High | Low | 2 hours |
| steam_enthalpy() | Properties | High | Medium | 3 hours |
| steam_entropy() | Properties | Medium | Medium | 3 hours |
| saturation_pressure() | Properties | High | Low | 2 hours |
| saturation_temperature() | Properties | High | Low | 2 hours |
| **Psychrometric** (4 functions) |
| relative_humidity() | Psychrometric | Medium | Low | 2 hours |
| dew_point() | Psychrometric | Medium | Low | 2 hours |
| wet_bulb_temp() | Psychrometric | Low | Medium | 3 hours |
| humidity_ratio() | Psychrometric | Medium | Low | 2 hours |
| **Thermal Fluids** (4 functions) |
| thermal_fluid_density() | Properties | High | Low | 2 hours |
| thermal_fluid_viscosity() | Properties | High | Low | 2 hours |
| thermal_fluid_specific_heat() | Properties | High | Low | 2 hours |
| thermal_fluid_thermal_conductivity() | Properties | Medium | Low | 2 hours |

**Total Estimated Time:** 35 hours (~5 days)

### 3.2 Implementation Strategy

#### Week 2-3: Water/Steam Properties

**Use CoolProp Interface:**
```python
from CoolProp.CoolProp import PropsSI

def water_density(
    temperature: float,
    pressure: float,
    phase: str = "liquid"
) -> float:
    """
    Calculate water density using CoolProp.

    Args:
        temperature: Temperature (°F)
        pressure: Pressure (psia)
        phase: "liquid" or "vapor"

    Returns:
        Density (lb/ft³)
    """
    # Convert to SI units
    T_K = (temperature - 32) * 5/9 + 273.15
    P_Pa = pressure * 6894.76

    # Call CoolProp
    rho_kg_m3 = PropsSI('D', 'T', T_K, 'P', P_Pa, 'Water')

    # Convert to Imperial
    rho_lb_ft3 = rho_kg_m3 * 0.062428

    return rho_lb_ft3
```

**Testing:**
- Unit tests with known values
- Compare to steam tables
- Validate against VBA
- Test edge cases (near saturation)

#### Week 3: Psychrometric & Thermal Fluids

**Psychrometric Functions:**
- Use standard psychrometric equations
- Implement iterative solvers where needed
- Validate against psychrometric charts

**Thermal Fluids:**
- Implement correlations for common fluids (Therminol, Dowtherm)
- Use polynomial fits from manufacturer data
- Test with temperature ranges from datasheets

### 3.3 Data Requirements

**Thermal Fluid Property Data:**
- Create `data/thermal_fluids.json`
- Include coefficients for:
  - Therminol VP-1
  - Dowtherm A
  - Marlotherm SH
  - Paratherm NF
- Source data from manufacturer datasheets

**Format:**
```json
{
  "Therminol VP-1": {
    "density": {
      "coefficients": [a, b, c, d],
      "units": "lb/ft³",
      "temp_range": [0, 750]
    },
    "viscosity": {...},
    "specific_heat": {...},
    "thermal_conductivity": {...}
  }
}
```

---

## 4. Testing & Validation

### 4.1 Test Targets

**Phase 4 Test Goals:**
- Total tests: 300+ (current 224 + 60 combustion + 30 fluids)
- Pass rate: 100% (no new failures)
- Coverage: >90% on all new modules
- Validation: <0.5% difference from VBA

### 4.2 Validation Test Cases

**New Validation Tests:**

**Test Case 4: Complete Combustion Workflow**
- Scenario: Natural gas boiler with efficiency calculation
- Functions: All combustion functions end-to-end
- Validation: Match VBA complete workflow
- Tests: 15 comprehensive tests

**Test Case 5: Water Properties Validation**
- Scenario: Water/steam at various T & P
- Functions: All water property functions
- Validation: Match steam tables
- Tests: 10 tests

**Test Case 6: Thermal Fluid Application**
- Scenario: Therminol VP-1 in heat transfer loop
- Functions: Thermal fluid properties
- Validation: Match manufacturer data
- Tests: 8 tests

### 4.3 Integration Tests

**New Integration Tests:**

**Test Suite: Boiler Efficiency with Air-Fuel**
- Test complete calculation including air requirements
- Validate air-fuel ratios
- Test excess air impact
- Tests: 6 tests

**Test Suite: Fluid Heating Duty**
- Test fluid property calculations in heating application
- Validate heat transfer calculations
- Test fluid selection logic
- Tests: 5 tests

---

## 5. Documentation

### 5.1 Code Documentation

**Requirements:**
- ✅ All functions have docstrings
- ✅ All parameters documented with units
- ✅ Return types documented
- ✅ Usage examples in docstrings
- ✅ References to source methods

### 5.2 User Documentation Updates

**Update getting_started.html:**
- Add section on fluids module
- Add examples for water properties
- Add thermal fluid examples
- Add complete system examples

**Update validation_results.html:**
- Add Phase 4 validation results
- Update test counts and coverage
- Add fluids validation section

### 5.3 API Documentation

**Create Sphinx Documentation:**
```bash
# Setup
pip install sphinx sphinx-rtd-theme sphinx-autodoc-typehints

# Initialize
cd docs
sphinx-quickstart

# Configure
# Edit conf.py with project settings

# Build
make html
```

**Structure:**
```
docs/
├── source/
│   ├── index.rst              # Landing page
│   ├── getting_started.rst    # Quick start
│   ├── api/
│   │   ├── combustion.rst     # Combustion API
│   │   ├── fluids.rst         # Fluids API
│   │   └── engineering.rst    # Engineering utilities
│   ├── examples/
│   │   ├── boiler.rst         # Boiler example
│   │   └── fluid_heating.rst  # Fluid heating example
│   └── validation.rst         # Validation results
└── build/
    └── html/                   # Generated docs
```

### 5.4 Tutorial Notebooks

**Create Jupyter Notebooks:**

**Tutorial 1: Basic Combustion Calculations**
- Load module
- Calculate heating values
- Calculate POC
- Calculate efficiency
- 30-45 minutes

**Tutorial 2: Water/Steam Properties**
- Use CoolProp interface
- Calculate properties at various conditions
- Plot saturation curve
- 30 minutes

**Tutorial 3: Complete Heater Design**
- Define fuel and fluid
- Calculate combustion
- Size heat exchanger
- Calculate duty
- 60 minutes

---

## 6. CI/CD Enhancements

### 6.1 Add Dependabot

**File:** `.github/dependabot.yml`

```yaml
version: 2
updates:
  - package-ecosystem: "pip"
    directory: "/"
    schedule:
      interval: "weekly"
    open-pull-requests-limit: 10
    labels:
      - "dependencies"
```

### 6.2 Add Documentation Building

**File:** `.github/workflows/docs.yml`

```yaml
name: Documentation

on:
  push:
    branches: [main, develop]
  pull_request:
    branches: [main]

jobs:
  build-docs:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
        with:
          python-version: "3.11"
      - name: Install dependencies
        run: |
          pip install -r requirements.txt
          pip install sphinx sphinx-rtd-theme
      - name: Build docs
        run: |
          cd docs
          make html
      - name: Upload artifacts
        uses: actions/upload-artifact@v3
        with:
          name: documentation
          path: docs/build/html/
```

### 6.3 Add Performance Testing

**File:** `.github/workflows/performance.yml`

```yaml
name: Performance Tests

on:
  push:
    branches: [main]

jobs:
  benchmark:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-python@v5
      - name: Install dependencies
        run: |
          pip install -r requirements-dev.txt
          pip install pytest-benchmark
      - name: Run benchmarks
        run: |
          pytest tests/performance/ --benchmark-only
      - name: Upload results
        uses: actions/upload-artifact@v3
        with:
          name: benchmarks
          path: .benchmarks/
```

---

## 7. Timeline & Milestones

### 7.1 Phase 4 Schedule (3 weeks)

#### Week 1 (Days 1-5)
**Focus:** Combustion completion & setup

| Day | Task | Deliverable |
|-----|------|-------------|
| Mon | Pre-phase checklist, setup | Branch created, plan reviewed |
| Tue | Air-fuel ratio functions (4) | Functions + tests (40) |
| Wed | Efficiency helper functions (2) | Functions + tests (20) |
| Thu | Flame temperature functions (2) | Functions + tests (20) |
| Fri | Emissions functions (2) | Functions + tests (20) |

**Week 1 Milestone:** Combustion module 100% complete ✅

#### Week 2 (Days 6-10)
**Focus:** Fluids module - water/steam

| Day | Task | Deliverable |
|-----|------|-------------|
| Mon | CoolProp integration setup | Interface tested |
| Tue | Water properties (4 functions) | Functions + tests (40) |
| Wed | Steam properties (4 functions) | Functions + tests (40) |
| Thu | Validation test case 4 & 5 | 25 validation tests |
| Fri | Integration tests | 11 integration tests |

**Week 2 Milestone:** Water/steam properties complete ✅

#### Week 3 (Days 11-15)
**Focus:** Fluids module - psychrometric & thermal fluids

| Day | Task | Deliverable |
|-----|------|-------------|
| Mon | Psychrometric functions (4) | Functions + tests (30) |
| Tue | Thermal fluid functions (4) | Functions + tests (30) |
| Wed | Validation test case 6 | 8 validation tests |
| Thu | Documentation (Sphinx, notebooks) | API docs + 1 tutorial |
| Fri | CI/CD enhancements, final review | Dependabot, docs workflow |

**Week 3 Milestone:** Phase 4 complete ✅

### 7.2 Key Milestones

| Milestone | Date | Criteria |
|-----------|------|----------|
| Combustion 100% Complete | End Week 1 | 30 functions, 241 tests |
| Water/Steam Properties | End Week 2 | 8 functions, 281 tests |
| Fluids Module Complete | End Week 3 | 16 functions, 311 tests |
| Documentation Complete | End Week 3 | Sphinx docs, 1 tutorial |
| Phase 4 Complete | Day 15 | All criteria met |

---

## 8. Resource Requirements

### 8.1 Development Resources

**Time Commitment:**
- Developer time: 3 weeks full-time (120 hours)
  - Coding: 70 hours (58%)
  - Testing: 30 hours (25%)
  - Documentation: 15 hours (12%)
  - Review/debugging: 5 hours (5%)

**Tools & Services:**
- GitHub (repository, CI/CD)
- CoolProp (free, open source)
- Sphinx (documentation)
- Jupyter (tutorials)

### 8.2 Data Requirements

**Required Data:**
- CoolProp (already available)
- Thermal fluid property data (collect from manufacturer sites)
- Steam tables (for validation)
- Psychrometric chart data (for validation)

**Data Collection:** 2-3 hours

### 8.3 Testing Resources

**Validation Data:**
- VBA Excel files (already available)
- Steam tables (ASME Steam Tables or NIST)
- Thermal fluid datasheets (Therminol, Dowtherm)

**Testing Time:** Included in 30-hour testing estimate

---

## 9. Success Criteria

### 9.1 Functional Criteria

**Combustion Module:**
- ✅ 30/30 functions implemented (100%)
- ✅ All functions tested (>90% coverage)
- ✅ All functions validated (<0.5% vs VBA)
- ✅ All edge cases handled
- ✅ Complete documentation

**Fluids Module:**
- ✅ 16+ core functions implemented
- ✅ CoolProp interface working
- ✅ All functions tested (>90% coverage)
- ✅ Validated against steam tables & datasheets
- ✅ Complete documentation

### 9.2 Quality Criteria

**Code Quality:**
- ✅ 100% type hints
- ✅ 100% docstrings
- ✅ Black formatted
- ✅ Ruff passing
- ✅ mypy passing (with pragmas if needed)

**Testing:**
- ✅ 300+ total tests
- ✅ 100% pass rate
- ✅ >90% coverage on new modules
- ✅ <0.5% validation accuracy

**Documentation:**
- ✅ Sphinx API documentation
- ✅ At least 1 Jupyter tutorial
- ✅ Updated HTML user guides
- ✅ Updated validation results

### 9.3 Process Criteria

**CI/CD:**
- ✅ All workflows passing
- ✅ Dependabot configured
- ✅ Documentation building automated
- ✅ Performance tests added

**Project Management:**
- ✅ All tasks tracked in project board
- ✅ Daily progress updates
- ✅ Blockers identified and resolved
- ✅ Phase 4 completion report written

---

## 10. Risk Management

### 10.1 Technical Risks

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| CoolProp integration issues | Medium | High | Test early, have fallback correlations |
| VBA validation failures | Low | High | Use Phase 3 patterns, test incrementally |
| Performance regression | Low | Medium | Run performance tests in CI |
| Flame temp convergence issues | Medium | Medium | Use robust iterative solver, test edge cases |

### 10.2 Schedule Risks

| Risk | Probability | Impact | Mitigation |
|------|------------|--------|------------|
| Underestimated complexity | Medium | Medium | Buffer time built in, prioritize core functions |
| Blocked by data availability | Low | Low | Collect data early in phase |
| Testing takes longer | Medium | Low | Parallel testing, focus on critical paths |

### 10.3 Contingency Plans

**If Behind Schedule:**
1. Prioritize combustion completion
2. Defer psychrometric functions to Phase 5
3. Defer thermal fluids to Phase 5
4. Focus on water/steam only

**If Technical Blockers:**
1. Document issue clearly
2. Implement workaround if possible
3. Defer to future phase if not critical
4. Continue with unblocked work

---

## 11. Action Items Summary

### 11.1 Immediate (This Week)

| # | Task | Owner | Due | Status |
|---|------|-------|-----|--------|
| 1 | Update README.md | Dev | Oct 23 | 🔲 TODO |
| 2 | Review Phase 3 lessons | Dev | Oct 23 | 🔲 TODO |
| 3 | Create Phase 4 branch | Dev | Oct 23 | 🔲 TODO |
| 4 | Set up project board | PM | Oct 23 | 🔲 TODO |
| 5 | Collect thermal fluid data | Dev | Oct 24 | 🔲 TODO |

### 11.2 Week 1 Deliverables

| # | Deliverable | Tests | Coverage | Due |
|---|-------------|-------|----------|-----|
| 1 | Air-fuel functions (4) | 40 | >90% | Oct 25 |
| 2 | Efficiency functions (2) | 20 | >90% | Oct 26 |
| 3 | Flame temp functions (2) | 20 | >90% | Oct 27 |
| 4 | Emissions functions (2) | 20 | >90% | Oct 28 |

### 11.3 Week 2 Deliverables

| # | Deliverable | Tests | Coverage | Due |
|---|-------------|-------|----------|-----|
| 1 | Water properties (4) | 40 | >90% | Nov 1 |
| 2 | Steam properties (4) | 40 | >90% | Nov 2 |
| 3 | Validation tests | 25 | - | Nov 3 |
| 4 | Integration tests | 11 | - | Nov 4 |

### 11.4 Week 3 Deliverables

| # | Deliverable | Tests | Coverage | Due |
|---|-------------|-------|----------|-----|
| 1 | Psychrometric (4) | 30 | >90% | Nov 7 |
| 2 | Thermal fluids (4) | 30 | >90% | Nov 8 |
| 3 | Sphinx docs | - | - | Nov 9 |
| 4 | Jupyter tutorial | - | - | Nov 9 |
| 5 | CI/CD enhancements | - | - | Nov 10 |

---

## 12. Appendices

### 12.1 Function Signatures Reference

**Combustion Module (Remaining):**

```python
# Air-Fuel Ratios
def stoich_air_mass_gas(composition: GasCompositionMass) -> float: ...
def stoich_air_vol_gas(composition: GasCompositionVolume) -> float: ...
def stoich_air_mass_liquid(fuel_type: str) -> float: ...
def excess_air_percent(actual_air: float, stoich_air: float) -> float: ...

# Flame Temperature
def adiabatic_flame_temp_gas(
    composition: GasCompositionMass,
    air_temp: float,
    fuel_temp: float,
    excess_air_percent: float
) -> float: ...

def adiabatic_flame_temp_liquid(
    fuel_type: str,
    air_temp: float,
    fuel_temp: float,
    excess_air_percent: float
) -> float: ...

# Efficiency
def combustion_efficiency(
    heat_input: float,
    stack_loss: float,
    radiation_loss: float = 0.0,
    blow_down_loss: float = 0.0
) -> float: ...

def stack_loss_percent(
    flue_gas_enthalpy: float,
    flue_gas_flow: float,
    heat_input: float
) -> float: ...

# Emissions
def nox_emissions(
    flame_temp: float,
    residence_time: float,
    oxygen_percent: float
) -> float: ...

def co2_emissions_per_mmbtu(
    composition: GasCompositionMass,
    hhv: float
) -> float: ...
```

**Fluids Module (Core):**

```python
# Water/Steam Properties
def water_density(temperature: float, pressure: float, phase: str = "liquid") -> float: ...
def water_viscosity(temperature: float, pressure: float, phase: str = "liquid") -> float: ...
def water_thermal_conductivity(temperature: float, pressure: float) -> float: ...
def water_specific_heat(temperature: float, pressure: float) -> float: ...
def steam_enthalpy(temperature: float, pressure: float) -> float: ...
def steam_entropy(temperature: float, pressure: float) -> float: ...
def saturation_pressure(temperature: float) -> float: ...
def saturation_temperature(pressure: float) -> float: ...

# Psychrometric
def relative_humidity(dry_bulb: float, wet_bulb: float, pressure: float) -> float: ...
def dew_point(temperature: float, relative_humidity: float) -> float: ...
def wet_bulb_temp(temperature: float, relative_humidity: float) -> float: ...
def humidity_ratio(temperature: float, relative_humidity: float) -> float: ...

# Thermal Fluids
def thermal_fluid_density(fluid: str, temperature: float) -> float: ...
def thermal_fluid_viscosity(fluid: str, temperature: float) -> float: ...
def thermal_fluid_specific_heat(fluid: str, temperature: float) -> float: ...
def thermal_fluid_thermal_conductivity(fluid: str, temperature: float) -> float: ...
```

### 12.2 References

**Technical References:**
- GPSA Engineering Data Book, 13th Edition
- ASME PTC 4 (Fired Steam Generators)
- ASME Steam Tables
- CoolProp Documentation
- EPA AP-42 (Emissions Factors)
- Therminol VP-1 Technical Bulletin
- Dowtherm A Product Data Sheet

**Project References:**
- Phase 3 Completion Summary
- Deployment Readiness Audit
- Migration Plan
- Developer Guide (CLAUDE.md)

---

**Document Prepared By:** Claude Code (AI Assistant)
**Date:** October 22, 2025
**Status:** Ready for Phase 4 kickoff
**Approval:** Pending project lead review

---

## Document Status

✅ **COMPLETE & READY FOR REVIEW**

This document provides a comprehensive, actionable plan for Phase 4. All tasks are defined with estimates, priorities, and success criteria.

**Next Action:** Review and approve to begin Phase 4 implementation.
