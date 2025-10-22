# Phase 4 Week 2 Implementation Plan

**Status:** ✅ Week 1 Complete - Ahead of Schedule
**Current:** Ready to begin Fluids Module
**Date:** October 22, 2025

---

## Week 1 Accomplishments

### ✅ COMPLETED AHEAD OF SCHEDULE

Original Week 1-2 plan called for:
- Days 1-2: Air-fuel ratios (4 functions) ✅ DONE
- Days 3-4: Efficiency helpers (4 functions) ✅ DONE
- Days 5-6: Flame temperature (2 functions) ✅ DONE
- Day 7: Emissions (2 functions) ✅ DONE

**Result:** Completed 12 functions in 5 days instead of 7 days planned!

**Combustion Module: 35/35 functions (100% COMPLETE)** 🎉

### Quality Metrics Achieved

| Metric | Target | Actual | Status |
|--------|--------|--------|--------|
| Functions | 12 | 12 | ✅ 100% |
| Tests | 120+ | 160 | ✅ 133% |
| Test Pass Rate | 100% | 100% | ✅ |
| Code Coverage | >90% | 97% avg | ✅ |
| VBA Compatibility | Yes | Yes | ✅ |

---

## Phase 4 Remaining Objectives

### Primary Goal: Fluids Module

The fluids module is the next major deliverable. It provides water/steam properties, psychrometric calculations, and thermal fluid properties essential for:
- Heat exchanger design
- Steam system calculations
- HVAC applications
- Thermal oil system design

---

## Week 2 Plan: Fluids Module - Part 1

### Objective

Implement core water/steam property functions (8-10 functions) with comprehensive testing and validation.

### Priority Functions (Days 6-10)

#### Day 6: Saturation Properties (2 functions)

**1. `saturation_pressure(temperature: float) -> float`**
- Calculate saturation pressure from temperature
- Use Antoine equation or IAPWS-IF97 correlation
- Range: 32°F to 705°F (0°C to 374°C)
- Validation against steam tables

**2. `saturation_temperature(pressure: float) -> float`**
- Calculate saturation temperature from pressure
- Inverse of saturation pressure
- Range: 0.09 psia to 3200 psia
- Validation against steam tables

**Deliverables:**
- 2 functions with full error handling
- ~15 unit tests covering full range
- Integration tests with realistic conditions
- Cross-validation (T→P→T should return original T)

**Estimated Time:** 6-8 hours

---

#### Day 7: Water Density & Viscosity (2 functions)

**3. `water_density(temperature: float, pressure: float = 14.7) -> float`**
- Liquid water density (lb/ft³)
- Temperature-dependent correlation
- Pressure correction for high pressures
- Range: 32°F to 400°F

**4. `water_viscosity(temperature: float) -> float`**
- Dynamic viscosity of liquid water (cP or lb/(ft·s))
- Andrade equation or polynomial fit
- Range: 32°F to 400°F
- Critical for pressure drop calculations

**Deliverables:**
- 2 functions with industry-standard correlations
- ~15 unit tests
- Validation against published data (Perry's, CRC)
- VBA compatibility wrappers

**Estimated Time:** 6-8 hours

---

#### Day 8: Specific Heat & Thermal Conductivity (2 functions)

**5. `water_specific_heat(temperature: float) -> float`**
- Specific heat capacity of liquid water (BTU/(lb·°F))
- Temperature-dependent polynomial
- Range: 32°F to 400°F
- Essential for heat duty calculations

**6. `water_thermal_conductivity(temperature: float) -> float`**
- Thermal conductivity (BTU/(hr·ft·°F))
- Temperature correlation
- Range: 32°F to 400°F
- Used in heat transfer calculations

**Deliverables:**
- 2 functions with validated correlations
- ~15 unit tests
- Integration examples (heat duty calculations)
- Documentation with references

**Estimated Time:** 6-8 hours

---

#### Day 9: Steam Enthalpy & Quality (2 functions)

**7. `steam_enthalpy(temperature: float, pressure: float, quality: float = 1.0) -> float`**
- Enthalpy of steam/water mixture (BTU/lb)
- Handles compressed liquid, saturated mixture, superheated vapor
- Quality: 0 = saturated liquid, 1 = saturated vapor
- IAPWS-IF97 correlations or simplified model

**8. `steam_quality(enthalpy: float, pressure: float) -> float`**
- Calculate quality from enthalpy and pressure
- Returns 0-1 (or <0 for compressed liquid, >1 for superheated)
- Inverse of steam_enthalpy
- Used in steam system analysis

**Deliverables:**
- 2 functions with phase determination logic
- ~20 unit tests (multiple phases)
- Validation against steam tables
- Integration with existing enthalpy functions

**Estimated Time:** 8-10 hours

---

#### Day 10: Integration & Documentation (1 day)

**Tasks:**
- Create integration examples (boiler, heat exchanger)
- Update `fluids/__init__.py` with all exports
- Comprehensive documentation in module docstrings
- README update with fluids module examples
- Cross-validation between functions

**Deliverables:**
- Complete fluids module structure
- ~70-80 total tests
- 100% test pass rate
- Documentation examples
- VBA compatibility confirmed

**Estimated Time:** 6-8 hours

---

## Week 2 Success Criteria

| Criterion | Target |
|-----------|--------|
| Functions implemented | 8 |
| Tests created | 70-80 |
| Test pass rate | 100% |
| Code coverage | >90% |
| VBA compatibility | All functions |
| Documentation | Complete with examples |

---

## Technical Approach

### Data Sources

1. **IAPWS-IF97 Standard** (International Association for the Properties of Water and Steam)
   - Industry standard for steam properties
   - High accuracy across full range
   - May be complex - consider simplified correlations first

2. **Simplified Correlations** (for initial implementation)
   - Antoine equation for saturation pressure
   - Polynomial fits for density, viscosity
   - Interpolation from steam tables for enthalpy
   - Easier to implement, good accuracy for typical ranges

3. **Validation Data**
   - ASME Steam Tables
   - Perry's Chemical Engineers' Handbook
   - CRC Handbook of Chemistry and Physics
   - NIST WebBook

### Implementation Strategy

**Phase 1: Basic Functions (Week 2)**
- Use simplified correlations for speed
- Cover typical industrial ranges (0-400°F, 0-200 psia)
- Focus on accuracy over extreme conditions
- Validate against standard references

**Phase 2: Enhanced Accuracy (Future)**
- Implement full IAPWS-IF97 if needed
- Extend to supercritical conditions
- Add more fluid types
- Performance optimization

### Error Handling

All functions should validate:
- Temperature/pressure within valid range
- Physical validity (e.g., quality 0-1 for two-phase)
- Consistent units (document assumptions)
- Meaningful error messages

---

## File Structure

```
src/sigma_thermal/fluids/
├── __init__.py                    # Module exports
├── water_properties.py            # Water/steam functions (8 functions)
├── psychrometric.py               # Psychrometric functions (future)
└── thermal_fluids.py              # Thermal oil properties (future)

tests/unit/
├── test_fluids_water_properties.py   # ~70-80 tests
├── test_fluids_integration.py        # Integration tests
```

---

## Dependencies & Tools

### Python Libraries

```python
# Already available
import numpy as np
import pytest

# May need for advanced correlations
from scipy.optimize import fsolve  # For inverse functions
from scipy.interpolate import interp1d  # For table interpolation
```

### Reference Data

Consider creating lookup tables from steam tables:
- `data/steam_tables_saturation.csv`
- `data/water_properties_temperature.csv`

Or embed small tables directly in code for critical points.

---

## Risk Management

### Potential Challenges

| Risk | Impact | Probability | Mitigation |
|------|--------|-------------|------------|
| IAPWS-IF97 complexity | High | Medium | Start with simplified correlations |
| Accuracy at extremes | Medium | Low | Define and document valid ranges |
| Phase determination logic | Medium | Medium | Use clear decision tree |
| Integration with combustion | Low | Low | Proven patterns from Phase 3 |

### Contingency Plan

If IAPWS-IF97 proves too complex:
1. Use CoolProp library (external dependency)
2. OR use simplified correlations with documented limitations
3. OR create lookup tables with interpolation

**Recommendation:** Start with simplified correlations, add IAPWS-IF97 later if needed.

---

## After Week 2

### Week 3 Options

**Option A: Complete Fluids Module**
- Psychrometric functions (4 functions)
- Thermal fluid properties (4 functions)
- Advanced steam functions (superheated, compressed liquid)

**Option B: Heat Transfer Module**
- Begin heat transfer coefficients
- Radiation heat transfer
- Convection correlations

**Option C: Validation & Optimization**
- Comprehensive validation against VBA
- Performance profiling
- Documentation enhancement
- Deployment preparation

**Recommendation:** Option A - Complete the fluids module for a cohesive deliverable, then move to heat transfer or validation based on project priorities.

---

## Long-Term Roadmap (Phase 4 Complete)

### Remaining Modules (Priority Order)

1. **Fluids Module** ← Week 2 starts here
   - Water/steam (Week 2: 8 functions)
   - Psychrometric (Week 3: 4 functions)
   - Thermal fluids (Week 3: 4 functions)
   - **Total:** 16 functions

2. **Heat Transfer Module** (Week 4-5)
   - Convection correlations (6 functions)
   - Radiation (4 functions)
   - Heat exchanger design (6 functions)
   - **Total:** 16 functions

3. **Process Calculations** (Week 6)
   - Pressure drop (4 functions)
   - Flow measurement (4 functions)
   - Pump/compressor (4 functions)
   - **Total:** 12 functions

4. **Validation & Polish** (Week 7-8)
   - Comprehensive VBA validation
   - Performance optimization
   - Documentation finalization
   - Deployment package

---

## Success Metrics (Phase 4 Overall)

| Metric | Current | Week 2 Target | Phase 4 Target |
|--------|---------|---------------|----------------|
| **Combustion Functions** | 35/35 (100%) | N/A | 35/35 (100%) |
| **Fluids Functions** | 0 | 8 | 16 |
| **Heat Transfer Functions** | 0 | 0 | 16 |
| **Total Functions** | 35 | 43 | 67+ |
| **Test Coverage** | 97% | >90% | >90% |
| **Tests Passing** | 297/297 | All | All |

---

## Immediate Next Steps (Day 6)

### 1. Research Water/Steam Correlations
- Review Antoine equation for saturation pressure
- Find polynomial fits for density, viscosity
- Identify validation data sources
- Determine accuracy requirements

### 2. Set Up Fluids Module Structure
```bash
# Create module structure
mkdir -p src/sigma_thermal/fluids
touch src/sigma_thermal/fluids/__init__.py
touch src/sigma_thermal/fluids/water_properties.py

# Create test structure
touch tests/unit/test_fluids_water_properties.py
```

### 3. Implement First Function
Start with `saturation_pressure()` as it's:
- Well-documented (Antoine equation)
- Foundational for other functions
- Easy to validate
- Clear requirements

### 4. Establish Testing Pattern
- Create test fixtures for common conditions
- Set up validation data
- Define accuracy tolerances
- Document test methodology

---

## Questions to Resolve

1. **Accuracy vs Complexity Trade-off**
   - Start with simplified correlations?
   - OR implement full IAPWS-IF97 from the start?
   - **Recommendation:** Simplified first, enhance later if needed

2. **Unit Conventions**
   - Use US customary (°F, psia, lb/ft³)?
   - OR SI units (°C, Pa, kg/m³) with conversion?
   - **Recommendation:** US customary (matches existing code)

3. **External Dependencies**
   - Allow CoolProp dependency for accuracy?
   - OR keep self-contained with correlations?
   - **Recommendation:** Self-contained correlations

4. **Validation Depth**
   - Test every 10°F increment?
   - OR test key points + edge cases?
   - **Recommendation:** Key points (32, 212, 400°F) + edges

---

## Resources

### Reference Materials
- **ASME Steam Tables** (definitive source)
- **Perry's Chemical Engineers' Handbook** (Chapter 2: Physical Properties)
- **CRC Handbook** (water properties tables)
- **NIST Chemistry WebBook** (online validation)
- **Crane TP-410** (flow of fluids, practical data)

### Python Libraries
- **CoolProp** (external, very accurate, consider for future)
- **iapws** (Python IAPWS-IF97 implementation)
- **scipy** (interpolation, optimization)

### Validation Tools
- Excel steam table add-ins
- Online steam calculators (NIST, TLV)
- VBA functions from Engineering-Functions.xlam

---

## Conclusion

**We are ahead of schedule and ready to begin the fluids module!**

The combustion module is complete and production-ready. The fluids module represents the next significant value add, providing essential property calculations for thermal system design.

**Week 2 will deliver 8 core water/steam property functions with ~80 comprehensive tests, maintaining our high quality standards.**

---

**Document Version:** 1.0
**Date:** October 22, 2025
**Status:** Ready to Begin Week 2
**Next Action:** Set up fluids module structure and begin saturation properties implementation
