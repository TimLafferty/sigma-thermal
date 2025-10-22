# Excel UDF Documentation

**Replace Excel VBA macros with Python-powered User Defined Functions**

---

## Overview

The Sigma Thermal Excel UDFs provide a modern replacement for legacy VBA macros. These Python-powered functions offer:

- ✅ **20+ calculation functions** for thermal engineering
- ✅ **Cross-platform** support (Windows and macOS)
- ✅ **Validated accuracy** (< 0.5% deviation from standards)
- ✅ **Easy migration** from existing VBA macros
- ✅ **Better maintainability** than VBA code

---

## Quick Start

### 1. Install xlwings

```bash
pip install xlwings
xlwings addin install
```

### 2. Copy Files to Workbook Folder

Copy these files to the same directory as your Excel workbook:
- `sigma_thermal_udf.py`
- `xlwings.conf`

(Files located in: `excel_udf/` directory in the repo)

### 3. Enable in Excel

1. Open your workbook
2. Go to xlwings ribbon tab
3. Click "Import Functions"
4. Enter: `sigma_thermal_udf`

### 4. Use Functions

```excel
=HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)  → 23389.33
=SATURATION_PRESSURE(212)                    → 14.648
=STEAM_ENTHALPY(212, 14.7, 1.0)             → 1156.21
```

---

## Documentation Files

### [Migration Guide](migration-guide.md)

**Complete step-by-step guide to replace VBA macros with Python UDFs**

Topics covered:
- Installation and setup
- VBA to Python function mapping
- Three migration strategies (manual, side-by-side, find & replace)
- Removing old VBA code
- Troubleshooting common issues
- Performance optimization
- Example migration scenarios

**Start here if you're migrating from VBA macros.**

### [Function Reference](function-reference.md)

**Complete documentation for all Excel UDF functions**

Topics covered:
- Installation instructions
- All 20+ functions with parameters and examples
- Usage examples and best practices
- Troubleshooting guide
- Performance tips
- Technical details

**Reference guide for all available functions.**

### [Quick Reference](quick-reference.md)

**One-page printable cheat sheet**

Topics covered:
- Function syntax and examples
- Common fuel compositions
- Units summary
- Quick troubleshooting

**Print this for quick reference at your desk.**

---

## Available Functions

### Heating Values
- `HHV_MASS_GAS()` - Higher heating value (mass basis)
- `LHV_MASS_GAS()` - Lower heating value (mass basis)
- `HHV_VOLUME_GAS()` - Higher heating value (volume basis)
- `LHV_VOLUME_GAS()` - Lower heating value (volume basis)

### Air Requirements
- `AIR_REQUIREMENT_MASS()` - Stoichiometric air (mass basis)
- `AIR_REQUIREMENT_VOLUME()` - Stoichiometric air (volume basis)

### Products of Combustion
- `POC_MASS()` - Products of combustion (mass basis)
- `POC_VOLUME()` - Products of combustion (volume basis)

### Flue Gas
- `FLUE_GAS_ENTHALPY()` - Flue gas sensible heat

### Steam Properties
- `SATURATION_PRESSURE()` - Saturation pressure from temperature
- `SATURATION_TEMPERATURE()` - Saturation temperature from pressure
- `STEAM_ENTHALPY()` - Steam enthalpy
- `STEAM_QUALITY()` - Vapor quality from enthalpy

### Water Properties
- `WATER_DENSITY()` - Water density
- `WATER_VISCOSITY()` - Dynamic viscosity
- `WATER_SPECIFIC_HEAT()` - Specific heat
- `WATER_THERMAL_CONDUCTIVITY()` - Thermal conductivity

### Helpers
- `HHV_NATURAL_GAS()` - Typical natural gas HHV
- `HHV_METHANE()` - Pure methane HHV

**Full details:** [Function Reference](function-reference.md)

---

## Example Usage

### Natural Gas Heating Value

```excel
A1: 85     (CH4 %)
B1: 10     (C2H6 %)
C1: 3      (C3H8 %)
D1: 1      (C4H10 %)
E1: 1      (CO2 %)

F1: =HHV_MASS_GAS(A1, B1, C1, D1, 0, 0, 0, E1, 0)
    Result: 23389.33 BTU/lb
```

### Steam Properties

```excel
A1: 212              (Temperature °F)
A2: 14.7             (Pressure psia)

B1: =SATURATION_PRESSURE(A1)        → 14.648 psia
B2: =SATURATION_TEMPERATURE(A2)     → 212.17 °F
B3: =STEAM_ENTHALPY(A1, A2, 1.0)    → 1156.21 BTU/lb (vapor)
B4: =STEAM_ENTHALPY(A1, A2, 0.0)    → 180.10 BTU/lb (liquid)
```

---

## Migration from VBA

### Quick Migration Steps

1. **Install xlwings** (one-time setup)
2. **Copy UDF files** to workbook folder
3. **Import functions** in Excel
4. **Replace function names** using Find & Replace

Example replacements:
```
Find: =HHVMass(          Replace: =HHV_MASS_GAS(
Find: =SaturationPressure(    Replace: =SATURATION_PRESSURE(
Find: =WaterDensity(     Replace: =WATER_DENSITY(
```

**Full guide:** [Migration Guide](migration-guide.md)

---

## Troubleshooting

### #NAME? Error

**Cause:** Excel cannot find the function

**Solution:**
1. Verify xlwings add-in is enabled
2. Re-import functions: xlwings tab → Import Functions
3. Restart Excel

### #VALUE! Error

**Cause:** Invalid input values

**Solution:**
1. Check inputs are numbers (not text)
2. Verify fuel composition sums to 100%
3. Check units (°F for temp, psia for pressure)

### Slow First Calculation

**Cause:** Python startup overhead

**Solution:**
- This is normal (2-5 seconds first time)
- Subsequent calculations are fast
- Use Manual calculation mode for large sheets

**Full troubleshooting:** [Function Reference](function-reference.md#troubleshooting)

---

## Requirements

- **Excel:** 2016 or later (Windows or macOS)
- **Python:** 3.11 or later
- **xlwings:** 0.30.0 or later
- **sigma-thermal:** 1.0.0 or later

---

## Files

Located in `excel_udf/` directory:

| File | Description |
|------|-------------|
| `sigma_thermal_udf.py` | Main UDF module with all functions |
| `xlwings.conf` | Configuration for xlwings add-in |
| `requirements.txt` | Python dependencies |

---

## Support

For help:
1. Check [Migration Guide](migration-guide.md) troubleshooting section
2. Review [Function Reference](function-reference.md) for detailed docs
3. Consult [Quick Reference](quick-reference.md) for syntax
4. Contact GTS Energy Inc.

---

## Next Steps

### New to Excel UDFs?
→ Start with [Migration Guide](migration-guide.md)

### Need function syntax?
→ Check [Quick Reference](quick-reference.md)

### Want detailed documentation?
→ Read [Function Reference](function-reference.md)

### Ready to migrate?
→ Follow [Migration Guide](migration-guide.md) step-by-step

---

**Back to:** [Main Documentation](../README.md)
