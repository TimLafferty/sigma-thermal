# Sigma Thermal Excel UDF Guide

Replace Excel VBA macros with Python-powered User Defined Functions (UDFs) for thermal engineering calculations.

**Author:** GTS Energy Inc.
**Date:** October 2025

---

## What Are Excel UDFs?

Excel User Defined Functions (UDFs) are custom functions that you can use directly in Excel formulas, just like built-in functions such as `SUM()` or `AVERAGE()`. The Sigma Thermal UDFs provide professional thermal engineering calculations powered by Python.

### Why Use These UDFs?

- Replace outdated VBA macros with modern Python
- Access validated thermal engineering calculations
- Use familiar Excel formula syntax
- Automatic recalculation when inputs change
- Cross-platform compatibility (Windows and macOS)

---

## Prerequisites

### Required Software

1. **Microsoft Excel** (Windows or macOS)
   - Excel 2016 or later recommended
   - Excel 365 fully supported

2. **Python 3.11** or later
   - Download from: https://www.python.org/downloads/
   - Verify installation: Open terminal/command prompt and run `python --version`

3. **xlwings** (installed in setup steps below)

### System Requirements

- **Windows:** Windows 10 or later
- **macOS:** macOS 10.15 (Catalina) or later

---

## Installation

### Step 1: Install Python Dependencies

Open terminal (macOS) or command prompt (Windows) and run:

```bash
# Install xlwings
pip install xlwings

# Install sigma-thermal package
cd /path/to/sigma-thermal
pip install -e .
```

### Step 2: Install xlwings Excel Add-in

```bash
# Install the xlwings add-in to Excel
xlwings addin install

# Verify installation
xlwings addin status
```

Expected output:
```
xlwings add-in is installed
```

### Step 3: Copy UDF Module

1. Locate your Excel workbook directory
2. Copy `sigma_thermal_udf.py` and `xlwings.conf` to the same folder as your Excel workbook

**Directory structure:**
```
My Documents/
└── Thermal Calculations/
    ├── calculations.xlsx
    ├── sigma_thermal_udf.py
    └── xlwings.conf
```

### Step 4: Configure Excel Workbook

1. Open your Excel workbook
2. Go to the xlwings tab in the ribbon
3. Click "Import Functions"
4. Select "UDF Modules" and enter: `sigma_thermal_udf`
5. Click OK

---

## Available Functions

### Heating Value Functions

Calculate higher and lower heating values for gaseous fuels.

#### HHV_MASS_GAS
**Description:** Higher heating value on mass basis
**Returns:** BTU/lb
**Parameters:**
- `ch4` - Methane mass % (required)
- `c2h6` - Ethane mass % (default: 0)
- `c3h8` - Propane mass % (default: 0)
- `c4h10` - Butane mass % (default: 0)
- `h2` - Hydrogen mass % (default: 0)
- `co` - Carbon monoxide mass % (default: 0)
- `h2s` - Hydrogen sulfide mass % (default: 0)
- `co2` - Carbon dioxide mass % (default: 0)
- `n2` - Nitrogen mass % (default: 0)

**Example:**
```excel
=HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)
Returns: 22487.23
```

#### LHV_MASS_GAS
**Description:** Lower heating value on mass basis
**Returns:** BTU/lb

**Example:**
```excel
=LHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)
Returns: 20256.45
```

#### HHV_VOLUME_GAS
**Description:** Higher heating value on volume basis
**Returns:** BTU/scf

**Example:**
```excel
=HHV_VOLUME_GAS(100, 0, 0, 0, 0, 0, 0, 0, 0)
Returns: 1012.00
```

#### LHV_VOLUME_GAS
**Description:** Lower heating value on volume basis
**Returns:** BTU/scf

**Example:**
```excel
=LHV_VOLUME_GAS(100, 0, 0, 0, 0, 0, 0, 0, 0)
Returns: 910.00
```

### Air Requirement Functions

Calculate stoichiometric air requirements for combustion.

#### AIR_REQUIREMENT_MASS
**Description:** Stoichiometric air requirement on mass basis
**Returns:** lb air/lb fuel

**Example:**
```excel
=AIR_REQUIREMENT_MASS(100, 0, 0, 0, 0, 0, 0)
Returns: 17.24
```

#### AIR_REQUIREMENT_VOLUME
**Description:** Stoichiometric air requirement on volume basis
**Returns:** scf air/scf fuel

**Example:**
```excel
=AIR_REQUIREMENT_VOLUME(100, 0, 0, 0, 0, 0, 0)
Returns: 9.52
```

### Products of Combustion Functions

Calculate flue gas products from combustion.

#### POC_MASS
**Description:** Products of combustion on mass basis
**Returns:** lb POC/lb fuel
**Parameters:**
- Fuel composition (ch4 through h2s)
- `excess_air` - Excess air % (default: 15)

**Example:**
```excel
=POC_MASS(100, 0, 0, 0, 0, 0, 0, 15)
Returns: 19.83
```

#### POC_VOLUME
**Description:** Products of combustion on volume basis
**Returns:** scf POC/scf fuel

**Example:**
```excel
=POC_VOLUME(100, 0, 0, 0, 0, 0, 0, 15)
Returns: 10.95
```

### Flue Gas Enthalpy Function

Calculate enthalpy of flue gas leaving the combustion system.

#### FLUE_GAS_ENTHALPY
**Description:** Flue gas sensible heat
**Returns:** BTU/lb fuel
**Parameters:**
- Fuel composition (ch4 through h2s)
- `flue_gas_temp` - Flue gas temperature °F (default: 350)
- `excess_air` - Excess air % (default: 15)
- `fuel_temp` - Fuel temperature °F (default: 60)
- `air_temp` - Combustion air temperature °F (default: 60)

**Example:**
```excel
=FLUE_GAS_ENTHALPY(100, 0, 0, 0, 0, 0, 0, 350, 15, 60, 60)
Returns: 1847.23
```

### Steam Properties Functions

Calculate thermodynamic properties of water and steam.

#### SATURATION_PRESSURE
**Description:** Saturation pressure from temperature
**Returns:** psia

**Example:**
```excel
=SATURATION_PRESSURE(212)
Returns: 14.696
```

#### SATURATION_TEMPERATURE
**Description:** Saturation temperature from pressure
**Returns:** °F

**Example:**
```excel
=SATURATION_TEMPERATURE(14.696)
Returns: 212.00
```

#### STEAM_ENTHALPY
**Description:** Steam enthalpy
**Returns:** BTU/lb
**Parameters:**
- `temperature` - Temperature °F
- `pressure` - Pressure psia
- `quality` - Quality 0-1 (default: 1.0, where 0=liquid, 1=vapor)

**Examples:**
```excel
=STEAM_ENTHALPY(212, 14.696, 1.0)
Returns: 1150.40  (saturated vapor)

=STEAM_ENTHALPY(212, 14.696, 0.0)
Returns: 180.10  (saturated liquid)

=STEAM_ENTHALPY(212, 14.696, 0.5)
Returns: 665.25  (50% quality)
```

#### STEAM_QUALITY
**Description:** Steam quality from enthalpy and pressure
**Returns:** Quality (0-1)

**Example:**
```excel
=STEAM_QUALITY(650, 14.696)
Returns: 0.484  (48.4% vapor)
```

### Water Properties Functions

Calculate physical properties of liquid water.

#### WATER_DENSITY
**Description:** Water density at given temperature
**Returns:** lb/ft³

**Example:**
```excel
=WATER_DENSITY(60)
Returns: 62.37
```

#### WATER_VISCOSITY
**Description:** Water dynamic viscosity
**Returns:** lb/ft·s

**Example:**
```excel
=WATER_VISCOSITY(60)
Returns: 0.000752
```

#### WATER_SPECIFIC_HEAT
**Description:** Water specific heat
**Returns:** BTU/lb·°F

**Example:**
```excel
=WATER_SPECIFIC_HEAT(60)
Returns: 0.999
```

#### WATER_THERMAL_CONDUCTIVITY
**Description:** Water thermal conductivity
**Returns:** BTU/hr·ft·°F

**Example:**
```excel
=WATER_THERMAL_CONDUCTIVITY(60)
Returns: 0.340
```

### Helper Functions

Quick reference functions for common fuel types.

#### HHV_NATURAL_GAS
**Description:** HHV for typical natural gas
**Composition:** 85% CH4, 10% C2H6, 3% C3H8, 1% C4H10, 1% CO2

**Example:**
```excel
=HHV_NATURAL_GAS()
Returns: 22487
```

#### HHV_METHANE
**Description:** HHV for pure methane

**Example:**
```excel
=HHV_METHANE()
Returns: 23875
```

---

## Usage Examples

### Example 1: Natural Gas Analysis

Calculate heating values and air requirements for a natural gas composition:

| Component | Cell | Value |
|-----------|------|-------|
| CH4 | A2 | 85 |
| C2H6 | A3 | 10 |
| C3H8 | A4 | 3 |
| C4H10 | A5 | 1 |
| CO2 | A6 | 1 |

**Formulas:**
```excel
B2: =HHV_MASS_GAS(A2, A3, A4, A5, 0, 0, 0, A6, 0)
B3: =LHV_MASS_GAS(A2, A3, A4, A5, 0, 0, 0, A6, 0)
B4: =AIR_REQUIREMENT_MASS(A2, A3, A4, A5, 0, 0, 0)
B5: =POC_MASS(A2, A3, A4, A5, 0, 0, 0, 15)
```

### Example 2: Combustion Efficiency

Calculate flue gas loss for efficiency analysis:

```excel
A1: Fuel: 100% Methane
A2: Flue Gas Temperature: 350°F
A3: Excess Air: 15%

B2: =FLUE_GAS_ENTHALPY(100, 0, 0, 0, 0, 0, 0, A2, A3, 60, 60)
B3: =HHV_MASS_GAS(100, 0, 0, 0, 0, 0, 0, 0, 0)
B4: =(B2 / B3) * 100    ' Flue gas loss %
```

### Example 3: Steam Properties Table

Create a steam table for saturated conditions:

| Pressure (psia) | Sat Temp (°F) | hf (BTU/lb) | hg (BTU/lb) |
|-----------------|---------------|-------------|-------------|
| =A2 | =SATURATION_TEMPERATURE(A2) | =STEAM_ENTHALPY(B2, A2, 0) | =STEAM_ENTHALPY(B2, A2, 1) |

Fill down for pressures: 14.696, 50, 100, 150, 200 psia

### Example 4: Water Property Analysis

Calculate Reynolds number for pipe flow:

```excel
A1: Temperature: 60°F
A2: Velocity: 5 ft/s
A3: Pipe Diameter: 0.5 ft

B1: =WATER_DENSITY(A1)
B2: =WATER_VISCOSITY(A1)
B3: =(B1 * A2 * A3) / B2    ' Reynolds number
```

---

## Tips and Best Practices

### Performance Optimization

1. **Minimize recalculations:** Set Excel to manual calculation mode (Formulas → Calculation Options → Manual) for large workbooks
2. **Use helper cells:** Break complex formulas into multiple cells for easier debugging
3. **Avoid volatile functions:** Don't nest UDFs inside volatile functions like `NOW()` or `RAND()`

### Error Handling

If a UDF returns an error:

1. **Check input values:** Ensure all percentages sum to 100% for composition functions
2. **Verify units:** All temperatures in °F, pressures in psia
3. **Check ranges:** Steam quality must be between 0 and 1

### Calculation Accuracy

- All functions validated against ASME Steam Tables and industry standards
- Typical accuracy: < 0.5% deviation from reference data
- See validation results in `web/resource.html`

---

## Troubleshooting

### Issue: "#NAME?" error in Excel

**Cause:** Excel cannot find the UDF function

**Solution:**
1. Verify xlwings add-in is installed: `xlwings addin status`
2. Check that `sigma_thermal_udf.py` is in the workbook directory
3. Verify UDF module is imported in xlwings ribbon
4. Restart Excel

### Issue: "#VALUE!" error in Excel

**Cause:** Invalid input values or Python error

**Solution:**
1. Check that input values are numeric
2. Verify percentages sum to 100% for composition functions
3. Check Python console for error messages
4. Enable `SHOW_LOG = True` in `xlwings.conf`

### Issue: Functions calculate slowly

**Cause:** Python startup overhead or complex calculations

**Solution:**
1. First calculation may be slow (Python initialization)
2. Subsequent calculations should be faster
3. Set Excel to manual calculation mode for large sheets
4. Consider using array formulas for repetitive calculations

### Issue: "Module not found" error

**Cause:** sigma_thermal package not installed

**Solution:**
```bash
cd /path/to/sigma-thermal
pip install -e .
```

### Issue: xlwings add-in not visible in Excel

**Windows:**
1. File → Options → Add-ins
2. Manage Excel Add-ins → Go
3. Check "xlwings"

**macOS:**
1. Tools → Excel Add-ins
2. Check "xlwings"

### Issue: Different results from VBA macros

**Cause:** Improved calculation methods or bug fixes

**Solution:**
- Python UDFs use updated correlations and methods
- Verify against published data in `web/resource.html`
- Python results are typically more accurate than old VBA code

---

## Uninstalling

To remove the xlwings add-in:

```bash
xlwings addin remove
```

To uninstall Python packages:

```bash
pip uninstall xlwings sigma-thermal
```

---

## Technical Details

### Function Naming Convention

- **ALL_CAPS:** Excel UDF functions use uppercase naming (e.g., `HHV_MASS_GAS`)
- **snake_case:** Python backend uses lowercase (e.g., `hhv_mass_gas`)

### Default Parameter Values

All UDFs use sensible defaults:
- Fuel components: 0% (except first required component)
- Excess air: 15%
- Temperatures: 60°F (ambient)
- Steam quality: 1.0 (saturated vapor)

### Calculation Methods

- **Heating values:** Mass-weighted component HHVs
- **Air requirements:** Stoichiometric oxygen demand
- **Steam properties:** IAPWS-IF97 industrial formulation
- **Water properties:** Polynomial correlations (60-212°F)

---

## Support

### Documentation

- **Web calculators:** See `web/resource.html` for detailed formulas
- **API documentation:** See `docs/` directory
- **Package documentation:** Run `python -c "import sigma_thermal; help(sigma_thermal)"`

### Testing

Test the UDF module directly from Python:

```bash
cd excel_udf
python sigma_thermal_udf.py
```

Expected output:
```
Testing Sigma Thermal UDFs...
HHV Natural Gas: 22487 BTU/lb
HHV Methane: 23875 BTU/lb
Air Requirement (mass): 17.24 lb/lb
POC (mass): 19.83 lb/lb
Flue Gas Enthalpy: 1847.23 BTU/lb
Saturation Pressure @ 212°F: 14.696 psia
Steam Enthalpy @ 212°F, 14.7 psia, x=1.0: 1150.4 BTU/lb
All tests passed!
```

### Getting Help

For issues or questions:
1. Check this guide's troubleshooting section
2. Review validation data in `web/resource.html`
3. Contact GTS Energy Inc.

---

## Version History

**v1.0.0** - October 2025
- Initial release
- 20+ Excel UDF functions
- Combustion calculations (heating values, air requirements, POC)
- Steam properties (saturation, enthalpy, quality)
- Water properties (density, viscosity, thermal properties)
- Validated against ASME standards

---

## Summary

The Sigma Thermal Excel UDFs provide:

- **20+ functions** for thermal engineering calculations
- **Professional accuracy** validated against industry standards
- **Easy integration** with existing Excel workflows
- **Cross-platform** support (Windows and macOS)
- **Python-powered** modern calculation engine

Replace your VBA macros with these validated, maintainable Python functions.

**Status:** Ready to use
**Platform:** Windows and macOS
**Excel Version:** 2016 or later
**Python Version:** 3.11+
