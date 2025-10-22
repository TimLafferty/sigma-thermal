# Sigma Thermal Excel UDFs

Python-powered User Defined Functions for Microsoft Excel - Replace VBA macros with validated thermal engineering calculations.

---

## Quick Start

### 1. Install Dependencies

```bash
pip install xlwings
pip install -e ..
xlwings addin install
```

### 2. Copy Files to Workbook Folder

Copy these files to the same folder as your Excel workbook:
- `sigma_thermal_udf.py`
- `xlwings.conf`

### 3. Use in Excel

Open Excel and use functions directly in formulas:

```excel
=HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)
=SATURATION_PRESSURE(212)
=STEAM_ENTHALPY(212, 14.7, 1.0)
```

---

## Documentation

- **[MIGRATION_GUIDE.md](MIGRATION_GUIDE.md)** - Step-by-step guide to replace VBA macros with Python UDFs
- **[EXCEL_UDF_GUIDE.md](EXCEL_UDF_GUIDE.md)** - Complete installation and usage guide
- **[QUICK_REFERENCE.md](QUICK_REFERENCE.md)** - One-page function reference (print-friendly)

---

## Available Functions (20+)

### Combustion Calculations
- `HHV_MASS_GAS()`, `LHV_MASS_GAS()` - Heating values (mass basis)
- `HHV_VOLUME_GAS()`, `LHV_VOLUME_GAS()` - Heating values (volume basis)
- `AIR_REQUIREMENT_MASS()`, `AIR_REQUIREMENT_VOLUME()` - Stoichiometric air
- `POC_MASS()`, `POC_VOLUME()` - Products of combustion
- `FLUE_GAS_ENTHALPY()` - Flue gas sensible heat

### Steam Properties
- `SATURATION_PRESSURE()`, `SATURATION_TEMPERATURE()` - Saturation properties
- `STEAM_ENTHALPY()` - Enthalpy of water/steam
- `STEAM_QUALITY()` - Vapor quality from enthalpy

### Water Properties
- `WATER_DENSITY()` - Density
- `WATER_VISCOSITY()` - Dynamic viscosity
- `WATER_SPECIFIC_HEAT()` - Specific heat
- `WATER_THERMAL_CONDUCTIVITY()` - Thermal conductivity

### Helper Functions
- `HHV_NATURAL_GAS()` - Typical natural gas HHV
- `HHV_METHANE()` - Pure methane HHV

---

## Testing

Test the module from command line:

```bash
python sigma_thermal_udf.py
```

Expected output:
```
Testing Sigma Thermal UDFs...
HHV Natural Gas: 22487 BTU/lb
HHV Methane: 23875 BTU/lb
...
All tests passed!
```

---

## Requirements

- **Microsoft Excel:** 2016 or later (Windows or macOS)
- **Python:** 3.11 or later
- **xlwings:** 0.30.0 or later
- **sigma-thermal:** 1.0.0 or later

---

## Files in this Directory

| File | Description |
|------|-------------|
| `sigma_thermal_udf.py` | Main UDF module with all functions |
| `xlwings.conf` | Configuration file for xlwings add-in |
| `requirements.txt` | Python package dependencies |
| `EXCEL_UDF_GUIDE.md` | Complete installation and usage guide |
| `QUICK_REFERENCE.md` | One-page function reference |
| `README.md` | This file |

---

## Support

For issues or questions:
1. See troubleshooting section in `EXCEL_UDF_GUIDE.md`
2. Check validation data at `../web/resource.html`
3. Contact GTS Energy Inc.

---

**Status:** Ready to use
**Version:** 1.0.0
**Date:** October 2025
**Author:** GTS Energy Inc.
