# Sigma Thermal Engineering Documentation

**Comprehensive documentation for the Sigma Thermal calculation library and deployment options**

---

## Quick Navigation

### For End Users

- **[Excel UDF Guide](excel-udf/)** - Replace VBA macros with Python functions in Excel
  - [Migration Guide](excel-udf/migration-guide.md) - Step-by-step VBA to Python migration
  - [Function Reference](excel-udf/function-reference.md) - Complete UDF documentation
  - [Quick Reference](excel-udf/quick-reference.md) - One-page cheat sheet

- **[Azure Deployment](azure-deployment/)** - Deploy web calculators to Azure
  - [Quick Start](azure-deployment/quick-start.md) - 5-minute deployment
  - [Deployment Guide](azure-deployment/deployment-guide.md) - Complete Azure setup

- **[Web Calculators](web-calculators/)** - HTML-based calculator interface
  - [HTML Calculators](web-calculators/html-calculators.md) - Web interface documentation

### For Developers

- **[Development Documentation](development/)** - Technical references
  - [Getting Started](development/getting_started.html) - Setup and development guide
  - [Validation Results](development/validation_results.html) - Test coverage and accuracy
  - [DOCUMENTATION_INDEX.md](DOCUMENTATION_INDEX.md) - Legacy documentation index
  - [EXECUTIVE_SUMMARY.md](EXECUTIVE_SUMMARY.md) - Project overview

---

## What is Sigma Thermal?

Sigma Thermal is a comprehensive Python library for industrial heater design and thermal calculations. It replaces legacy Excel VBA macros with modern, validated Python implementations.

### Key Features

- **43 calculation functions** covering combustion, steam, and fluid properties
- **Validated against industry standards** (ASME Steam Tables, VBA benchmarks)
- **Multiple deployment options** (Python library, Excel UDFs, web calculators, Azure)
- **Cross-platform** (Windows, macOS, Linux)
- **Well-documented** with comprehensive guides and examples

---

## Deployment Options

### Option 1: Excel UDFs (Recommended for Excel Users)

Replace your VBA macros with Python-powered User Defined Functions.

**Best for:** Existing Excel workflows, gradual migration from VBA

**Documentation:** [Excel UDF Guide](excel-udf/)

**Example:**
```excel
=HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)  → 23389.33 BTU/lb
=SATURATION_PRESSURE(212)                    → 14.648 psia
=STEAM_ENTHALPY(212, 14.7, 1.0)             → 1156.21 BTU/lb
```

### Option 2: Web Calculators (Azure Static Web Apps)

Professional HTML forms with Azure Functions API backend.

**Best for:** Team collaboration, mobile access, centralized deployment

**Documentation:** [Azure Deployment Guide](azure-deployment/)

**Live Demo:** Deploy in 5 minutes to `https://[your-app].azurestaticapps.net`

### Option 3: Python Library

Direct Python API for integration into custom applications.

**Best for:** Automated calculations, data pipelines, custom integrations

**Example:**
```python
from sigma_thermal.combustion import GasComposition, hhv_mass_gas
from sigma_thermal.fluids import saturation_pressure

fuel = GasComposition(methane_mass=100)
hhv = hhv_mass_gas(fuel)  # 23875.0 BTU/lb

p_sat = saturation_pressure(212)  # 14.648 psia
```

---

## Available Calculations

### Combustion Calculations

- **Heating Values:** HHV/LHV on mass and volume basis
- **Air Requirements:** Stoichiometric air (mass and volume)
- **Products of Combustion:** Flue gas composition and quantities
- **Flue Gas Enthalpy:** Heat loss calculations
- **Combustion Efficiency:** Performance metrics

### Steam Properties

- **Saturation Properties:** Pressure and temperature
- **Enthalpy:** Liquid, vapor, and two-phase
- **Quality:** Vapor fraction calculations

### Water Properties

- **Density:** Temperature-dependent
- **Viscosity:** Dynamic viscosity
- **Specific Heat:** Cp calculations
- **Thermal Conductivity:** Heat transfer properties

### Flash Steam

- **Flash Calculations:** Steam generation from pressure drop

---

## Getting Started

### For Excel Users

1. **Install xlwings:**
   ```bash
   pip install xlwings
   xlwings addin install
   ```

2. **Copy files to workbook folder:**
   - `sigma_thermal_udf.py`
   - `xlwings.conf`

3. **Import functions in Excel:**
   - xlwings tab → Import Functions → Enter: `sigma_thermal_udf`

4. **Start using functions:**
   ```excel
   =HHV_METHANE()  → 23875
   ```

**Full guide:** [Excel UDF Migration Guide](excel-udf/migration-guide.md)

### For Azure Deployment

1. **Deploy to Azure:**
   ```bash
   ./deploy-azure.sh
   ```

2. **Access your app:**
   ```
   https://sigma-thermal-calculators.azurestaticapps.net
   ```

**Full guide:** [Azure Quick Start](azure-deployment/quick-start.md)

### For Python Developers

1. **Install package:**
   ```bash
   pip install -e .
   ```

2. **Import and use:**
   ```python
   from sigma_thermal.combustion import hhv_mass_gas, GasComposition

   fuel = GasComposition(methane_mass=100)
   result = hhv_mass_gas(fuel)
   ```

**Full guide:** [Development Documentation](development/getting_started.html)

---

## Documentation Structure

```
docs/
├── README.md                          # This file - master documentation index
│
├── excel-udf/                         # Excel UDF documentation
│   ├── README.md                      # Excel UDF overview
│   ├── migration-guide.md             # VBA to Python migration
│   ├── function-reference.md          # Complete function documentation
│   └── quick-reference.md             # One-page cheat sheet
│
├── azure-deployment/                  # Azure deployment documentation
│   ├── README.md                      # Azure deployment overview
│   ├── quick-start.md                 # 5-minute deployment guide
│   └── deployment-guide.md            # Complete Azure setup
│
├── web-calculators/                   # Web interface documentation
│   ├── README.md                      # Web calculators overview
│   └── html-calculators.md            # HTML interface details
│
├── development/                       # Developer documentation
│   ├── getting_started.html           # Development setup
│   ├── validation_results.html        # Test results and validation
│   └── ...                            # Additional technical docs
│
└── archive/                           # Historical project documentation
    ├── PHASE1_COMPLETION_SUMMARY.md
    ├── PHASE2_PROGRESS.md
    └── ...                            # Legacy docs
```

---

## Key Documentation Files

### User Guides

| Document | Description | Audience |
|----------|-------------|----------|
| [Excel UDF Migration Guide](excel-udf/migration-guide.md) | Replace VBA macros with Python | Excel users migrating from VBA |
| [Excel UDF Function Reference](excel-udf/function-reference.md) | Complete UDF documentation | Excel users |
| [Excel UDF Quick Reference](excel-udf/quick-reference.md) | One-page function cheat sheet | Excel users |
| [Azure Quick Start](azure-deployment/quick-start.md) | 5-minute Azure deployment | DevOps, administrators |
| [Azure Deployment Guide](azure-deployment/deployment-guide.md) | Complete Azure setup | DevOps, administrators |
| [HTML Calculators](web-calculators/html-calculators.md) | Web interface documentation | Web developers |

### Technical References

| Document | Description | Audience |
|----------|-------------|----------|
| [Getting Started](development/getting_started.html) | Development environment setup | Python developers |
| [Validation Results](development/validation_results.html) | Test coverage and accuracy | QA, validators |
| [EXECUTIVE_SUMMARY.md](EXECUTIVE_SUMMARY.md) | Project overview | Management, stakeholders |
| [DOCUMENTATION_INDEX.md](DOCUMENTATION_INDEX.md) | Legacy documentation index | All users |

---

## Common Tasks

### Migrate Excel VBA Macros to Python

**Guide:** [Excel UDF Migration Guide](excel-udf/migration-guide.md)

**Steps:**
1. Install xlwings: `pip install xlwings && xlwings addin install`
2. Copy `sigma_thermal_udf.py` and `xlwings.conf` to workbook folder
3. Import functions in Excel
4. Replace VBA function names with Python UDF names
5. Validate results

**Example migration:**
```excel
Before: =HHVMass(A1, B1, C1, D1, 0, 0, 0, E1, 0)
After:  =HHV_MASS_GAS(A1, B1, C1, D1, 0, 0, 0, E1, 0)
```

### Deploy Web Calculators to Azure

**Guide:** [Azure Quick Start](azure-deployment/quick-start.md)

**Steps:**
1. Run: `./deploy-azure.sh`
2. Add GitHub secret (provided by script)
3. Push to main branch
4. Access at: `https://[your-app].azurestaticapps.net`

### Use Python Library Directly

**Guide:** [Development Getting Started](development/getting_started.html)

**Steps:**
1. Install: `pip install -e .`
2. Import: `from sigma_thermal.combustion import hhv_mass_gas, GasComposition`
3. Use: `result = hhv_mass_gas(fuel)`

---

## Validation and Testing

All calculations validated against:
- **ASME Steam Tables** (steam/water properties)
- **Excel VBA benchmarks** (combustion calculations)
- **Industry standards** (NIST, Perry's Handbook)

**Test coverage:**
- 412 unit tests
- < 0.5% deviation from reference data
- Validated across full operating range

**Results:** [Validation Results](development/validation_results.html)

---

## Support and Troubleshooting

### Excel UDF Issues

**Problem:** `#NAME?` error in Excel
**Solution:** Verify xlwings add-in enabled, re-import functions

**Problem:** `#VALUE!` error
**Solution:** Check input values, ensure fuel composition sums to 100%

**Problem:** Slow first calculation
**Solution:** Normal - Python startup takes 2-5 seconds, subsequent calculations fast

**Full troubleshooting:** [Excel UDF Function Reference](excel-udf/function-reference.md#troubleshooting)

### Azure Deployment Issues

**Problem:** 404 on API endpoint
**Solution:** Verify route in `function.json`, check deployment logs

**Problem:** CORS errors
**Solution:** Check `staticwebapp.config.json` headers

**Full troubleshooting:** [Azure Deployment Guide](azure-deployment/deployment-guide.md#troubleshooting)

### Python Library Issues

**Problem:** Import errors
**Solution:** Run `pip install -e .` in repo root

**Problem:** Validation failures
**Solution:** Check input units (°F for temp, psia for pressure)

**Full troubleshooting:** [Getting Started](development/getting_started.html)

---

## Version Information

**Version:** 1.0.0
**Date:** October 2025
**Author:** GTS Energy Inc.
**Python Version:** 3.11+
**Excel Version:** 2016+ (Windows and macOS)

---

## Project Files

### Source Code
- `src/sigma_thermal/` - Python library source
- `excel_udf/` - Excel UDF module and configuration
- `web/` - HTML calculator interface
- `api/` - Azure Functions API backend

### Documentation
- `docs/` - This documentation directory
- `README.md` - Project root README

### Deployment
- `deploy-azure.sh` - Azure deployment script
- `.github/workflows/` - CI/CD pipelines

### Testing
- `tests/` - Unit tests
- `validation/` - Validation test suites

---

## Next Steps

### New Users

1. Choose deployment option (Excel UDFs, Azure, or Python library)
2. Follow relevant getting started guide
3. Try example calculations
4. Refer to function reference as needed

### Excel Users Migrating from VBA

1. Start with [Excel UDF Migration Guide](excel-udf/migration-guide.md)
2. Install xlwings and copy UDF files
3. Test functions side-by-side with VBA
4. Migrate formulas using find & replace
5. Validate results and remove VBA

### Developers

1. Review [Development Getting Started](development/getting_started.html)
2. Install package: `pip install -e .`
3. Run tests: `pytest`
4. Review [Validation Results](development/validation_results.html)
5. Start integrating into your application

### DevOps/Administrators

1. Review [Azure Deployment Guide](azure-deployment/deployment-guide.md)
2. Run deployment script: `./deploy-azure.sh`
3. Configure custom domain (optional)
4. Set up monitoring with Application Insights
5. Review cost estimates

---

## Contributing

For questions, issues, or contributions:
1. Check relevant documentation section
2. Review troubleshooting guides
3. Consult validation results
4. Contact GTS Energy Inc.

---

## License

Copyright © 2025 GTS Energy Inc.
All rights reserved.

---

**Last Updated:** October 22, 2025
**Documentation Version:** 1.0.0
