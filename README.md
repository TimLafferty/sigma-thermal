# Sigma Thermal

**Industrial heater design and thermal engineering calculation library**

---

## Overview

Sigma Thermal is a comprehensive Python library for thermal engineering calculations. It replaces legacy Excel VBA macros with modern, validated Python implementations.

### Key Features

- ✅ **43 calculation functions** for combustion, steam, and fluid properties
- ✅ **Validated accuracy** (< 0.5% deviation from ASME standards)
- ✅ **Multiple deployment options** (Python library, Excel UDFs, web calculators, Azure)
- ✅ **Cross-platform** support (Windows, macOS, Linux)
- ✅ **Comprehensive documentation** with guides and examples

---

## Quick Navigation

### 📚 [Complete Documentation](docs/)

All documentation consolidated in the `docs/` directory:

- **[Excel UDF Guide](docs/excel-udf/)** - Replace VBA macros with Python functions
- **[Azure Deployment](docs/azure-deployment/)** - Deploy web calculators to Azure
- **[Web Calculators](docs/web-calculators/)** - HTML-based calculator interface
- **[Development Docs](docs/development/)** - Technical references and validation

### 🚀 Quick Start Guides

- **[Excel Users](docs/excel-udf/migration-guide.md)** - Migrate from VBA to Python UDFs
- **[Azure Deployment](docs/azure-deployment/quick-start.md)** - 5-minute Azure setup
- **[Python Developers](docs/development/getting_started.html)** - Setup and API reference

---

## Installation

### For Excel UDFs

```bash
pip install xlwings
xlwings addin install
```

Then copy `excel_udf/sigma_thermal_udf.py` and `excel_udf/xlwings.conf` to your workbook folder.

**Full guide:** [Excel UDF Migration](docs/excel-udf/migration-guide.md)

### For Azure Deployment

```bash
./deploy-azure.sh
```

**Full guide:** [Azure Quick Start](docs/azure-deployment/quick-start.md)

### For Python Development

```bash
git clone https://github.com/gts-energy/sigma-thermal.git
cd sigma-thermal
pip install -e ".[dev]"
```

**Full guide:** [Development Setup](docs/development/getting_started.html)

---

## Usage Examples

### Excel UDFs

```excel
=HHV_MASS_GAS(85, 10, 3, 1, 0, 0, 0, 1, 0)  → 23389.33 BTU/lb
=SATURATION_PRESSURE(212)                    → 14.648 psia
=STEAM_ENTHALPY(212, 14.7, 1.0)             → 1156.21 BTU/lb
=WATER_DENSITY(60)                           → 62.33 lb/ft³
```

**Full reference:** [Excel UDF Function Reference](docs/excel-udf/function-reference.md)

### Python API

```python
from sigma_thermal.combustion import GasComposition, hhv_mass_gas
from sigma_thermal.fluids import saturation_pressure, steam_enthalpy

# Heating value calculation
fuel = GasComposition(methane_mass=100)
hhv = hhv_mass_gas(fuel)  # 23875.0 BTU/lb

# Steam properties
p_sat = saturation_pressure(212)  # 14.648 psia
h_steam = steam_enthalpy(212, 14.7, 1.0)  # 1156.21 BTU/lb
```

### Web Calculators

Access deployed calculators at:
```
https://[your-app].azurestaticapps.net
```

**Full guide:** [Web Calculators](docs/web-calculators/)

---

## Available Calculations

### Combustion

- Heating Values (HHV/LHV on mass and volume basis)
- Air Requirements (stoichiometric air, mass and volume)
- Products of Combustion (flue gas composition)
- Flue Gas Enthalpy (heat loss calculations)
- Combustion Efficiency

### Steam Properties

- Saturation Pressure and Temperature
- Steam Enthalpy (liquid, vapor, two-phase)
- Steam Quality (vapor fraction)

### Water Properties

- Density (temperature-dependent)
- Viscosity (dynamic viscosity)
- Specific Heat (Cp calculations)
- Thermal Conductivity

### Flash Steam

- Flash calculations (steam generation from pressure drop)

**Full details:** [Documentation](docs/README.md)

---

## Deployment Options

### Option 1: Excel UDFs

Replace VBA macros with Python-powered User Defined Functions.

**Best for:** Existing Excel workflows, gradual VBA migration

**Guide:** [Excel UDF Migration](docs/excel-udf/migration-guide.md)

### Option 2: Web Calculators (Azure)

Professional HTML forms with Azure Functions backend.

**Best for:** Team collaboration, mobile access, centralized deployment

**Guide:** [Azure Deployment](docs/azure-deployment/)

### Option 3: Python Library

Direct Python API for custom applications.

**Best for:** Automated calculations, data pipelines, integrations

**Guide:** [Development Documentation](docs/development/getting_started.html)

---

## Project Structure

```
sigma-thermal/
├── src/sigma_thermal/        # Python library source
│   ├── combustion/            # Combustion calculations
│   ├── fluids/                # Fluid properties
│   ├── heat_transfer/         # Heat transfer
│   └── engineering/           # Engineering utilities
├── excel_udf/                 # Excel UDF module
│   ├── sigma_thermal_udf.py   # UDF functions
│   └── xlwings.conf           # Configuration
├── web/                       # Web calculator interface
│   ├── index.html             # Landing page
│   ├── resource.html          # Technical reference
│   └── calculators/           # Calculator pages
├── api/                       # Azure Functions backend
│   ├── heating_value/         # Heating value API
│   └── host.json              # Functions config
├── tests/                     # Unit tests
├── docs/                      # Documentation
│   ├── excel-udf/             # Excel UDF docs
│   ├── azure-deployment/      # Azure deployment docs
│   ├── web-calculators/       # Web interface docs
│   └── development/           # Technical docs
└── deploy-azure.sh            # Azure deployment script
```

---

## Testing

Run the test suite:

```bash
pytest
```

View validation results:

```bash
open docs/development/validation_results.html
```

**Test coverage:** 412 unit tests, < 0.5% deviation from standards

---

## Documentation

### Complete Documentation

All documentation is in the `docs/` directory:

**[📚 View Complete Documentation](docs/README.md)**

### Quick Links

| Guide | Description |
|-------|-------------|
| [Excel UDF Migration](docs/excel-udf/migration-guide.md) | Replace VBA macros with Python |
| [Excel Function Reference](docs/excel-udf/function-reference.md) | Complete UDF documentation |
| [Azure Quick Start](docs/azure-deployment/quick-start.md) | 5-minute Azure deployment |
| [Azure Deployment Guide](docs/azure-deployment/deployment-guide.md) | Complete Azure setup |
| [Web Calculators](docs/web-calculators/) | HTML interface documentation |
| [Development Setup](docs/development/getting_started.html) | Python development guide |
| [Validation Results](docs/development/validation_results.html) | Test coverage and accuracy |

---

## Support

### Resources

- **Documentation:** [docs/](docs/README.md)
- **Excel UDF Help:** [Excel UDF Guide](docs/excel-udf/)
- **Azure Deployment Help:** [Azure Deployment](docs/azure-deployment/)
- **Technical Reference:** [Development Docs](docs/development/)

### Getting Help

1. Check relevant documentation in `docs/`
2. Review troubleshooting sections
3. Consult validation results
4. Contact GTS Energy Inc.

---

## Requirements

- **Python:** 3.11 or later
- **Excel (for UDFs):** 2016+ (Windows or macOS)
- **xlwings (for UDFs):** 0.30.0 or later
- **Azure CLI (for deployment):** Latest version

---

## License

Copyright © 2025 GTS Energy Inc.
All rights reserved.

---

## Version

**Version:** 1.0.0
**Date:** October 2025
**Author:** GTS Energy Inc.

---

## Quick Start

### For Excel Users

```bash
pip install xlwings && xlwings addin install
```

**Then:** [Follow Excel UDF Migration Guide](docs/excel-udf/migration-guide.md)

### For Azure Deployment

```bash
./deploy-azure.sh
```

**Then:** [Follow Azure Quick Start](docs/azure-deployment/quick-start.md)

### For Python Developers

```bash
pip install -e .
```

**Then:** [Follow Development Guide](docs/development/getting_started.html)

---

**For complete documentation, visit:** [docs/](docs/README.md)
