# Sigma Thermal Engineering Calculators

Web-based calculator interface for Sigma Thermal engineering functions.

## Features

- **9 Interactive Calculators:**
  - 🔥 Heating Value Calculator (COMPLETE)
  - 💨 Air Requirement Calculator (Placeholder)
  - 📊 Products of Combustion (Placeholder)
  - 🌡️ Flue Gas Enthalpy (Placeholder)
  - ⚡ Combustion Efficiency (Placeholder)
  - 💧 Steam Properties Calculator (COMPLETE)
  - 💦 Water Properties (Placeholder)
  - ⚙️ Flash Steam Calculator (Placeholder)
  - 🔍 Excel Comparison Tool (Placeholder)

- **User-Friendly Interface:**
  - Clean, professional design
  - Example scenario presets
  - Real-time validation
  - Export results (JSON)
  - Charts and visualizations

- **Validation Features:**
  - Compare Python vs Excel VBA
  - ASME Steam Table validation
  - Deviation analysis

## Installation

### Requirements

```bash
pip install streamlit plotly pandas
```

Or install from the sigma-thermal package:

```bash
pip install -e .
```

## Running the App

### Local Development

From the `src/sigma_thermal/calculators` directory:

```bash
streamlit run app.py
```

Or from anywhere:

```bash
python -m streamlit run src/sigma_thermal/calculators/app.py
```

The app will open in your browser at `http://localhost:8501`

### Docker (Optional)

```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY . .

RUN pip install -e .

EXPOSE 8501

CMD ["streamlit", "run", "src/sigma_thermal/calculators/app.py"]
```

Build and run:

```bash
docker build -t sigma-thermal-calculators .
docker run -p 8501:8501 sigma-thermal-calculators
```

## Project Structure

```
calculators/
├── app.py                      # Main Streamlit application
├── pages/                      # Calculator pages
│   ├── heating_value.py        # ✅ Heating value calculator
│   ├── steam_properties.py     # ✅ Steam properties calculator
│   ├── air_requirement.py      # 🚧 Placeholder
│   ├── products_combustion.py  # 🚧 Placeholder
│   ├── flue_gas_enthalpy.py    # 🚧 Placeholder
│   ├── combustion_efficiency.py # 🚧 Placeholder
│   ├── water_properties.py     # 🚧 Placeholder
│   ├── flash_steam.py          # 🚧 Placeholder
│   └── excel_comparison.py     # 🚧 Placeholder
├── utils/                      # Utility functions
│   └── ui_components.py        # Reusable UI components
└── data/                       # Data and presets
    └── presets.py              # Fuel compositions, conditions
```

## Usage Examples

### Heating Value Calculator

1. Navigate to "🔥 Heating Value Calculator"
2. Select a fuel preset (e.g., "Natural Gas (Typical)")
3. Adjust composition if needed
4. Click "Calculate Heating Values"
5. View results on mass and volume basis
6. Export results as JSON

### Steam Properties Calculator

1. Navigate to "💧 Steam Properties"
2. Choose calculation mode:
   - Temperature & Pressure (Known)
   - Enthalpy & Pressure (Known)
   - Saturation Properties Only
3. Enter parameters
4. Click "Calculate Steam Properties"
5. View phase, enthalpy, quality
6. See T-s diagram visualization

## Implemented Functions

### Heating Value Calculator
- `hhv_mass_gas()` - Higher heating value (mass basis)
- `lhv_mass_gas()` - Lower heating value (mass basis)
- `hhv_volume_gas()` - Higher heating value (volume basis)
- `lhv_volume_gas()` - Lower heating value (volume basis)

### Steam Properties Calculator
- `saturation_pressure()` - Psat from temperature
- `saturation_temperature()` - Tsat from pressure
- `steam_enthalpy()` - Enthalpy (T, P, quality)
- `steam_quality()` - Quality from enthalpy

## Example Presets

### Fuel Compositions
- Pure Methane
- Natural Gas (Typical)
- Natural Gas (High BTU)
- Natural Gas (Lean)
- Landfill Gas
- Digester Gas
- Refinery Gas
- Coke Oven Gas
- Blast Furnace Gas

### Operating Conditions
- Standard (77°F, 10% XSA)
- Boiler (Low/Moderate excess air)
- Furnace
- Heater
- Incinerator (High excess air)
- Cold Weather
- Hot & Humid

### Steam Pressures
- Vacuum (Evaporator): 2 psia
- Low Pressure (HVAC): 15 psia
- Atmospheric: 14.7 psia
- Low Steam: 64.7 psia (50 psig)
- Medium Steam: 114.7 psia (100 psig)
- High Steam: 164.7 psia (150 psig)
- And more...

## Settings

Access settings in the sidebar:
- **Unit System:** US Customary (SI coming soon)
- **Decimal Places:** 0-6 (default: 2)

## Validation

The calculators include built-in validation against:
- Excel VBA functions (where implemented)
- ASME Steam Tables
- NIST reference data

Deviation analysis shows:
- ✅ PASS: <1% deviation
- 🟡 WARNING: 1-2% deviation
- ❌ FAIL: >2% deviation

## Development Roadmap

### Week 3 (Current)
- [x] Main app structure
- [x] Heating Value Calculator
- [x] Steam Properties Calculator
- [ ] Products of Combustion Calculator
- [ ] Water Properties Calculator
- [ ] Excel comparison tool

### Week 4
- [ ] Combustion Efficiency Calculator
- [ ] Air Requirement Calculator
- [ ] Flue Gas Enthalpy Calculator
- [ ] Flash Steam Calculator

### Future
- [ ] SI unit support
- [ ] PDF report generation
- [ ] User accounts & saved calculations
- [ ] Mobile app
- [ ] API access

## Contributing

To add a new calculator:

1. Create page file in `pages/`
2. Import sigma_thermal functions
3. Use `ui_components` for consistent styling
4. Add navigation link in `app.py`
5. Test thoroughly
6. Update this README

## Support

For issues or questions:
- GitHub Issues: https://github.com/gts-energy/sigma-thermal/issues
- Documentation: See `/docs/CALCULATOR_UI_REQUIREMENTS.md`

## License

© 2025 GTS Energy Inc.

---

*Version: 1.0*
*Last Updated: October 22, 2025*
*Status: 2 calculators complete, 7 placeholders*
