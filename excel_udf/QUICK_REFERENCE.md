# Sigma Thermal Excel UDF Quick Reference

**GTS Energy Inc. | October 2025**

---

## Installation

```bash
pip install xlwings
pip install -e /path/to/sigma-thermal
xlwings addin install
```

Copy `sigma_thermal_udf.py` and `xlwings.conf` to your workbook folder.

---

## Heating Value Functions

| Function | Description | Units | Example |
|----------|-------------|-------|---------|
| `HHV_MASS_GAS(ch4, c2h6, c3h8, c4h10, h2, co, h2s, co2, n2)` | Higher heating value (mass) | BTU/lb | `=HHV_MASS_GAS(85,10,3,1,0,0,0,1,0)` → 22487 |
| `LHV_MASS_GAS(ch4, c2h6, c3h8, c4h10, h2, co, h2s, co2, n2)` | Lower heating value (mass) | BTU/lb | `=LHV_MASS_GAS(85,10,3,1,0,0,0,1,0)` → 20256 |
| `HHV_VOLUME_GAS(ch4, c2h6, c3h8, c4h10, h2, co, h2s, co2, n2)` | Higher heating value (volume) | BTU/scf | `=HHV_VOLUME_GAS(100,0,0,0,0,0,0,0,0)` → 1012 |
| `LHV_VOLUME_GAS(ch4, c2h6, c3h8, c4h10, h2, co, h2s, co2, n2)` | Lower heating value (volume) | BTU/scf | `=LHV_VOLUME_GAS(100,0,0,0,0,0,0,0,0)` → 910 |

---

## Air Requirement Functions

| Function | Description | Units | Example |
|----------|-------------|-------|---------|
| `AIR_REQUIREMENT_MASS(ch4, c2h6, c3h8, c4h10, h2, co, h2s)` | Stoichiometric air (mass) | lb/lb | `=AIR_REQUIREMENT_MASS(100,0,0,0,0,0,0)` → 17.24 |
| `AIR_REQUIREMENT_VOLUME(ch4, c2h6, c3h8, c4h10, h2, co, h2s)` | Stoichiometric air (volume) | scf/scf | `=AIR_REQUIREMENT_VOLUME(100,0,0,0,0,0,0)` → 9.52 |

---

## Products of Combustion

| Function | Description | Units | Example |
|----------|-------------|-------|---------|
| `POC_MASS(ch4, c2h6, c3h8, c4h10, h2, co, h2s, excess_air)` | Products of combustion (mass) | lb/lb | `=POC_MASS(100,0,0,0,0,0,0,15)` → 19.83 |
| `POC_VOLUME(ch4, c2h6, c3h8, c4h10, h2, co, h2s, excess_air)` | Products of combustion (volume) | scf/scf | `=POC_VOLUME(100,0,0,0,0,0,0,15)` → 10.95 |

**Default:** `excess_air = 15%`

---

## Flue Gas Enthalpy

| Function | Description | Units | Example |
|----------|-------------|-------|---------|
| `FLUE_GAS_ENTHALPY(ch4, c2h6, c3h8, c4h10, h2, co, h2s, flue_gas_temp, excess_air, fuel_temp, air_temp)` | Flue gas sensible heat | BTU/lb | `=FLUE_GAS_ENTHALPY(100,0,0,0,0,0,0,350,15,60,60)` → 1847 |

**Defaults:** `flue_gas_temp=350`, `excess_air=15`, `fuel_temp=60`, `air_temp=60`

---

## Steam Properties

| Function | Description | Units | Example |
|----------|-------------|-------|---------|
| `SATURATION_PRESSURE(temperature)` | Saturation pressure | psia | `=SATURATION_PRESSURE(212)` → 14.696 |
| `SATURATION_TEMPERATURE(pressure)` | Saturation temperature | °F | `=SATURATION_TEMPERATURE(14.696)` → 212 |
| `STEAM_ENTHALPY(temperature, pressure, quality)` | Steam enthalpy | BTU/lb | `=STEAM_ENTHALPY(212,14.7,1.0)` → 1150 |
| `STEAM_QUALITY(enthalpy, pressure)` | Steam quality | 0-1 | `=STEAM_QUALITY(650,14.696)` → 0.484 |

**Default:** `quality = 1.0` (saturated vapor)

**Quality values:** 0 = liquid, 1 = vapor, 0-1 = two-phase

---

## Water Properties

| Function | Description | Units | Example |
|----------|-------------|-------|---------|
| `WATER_DENSITY(temperature)` | Density | lb/ft³ | `=WATER_DENSITY(60)` → 62.37 |
| `WATER_VISCOSITY(temperature)` | Dynamic viscosity | lb/ft·s | `=WATER_VISCOSITY(60)` → 0.000752 |
| `WATER_SPECIFIC_HEAT(temperature)` | Specific heat | BTU/lb·°F | `=WATER_SPECIFIC_HEAT(60)` → 0.999 |
| `WATER_THERMAL_CONDUCTIVITY(temperature)` | Thermal conductivity | BTU/hr·ft·°F | `=WATER_THERMAL_CONDUCTIVITY(60)` → 0.340 |

---

## Helper Functions

| Function | Description | Result |
|----------|-------------|--------|
| `HHV_NATURAL_GAS()` | Typical natural gas (85% CH4, 10% C2H6, 3% C3H8, 1% C4H10, 1% CO2) | 22487 BTU/lb |
| `HHV_METHANE()` | Pure methane (100% CH4) | 23875 BTU/lb |

---

## Common Fuel Compositions

| Fuel Type | CH4 | C2H6 | C3H8 | C4H10 | CO2 | N2 | Formula |
|-----------|-----|------|------|-------|-----|----|----|
| Pure Methane | 100 | 0 | 0 | 0 | 0 | 0 | `=HHV_MASS_GAS(100,0,0,0,0,0,0,0,0)` |
| Natural Gas | 85 | 10 | 3 | 1 | 1 | 0 | `=HHV_MASS_GAS(85,10,3,1,0,0,0,1,0)` |
| Pipeline Gas | 95 | 3 | 0.5 | 0.2 | 0.8 | 0.5 | `=HHV_MASS_GAS(95,3,0.5,0.2,0,0,0,0.8,0.5)` |

---

## Units Summary

| Property | Input Units | Output Units |
|----------|-------------|--------------|
| Temperature | °F | °F |
| Pressure | psia | psia |
| Mass composition | % | % |
| Heating value | - | BTU/lb or BTU/scf |
| Air requirement | - | lb/lb or scf/scf |
| POC | - | lb/lb or scf/scf |
| Enthalpy | - | BTU/lb |
| Quality | 0-1 | 0-1 |
| Density | - | lb/ft³ |
| Viscosity | - | lb/ft·s |

---

## Troubleshooting

| Error | Cause | Solution |
|-------|-------|----------|
| `#NAME?` | Function not found | Verify xlwings add-in installed, restart Excel |
| `#VALUE!` | Invalid input | Check numeric values, composition sums to 100% |
| Slow calculation | Python startup | First calculation slower, subsequent faster |
| Module not found | Package not installed | Run `pip install -e /path/to/sigma-thermal` |

---

## Tips

- All fuel compositions must sum to 100%
- Use helper cells for complex formulas
- Set Excel to manual calculation for large sheets
- Temperature range: 32-700°F (water/steam functions)
- Pressure range: 0.1-3000 psia (steam functions)

---

**For full documentation, see:** `EXCEL_UDF_GUIDE.md`

**Web calculators:** `web/resource.html`

**Support:** GTS Energy Inc.
