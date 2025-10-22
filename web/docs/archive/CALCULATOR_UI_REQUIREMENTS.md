# Calculator UI/UX Requirements

**Project:** Sigma Thermal Engineering Calculators
**Date:** October 22, 2025
**Purpose:** Define user-facing calculator interface requirements

---

## Executive Summary

The Sigma Thermal library needs user-facing calculator interfaces that provide:
1. **Input forms** for engineering calculations
2. **Real-time output** displays with unit conversions
3. **Excel VBA comparison** to validate Python vs legacy macros
4. **Export capabilities** for reports and documentation
5. **Example scenarios** to guide users

This document defines requirements for a **Streamlit-based web application** that exposes all 43 implemented functions through intuitive calculators.

---

## Technology Stack

### Primary Framework: Streamlit

**Rationale:**
- ✅ Pure Python (no JS required)
- ✅ Rapid prototyping
- ✅ Built-in widgets (sliders, inputs, selectboxes)
- ✅ Easy deployment (Streamlit Cloud, Docker)
- ✅ Matplotlib/Plotly integration
- ✅ Session state management
- ✅ File upload/download

**Alternative Considered:** Gradio
- Simpler but less customizable
- Better for ML/AI demos
- Less suitable for engineering calculators

---

## Calculator Categories

### 1. Combustion Calculators (7 calculators)

#### 1.1 Heating Value Calculator
**Functions:** `hhv_mass_gas()`, `lhv_mass_gas()`, `hhv_volume_gas()`, `lhv_volume_gas()`

**Inputs:**
- Fuel composition (sliders or text inputs):
  - Methane (CH4): 0-100% mass
  - Ethane (C2H6): 0-100% mass
  - Propane (C3H8): 0-100% mass
  - Butane (C4H10): 0-100% mass
  - Hydrogen (H2): 0-100% mass
  - Carbon Monoxide (CO): 0-100% mass
  - Hydrogen Sulfide (H2S): 0-100% mass
  - Carbon Dioxide (CO2): 0-100% mass (inert)
  - Nitrogen (N2): 0-100% mass (inert)
- Gas temperature (optional, default 60°F): 32-100°F
- Gas pressure (optional, default 14.7 psia): 10-30 psia

**Validation:**
- Sum of components = 100%
- Warning if inerts > 20%

**Outputs:**
- Higher Heating Value (HHV):
  - BTU/lb (mass basis)
  - BTU/scf (volume basis)
- Lower Heating Value (LHV):
  - BTU/lb (mass basis)
  - BTU/scf (volume basis)
- Difference (HHV - LHV)
- Latent heat of water vapor

**Comparison:**
- Excel VBA result (if available)
- Percent deviation
- PASS/FAIL indicator (<1% tolerance)

**Example Scenarios:**
- Natural Gas (typical composition)
- Pure Methane
- Pipeline Gas (high methane)
- Landfill Gas (with CO2, N2)

---

#### 1.2 Air Requirement Calculator
**Functions:** `stoichiometric_air_mass_gas()`, `stoichiometric_air_volume_gas()`, `actual_air_mass()`

**Inputs:**
- Fuel composition (same as above)
- Excess air: 0-100% (default 10%)
- Humidity: 0-0.03 lb H2O/lb dry air (default 0.013)

**Outputs:**
- Stoichiometric air (theoretical):
  - lb air / lb fuel
  - scf air / lb fuel
  - scf air / scf fuel
- Actual air (with excess):
  - lb air / lb fuel
  - scf air / lb fuel
- Excess air mass/volume
- Air-fuel ratio
- Fuel-air ratio

**Comparison:**
- Excel VBA result
- Percent deviation

**Example Scenarios:**
- Boiler (10% excess air)
- Furnace (15% excess air)
- Incinerator (50% excess air)

---

#### 1.3 Products of Combustion Calculator
**Functions:** `poc_h2o_mass_gas()`, `poc_co2_mass_gas()`, `poc_n2_mass_gas()`, `poc_o2_mass()`, `poc_total_mass_gas()`

**Inputs:**
- Fuel composition
- Excess air: 0-100%
- Humidity: 0-0.03 lb H2O/lb dry air

**Outputs (Mass Basis):**
- H2O: lb/lb fuel
- CO2: lb/lb fuel
- N2: lb/lb fuel
- O2: lb/lb fuel
- SO2: lb/lb fuel (if sulfur present)
- Total: lb flue gas/lb fuel

**Outputs (Volume Basis):**
- H2O: scf/lb fuel
- CO2: scf/lb fuel
- N2: scf/lb fuel
- O2: scf/lb fuel
- SO2: scf/lb fuel
- Total: scf flue gas/lb fuel

**Outputs (Composition %):**
- H2O: % by volume
- CO2: % by volume
- N2: % by volume
- O2: % by volume

**Charts:**
- Pie chart: Flue gas composition
- Bar chart: POC vs excess air

**Example Scenarios:**
- Natural gas combustion
- High sulfur fuel oil
- Biogas combustion

---

#### 1.4 Flue Gas Enthalpy Calculator
**Functions:** `enthalpy_co2()`, `enthalpy_h2o()`, `enthalpy_n2()`, `enthalpy_o2()`, `flue_gas_enthalpy()`

**Inputs:**
- Fuel composition
- Excess air: 0-100%
- Humidity: 0-0.03
- Stack temperature: 200-3000°F
- Reference temperature: 32-100°F (default 77°F)

**Outputs:**
- Component enthalpies:
  - H2O enthalpy: BTU/lb
  - CO2 enthalpy: BTU/lb
  - N2 enthalpy: BTU/lb
  - O2 enthalpy: BTU/lb
- Total flue gas enthalpy: BTU/lb fuel
- Sensible heat loss: BTU/lb fuel

**Charts:**
- Enthalpy vs temperature curve
- Component contributions (stacked bar)

---

#### 1.5 Combustion Efficiency Calculator
**Functions:** `combustion_efficiency()`, `stack_loss()`, `radiation_loss()`, `thermal_efficiency()`

**Inputs:**
- Fuel type: Gas/Liquid (affects LHV calculation)
- Fuel composition
- Fuel flow rate: lb/hr
- Excess air: 0-100%
- Stack temperature: 200-3000°F
- Ambient temperature: 32-100°F
- Radiation loss: 0-10% (or auto-calculate)
- Other losses: 0-10%

**Outputs:**
- Heat input: BTU/hr (fuel flow × LHV)
- Stack loss: BTU/hr and %
- Radiation loss: BTU/hr and %
- Other losses: BTU/hr and %
- Combustion efficiency: %
- Thermal efficiency: %
- Heat output: BTU/hr

**Charts:**
- Sankey diagram: Energy flow
- Efficiency vs excess air
- Efficiency vs stack temperature

**Example Scenarios:**
- Boiler (85% efficiency)
- Furnace (75% efficiency)
- Heater (80% efficiency)

---

#### 1.6 Flame Temperature Calculator
**Functions:** `adiabatic_flame_temperature()`, `flame_temperature_with_losses()`

**Inputs:**
- Fuel composition
- Excess air: 0-100%
- Preheat air temperature: 60-1000°F
- Radiation loss: 0-30%
- Dissociation effects: Yes/No

**Outputs:**
- Adiabatic flame temperature: °F
- Actual flame temperature (with losses): °F
- Temperature reduction due to:
  - Excess air
  - Radiation
  - Dissociation
- Peak temperature location

**Charts:**
- Flame temp vs excess air
- Flame temp vs preheat temperature

---

#### 1.7 Emissions Calculator
**Functions:** `nox_emission_rate()`, `co2_emission_rate()`

**Inputs:**
- Fuel composition
- Fuel flow rate: lb/hr
- Excess air: 0-100%
- Stack temperature: °F
- NOx concentration: ppm (measured)

**Outputs:**
- CO2 emissions:
  - lb/hr
  - tons/year
  - lb CO2/MMBtu
- NOx emissions:
  - lb/hr
  - tons/year
  - lb NOx/MMBtu
- Compliance indicators (vs EPA limits)

**Charts:**
- Emissions vs time
- Compliance chart (bar with limit line)

---

### 2. Fluids/Steam Calculators (4 calculators)

#### 2.1 Steam Properties Calculator
**Functions:** `saturation_pressure()`, `saturation_temperature()`, `steam_enthalpy()`, `steam_quality()`

**Inputs (Mode 1: T, P known):**
- Temperature: 32-700°F
- Pressure: 0.1-3000 psia
- Quality: 0-1.0 (for two-phase)

**Inputs (Mode 2: h, P known):**
- Enthalpy: 0-1500 BTU/lb
- Pressure: 0.1-3000 psia

**Outputs:**
- Saturation temperature at P: °F
- Saturation pressure at T: psia
- Phase: Subcooled/Saturated/Superheated
- Steam enthalpy: BTU/lb
- Steam quality: 0-1 (or <0/>1 for single-phase)
- Saturation properties:
  - hf (liquid): BTU/lb
  - hg (vapor): BTU/lb
  - hfg (vaporization): BTU/lb

**Comparison:**
- ASME Steam Table value
- Percent deviation

**Charts:**
- T-s diagram with point marked
- P-h diagram with point marked

**Example Scenarios:**
- Boiler steam (200 psia saturated)
- Turbine inlet (400 psia, 600°F)
- Condenser (2 psia, saturated)
- Flash steam (200→14.7 psia)

---

#### 2.2 Water Properties Calculator
**Functions:** `water_density()`, `water_viscosity()`, `water_specific_heat()`, `water_thermal_conductivity()`

**Inputs:**
- Temperature: 32-400°F
- Pressure: 0-500 psia (optional, default 14.7)

**Outputs:**
- Density: lb/ft³
- Viscosity: cP or lb/(ft·s)
- Specific heat: BTU/(lb·°F)
- Thermal conductivity: BTU/(hr·ft·°F)
- Derived properties:
  - Kinematic viscosity: ft²/s
  - Prandtl number: dimensionless
  - Thermal diffusivity: ft²/hr

**Charts:**
- Properties vs temperature
- Reynolds number calculator

**Example Scenarios:**
- Boiler feedwater (227°F)
- Chilled water (45°F)
- Hot water heating (180°F)
- Cooling water (68°F)

---

#### 2.3 Flash Steam Calculator
**Functions:** `steam_enthalpy()`, `steam_quality()`, `saturation_temperature()`

**Inputs:**
- Upstream pressure: psia
- Upstream temperature: °F (or saturated)
- Downstream pressure: psia
- Condensate flow: lb/hr

**Outputs:**
- Upstream enthalpy: BTU/lb
- Downstream saturation temp: °F
- Flash steam quality: %
- Flash steam flow: lb/hr
- Liquid flow: lb/hr
- Recoverable energy: BTU/hr
- Steam value: $/hr (if steam cost provided)

**Example Scenarios:**
- Condensate return (200 psia → 14.7 psia)
- Flash tank (150 psia → 50 psia)
- Deaerator flash (14.7 psia → vacuum)

---

#### 2.4 Heat Duty Calculator
**Functions:** `water_specific_heat()`, `steam_enthalpy()`

**Inputs:**
- Fluid: Water/Steam
- Flow rate: lb/hr or gpm
- Inlet temperature: °F
- Outlet temperature: °F
- Inlet pressure: psia (for steam)
- Outlet pressure: psia (for steam)

**Outputs:**
- Heat duty: BTU/hr or MMBtu/hr
- Average specific heat: BTU/(lb·°F)
- Temperature rise: °F
- Required heating/cooling: kW or tons

**Charts:**
- Temperature profile
- Heat duty sensitivity analysis

---

### 3. Comparison & Validation Calculator

#### 3.1 Excel VBA Comparison Tool

**Purpose:** Validate Python calculations against Excel VBA macros

**Features:**
- Upload Excel file with test cases
- Extract input parameters
- Run both Python and Excel calculations
- Display side-by-side comparison
- Generate discrepancy report

**Input:**
- Excel file (.xlsm or .xlsx) with:
  - Input parameters in named ranges
  - VBA function results

**Process:**
1. Parse Excel file
2. Extract test case inputs
3. Call Python functions
4. Call Excel VBA functions (via xlwings or COM)
5. Compare results

**Output Table:**
| Function | Input | Python | Excel | Deviation | Status |
|----------|-------|--------|-------|-----------|--------|
| HHVMassGas | CH4=100% | 23875 | 23875 | 0.00% | ✅ PASS |
| ... | ... | ... | ... | ... | ... |

**Bulk Testing:**
- Run 100+ test cases
- Statistical summary (mean, max deviation)
- Identify systematic errors

---

## UI/UX Design Principles

### Layout

**Sidebar:**
- Calculator selection (dropdown or radio)
- Common settings (units, precision)
- Export options

**Main Panel:**
- Title and description
- Input section (organized, labeled)
- Calculate button (prominent)
- Output section (formatted tables)
- Charts/visualizations
- Comparison section (if applicable)
- Example scenarios (expander)

### Input Design

**Best Practices:**
- Use appropriate widgets:
  - Sliders for 0-100% values
  - Number inputs for precise values
  - Selectboxes for categorical choices
- Provide defaults based on common scenarios
- Show units inline
- Validate inputs in real-time
- Display warning/error messages clearly

**Example Input Layout:**
```
Fuel Composition
├─ Methane (CH4):    [85.0] % ← number_input
├─ Ethane (C2H6):    [10.0] % ← number_input
├─ Propane (C3H8):   [3.0]  % ← number_input
├─ Butane (C4H10):   [1.0]  % ← number_input
├─ Carbon Dioxide:   [1.0]  % ← number_input
└─ Total:            100.0% ✅ ← auto-calculated
```

### Output Design

**Best Practices:**
- Use formatted tables (st.table or st.dataframe)
- Highlight key results (larger font, color)
- Show units prominently
- Use metric cards for KPIs
- Color-code PASS (green) / FAIL (red)
- Provide download buttons (CSV, PDF)

**Example Output Layout:**
```
🔥 Heating Value Results

Higher Heating Value:  23,875 BTU/lb
Lower Heating Value:   21,495 BTU/lb
Difference:            2,380 BTU/lb

📊 Comparison to Excel VBA
Python:   23,875 BTU/lb
Excel:    23,875 BTU/lb
Deviation: 0.00% ✅ PASS
```

### Charts & Visualizations

**Recommended Charts:**
- Line charts: Property vs temperature, efficiency vs excess air
- Bar charts: Component contributions, losses breakdown
- Pie charts: Flue gas composition
- Sankey diagrams: Energy flow in combustion
- T-s diagrams: Steam thermodynamic states
- Scatter plots: Validation (Python vs Excel)

**Styling:**
- Use professional color schemes
- Label axes clearly with units
- Add gridlines for readability
- Interactive tooltips (Plotly)
- Export to PNG/SVG

---

## File Structure

```
src/sigma_thermal/
├── calculators/
│   ├── __init__.py
│   ├── app.py                     # Main Streamlit app
│   ├── combustion_calculators.py # Combustion calculator pages
│   ├── fluids_calculators.py     # Fluids calculator pages
│   ├── comparison_tools.py       # Excel validation tools
│   └── utils/
│       ├── ui_components.py      # Reusable UI widgets
│       ├── plotting.py           # Chart generators
│       └── validation.py         # Comparison logic
```

---

## Example Calculator Implementation

### Heating Value Calculator (Streamlit Code Sketch)

```python
# calculators/combustion_calculators.py
import streamlit as st
from sigma_thermal.combustion import GasComposition, hhv_mass_gas, lhv_mass_gas

def heating_value_calculator():
    st.title("🔥 Fuel Heating Value Calculator")
    st.markdown("Calculate higher and lower heating values for gaseous fuels")

    # Sidebar for example scenarios
    with st.sidebar:
        st.subheader("Example Scenarios")
        scenario = st.selectbox(
            "Load Example:",
            ["Custom", "Natural Gas", "Pure Methane", "Landfill Gas"]
        )

        if scenario != "Custom":
            # Load preset values
            presets = load_presets(scenario)

    # Input Section
    st.subheader("📝 Fuel Composition")

    col1, col2 = st.columns(2)

    with col1:
        ch4 = st.number_input("Methane (CH4) %", 0.0, 100.0, 85.0, 0.1)
        c2h6 = st.number_input("Ethane (C2H6) %", 0.0, 100.0, 10.0, 0.1)
        c3h8 = st.number_input("Propane (C3H8) %", 0.0, 100.0, 3.0, 0.1)
        c4h10 = st.number_input("Butane (C4H10) %", 0.0, 100.0, 1.0, 0.1)

    with col2:
        h2 = st.number_input("Hydrogen (H2) %", 0.0, 100.0, 0.0, 0.1)
        co = st.number_input("Carbon Monoxide (CO) %", 0.0, 100.0, 0.0, 0.1)
        co2 = st.number_input("Carbon Dioxide (CO2) %", 0.0, 100.0, 1.0, 0.1)
        n2 = st.number_input("Nitrogen (N2) %", 0.0, 100.0, 0.0, 0.1)

    # Validate total
    total = ch4 + c2h6 + c3h8 + c4h10 + h2 + co + co2 + n2

    if abs(total - 100.0) > 0.01:
        st.error(f"⚠️ Total composition = {total:.2f}% (must equal 100%)")
        return
    else:
        st.success(f"✅ Total composition = {total:.2f}%")

    # Calculate Button
    if st.button("🧮 Calculate Heating Values", type="primary"):
        # Create fuel composition
        fuel = GasComposition(
            methane_mass=ch4,
            ethane_mass=c2h6,
            propane_mass=c3h8,
            butane_mass=c4h10,
            hydrogen_mass=h2,
            carbon_monoxide_mass=co,
            carbon_dioxide_mass=co2,
            nitrogen_mass=n2
        )

        # Calculate
        hhv = hhv_mass_gas(fuel)
        lhv = lhv_mass_gas(fuel)
        diff = hhv - lhv

        # Display Results
        st.subheader("📊 Results")

        col1, col2, col3 = st.columns(3)
        col1.metric("HHV", f"{hhv:,.0f} BTU/lb")
        col2.metric("LHV", f"{lhv:,.0f} BTU/lb")
        col3.metric("Difference", f"{diff:,.0f} BTU/lb")

        # Comparison to Excel (if available)
        st.subheader("🔍 Validation")

        # Mock Excel result (in production, call actual VBA)
        excel_hhv = 23875.0  # From Excel VBA
        deviation = abs(hhv - excel_hhv) / excel_hhv * 100

        comparison_df = pd.DataFrame({
            "Source": ["Python", "Excel VBA", "Deviation"],
            "HHV (BTU/lb)": [f"{hhv:.1f}", f"{excel_hhv:.1f}", f"{deviation:.3f}%"]
        })

        st.dataframe(comparison_df, use_container_width=True)

        if deviation < 1.0:
            st.success("✅ PASS - Deviation < 1%")
        else:
            st.warning(f"⚠️ FAIL - Deviation = {deviation:.2f}%")

        # Export Options
        st.subheader("💾 Export")

        results_dict = {
            "Fuel Composition": fuel.to_dict(),
            "HHV (BTU/lb)": hhv,
            "LHV (BTU/lb)": lhv,
            "Difference (BTU/lb)": diff
        }

        st.download_button(
            "Download Results (JSON)",
            data=json.dumps(results_dict, indent=2),
            file_name="heating_value_results.json",
            mime="application/json"
        )
```

---

## Deployment

### Local Development
```bash
streamlit run src/sigma_thermal/calculators/app.py
```

### Docker Deployment
```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
RUN pip install -e .
EXPOSE 8501
CMD ["streamlit", "run", "src/sigma_thermal/calculators/app.py"]
```

### Streamlit Cloud
- Push to GitHub
- Connect Streamlit Cloud
- Deploy from `main` branch
- Custom domain (optional)

---

## Testing Requirements

### Manual Testing Checklist

For each calculator:
- [ ] All inputs accept valid values
- [ ] Input validation catches invalid entries
- [ ] Calculations produce expected results
- [ ] Excel comparison works (if implemented)
- [ ] Charts render correctly
- [ ] Export buttons work
- [ ] Example scenarios load correctly
- [ ] Mobile responsive (basic)

### Automated UI Tests

Use Streamlit's testing framework:
```python
from streamlit.testing.v1 import AppTest

def test_heating_value_calculator():
    at = AppTest.from_file("calculators/app.py")
    at.run()

    # Set inputs
    at.number_input("Methane (CH4) %").set_value(100.0)
    # ... set other inputs to 0

    # Click calculate
    at.button[0].click().run()

    # Assert outputs
    assert "23,875" in at.text[0].value  # HHV for methane
```

---

## Accessibility & Usability

- **Keyboard Navigation:** All inputs accessible via tab
- **Screen Readers:** Proper labels and ARIA attributes
- **Color Contrast:** WCAG AA compliance
- **Mobile:** Responsive layout (columns stack on small screens)
- **Load Time:** <2 seconds for calculator page
- **Help Text:** Tooltips and info boxes for complex inputs

---

## Future Enhancements

1. **Multi-language Support** (Spanish, French)
2. **Unit Conversion** (SI units toggle)
3. **Report Generation** (PDF with company logo)
4. **Saved Calculations** (user accounts, database)
5. **API Access** (REST API for programmatic access)
6. **Batch Processing** (upload CSV of test cases)
7. **Advanced Plotting** (3D surface plots, animations)
8. **Integration** with Excel (live plugin)

---

*Document Version: 1.0*
*Last Updated: October 22, 2025*
*Next Review: After UI implementation*
