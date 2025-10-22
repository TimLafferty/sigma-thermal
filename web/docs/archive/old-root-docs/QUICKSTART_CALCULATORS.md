# Quick Start: Sigma Thermal Calculators

## ✅ Dependencies Installed

All required packages are now installed in your virtual environment (.venv):
- ✅ streamlit 1.50.0
- ✅ plotly 6.3.1
- ✅ pandas, numpy, etc.

---

## 🚀 Running the Calculator App

### Option 1: Use the Launch Script (Recommended)

```bash
./run_calculators.sh
```

### Option 2: Activate venv and run streamlit

```bash
source .venv/bin/activate
streamlit run src/sigma_thermal/calculators/app.py
```

### Option 3: Run directly with venv python

```bash
.venv/bin/streamlit run src/sigma_thermal/calculators/app.py
```

---

## 🌐 Access the App

The app will automatically open in your browser at:
```
http://localhost:8501
```

If it doesn't open automatically, copy that URL into your browser.

---

## 🎯 Available Calculators

### ✅ Fully Functional:

1. **🔥 Heating Value Calculator**
   - Calculate HHV/LHV for gaseous fuels
   - 9 fuel presets (Natural Gas, Methane, Landfill Gas, etc.)
   - Excel VBA comparison
   - JSON export

2. **💧 Steam Properties Calculator**
   - Calculate saturation properties (P↔T)
   - Steam enthalpy and quality
   - Phase determination (subcooled/saturated/superheated)
   - Interactive T-s diagram
   - ASME Steam Table validation

### 🚧 Coming Soon (Placeholders):

3. Air Requirement Calculator
4. Products of Combustion
5. Flue Gas Enthalpy
6. Combustion Efficiency
7. Water Properties
8. Flash Steam Calculator
9. Excel Comparison Tool

---

## 📖 How to Use

### Heating Value Calculator Example:

1. Click "🔥 Heating Value Calculator" in sidebar
2. Select a fuel preset (e.g., "Natural Gas (Typical)")
3. Composition automatically fills in
4. Adjust percentages if needed (must total 100%)
5. Click "🧮 Calculate Heating Values"
6. View results:
   - HHV: ~22,487 BTU/lb
   - LHV: ~20,256 BTU/lb
7. Export as JSON if needed

### Steam Properties Calculator Example:

1. Click "💧 Steam Properties" in sidebar
2. Select mode: "Temperature & Pressure (Known)"
3. Enter:
   - Temperature: 212°F
   - Pressure: 14.7 psia (or select "Atmospheric" preset)
   - Quality: 1.0 (saturated vapor)
4. Click "🧮 Calculate Steam Properties"
5. View results:
   - Phase: Saturated (Two-Phase)
   - Enthalpy: ~1156 BTU/lb
   - Saturation properties table
   - T-s diagram with point marked
6. Compare to ASME Steam Tables (automatic at reference points)

---

## ⚙️ Settings

Access settings in the sidebar:
- **Decimal Places:** 0-6 (default: 2)
- **Unit System:** US Customary (SI coming soon)

---

## 🛑 Stopping the Server

Press `Ctrl+C` in the terminal to stop the Streamlit server.

---

## 🐛 Troubleshooting

### "ModuleNotFoundError: No module named 'streamlit'"

**Solution:** Make sure you're using the virtual environment:

```bash
# Activate venv first
source .venv/bin/activate

# Then run
streamlit run src/sigma_thermal/calculators/app.py
```

Or use the launch script which handles this automatically:
```bash
./run_calculators.sh
```

### Port 8501 Already in Use

**Solution:** Kill the existing Streamlit process or use a different port:

```bash
.venv/bin/streamlit run src/sigma_thermal/calculators/app.py --server.port 8502
```

### App Not Opening in Browser

**Solution:** Manually navigate to `http://localhost:8501` in your browser.

---

## 📚 Documentation

- **Full Calculator Docs:** `docs/CALCULATOR_UI_REQUIREMENTS.md`
- **Implementation Details:** `docs/UI_IMPLEMENTATION_SUMMARY.md`
- **Overall Progress:** `docs/COMPREHENSIVE_PROGRESS.md`
- **Excel Validation:** `docs/EXCEL_VBA_DISCREPANCIES.md`

---

## 🎨 Features Demonstrated

**Heating Value Calculator:**
- ✅ Multi-component fuel inputs
- ✅ Preset management
- ✅ Real-time validation
- ✅ Excel comparison
- ✅ Professional styling

**Steam Properties Calculator:**
- ✅ Multi-mode interface
- ✅ Phase determination
- ✅ Interactive Plotly charts
- ✅ ASME validation
- ✅ Color-coded results

---

## 💻 Development

To add a new calculator:

1. Create `pages/your_calculator.py`
2. Import sigma_thermal functions
3. Use UI components from `utils/ui_components.py`
4. Add navigation link in `app.py` sidebar
5. Test and deploy

Example template in each placeholder file.

---

## 🚀 Next Steps

**Try the calculators:**
```bash
./run_calculators.sh
```

**Then navigate to:**
- 🔥 Heating Value Calculator
- 💧 Steam Properties Calculator

**Enjoy!** 🎉

---

*Last Updated: October 22, 2025*
*Status: 2 calculators operational, 7 coming soon*
