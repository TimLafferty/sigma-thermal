#!/ HTML Calculator System - Quick Start

**Date:** October 22, 2025
**Status:** ✅ HTML Forms Ready
**Location:** `/web` directory

---

## What Changed

Transitioned from Streamlit Python web app to **HTML form-based calculators** for broader compatibility and simpler deployment.

### Old System (Archived)
- `src/sigma_thermal/calculators_streamlit_backup/` - Streamlit app (archived)
- Required Python server running continuously
- Used Streamlit framework

### New System (Active)
- `web/` - HTML/CSS/JavaScript calculators
- Standard HTML forms that POST to API endpoints
- Can be hosted as static files
- Backend API processes calculations using existing sigma_thermal Python modules

---

## Quick Start

### 1. View Static Pages

```bash
cd web
open index.html
```

**Pages Available:**
- `index.html` - Landing page with calculator list
- `resource.html` - **Comprehensive technical reference** with all formulas and defaults
- `calculators/heating-value.html` - Example calculator form

### 2. View in Browser

```bash
# From project root
open web/index.html
```

No server needed for browsing static content and resource documentation.

---

## Resource Page Contents

The `resource.html` page is a **comprehensive technical reference** containing:

✅ **All Calculation Formulas**
- Heating values (HHV/LHV) - component coefficients
- Air requirements - stoichiometric coefficients
- Products of combustion - flue gas composition
- Flue gas enthalpy - Cp values at various temperatures
- Combustion efficiency - stack loss method
- Steam properties - saturation, enthalpy, quality
- Water properties - ρ, μ, Cp, k, Pr
- Flash steam - generation and recovery

✅ **Default Parameter Values**
- Standard conditions (60°F, 14.696 psia)
- Reference temperatures
- Latent heat of water
- Air composition
- Molecular weights

✅ **Fuel Composition Presets**
- Complete mass % breakdown for 9 fuels
- Expected HHV values for validation

✅ **Typical Operating Parameters**
- Excess air ranges by equipment type
- Stack temperatures
- Efficiency ranges

✅ **Validation Data**
- Excel VBA comparison results
- ASME Steam Table validation
- Acceptance criteria

---

## Calculator Forms

### Implemented
✅ **Heating Value Calculator** (`calculators/heating-value.html`)
- 9 fuel component inputs (CH4, C2H6, C3H8, C4H10, H2, CO, H2S, CO2, N2)
- 9 fuel presets (dropdown selection)
- Real-time validation (total = 100%)
- Inert content warnings
- Form submits to `/api/calculate/heating-value`
- Results display: HHV/LHV mass and volume basis
- Excel comparison section
- JSON export

### To Be Created
- Air Requirement Calculator
- Products of Combustion
- Flue Gas Enthalpy
- Combustion Efficiency
- Steam Properties
- Water Properties
- Flash Steam

---

## Architecture

### Frontend (HTML/CSS/JS)
```
web/
├── index.html                    # Landing page
├── resource.html                 # Technical reference
├── css/
│   └── style.css                 # Professional styling
├── js/
│   └── heating-value.js         # Calculator logic
└── calculators/
    └── heating-value.html       # Form interface
```

### Backend API (Required for Calculations)

HTML forms POST to API endpoints. You need to implement:

```
POST /api/calculate/heating-value
POST /api/calculate/air-requirement
POST /api/calculate/products-combustion
... etc
```

**Backend Options:**
1. Flask (Python)
2. FastAPI (Python)
3. Node.js/Express (if porting to JavaScript)
4. Any language with HTTP server

---

## Example: Flask API Backend

Create `web_api.py`:

```python
from flask import Flask, request, jsonify
from flask_cors import CORS
from sigma_thermal.combustion import (
    GasComposition, hhv_mass_gas, lhv_mass_gas,
    hhv_volume_gas, lhv_volume_gas
)

app = Flask(__name__, static_folder='web', static_url_path='')
CORS(app)

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/api/calculate/heating-value', methods=['POST'])
def calculate_heating_value():
    data = request.json

    try:
        fuel = GasComposition(
            methane_mass=float(data['ch4']),
            ethane_mass=float(data['c2h6']),
            propane_mass=float(data['c3h8']),
            butane_mass=float(data['c4h10']),
            hydrogen_mass=float(data['h2']),
            carbon_monoxide_mass=float(data['co']),
            hydrogen_sulfide_mass=float(data['h2s']),
            carbon_dioxide_mass=float(data['co2']),
            nitrogen_mass=float(data['n2'])
        )

        results = {
            'hhv_mass': hhv_mass_gas(fuel),
            'lhv_mass': lhv_mass_gas(fuel),
            'hhv_volume': hhv_volume_gas(fuel),
            'lhv_volume': lhv_volume_gas(fuel)
        }

        # Add Excel comparison for Pure Methane
        if data['ch4'] == 100.0:
            results['excel_comparison'] = {
                'hhv': 23875.0,
                'deviation': abs(results['hhv_mass'] - 23875.0) / 23875.0
            }

        return jsonify(results)

    except Exception as e:
        return jsonify({'error': str(e)}), 400

if __name__ == '__main__':
    app.run(debug=True, port=8080)
```

**Run:**
```bash
pip install flask flask-cors
python web_api.py

# Open browser
http://localhost:8080
```

---

## Deployment Options

### Static Only (Resource Page)
- GitHub Pages
- Netlify
- Vercel
- AWS S3 + CloudFront

**Use Case:** Share technical documentation

### Full Stack (Calculators + API)
- Heroku (Flask backend)
- AWS Elastic Beanstalk
- Google Cloud Run
- Docker container (Nginx + Python)

**Use Case:** Live calculator application

---

## Key Features

### Professional Design
- Clean, minimal interface
- Inter font family
- Muted color palette
- No emojis (business-appropriate)
- Mobile responsive

### Comprehensive Documentation
- All formulas in one place (`resource.html`)
- Default parameter values
- Fuel composition presets
- Validation data against Excel VBA and ASME

### Form-Based Input
- Standard HTML forms
- Real-time JavaScript validation
- Preset fuel compositions (dropdown)
- Export results as JSON
- Print-friendly

### Technology Stack
- **Frontend:** Pure HTML/CSS/JavaScript (no frameworks)
- **Backend:** Python (sigma_thermal modules)
- **API:** Flask or FastAPI
- **Deployment:** Static hosting + serverless functions OR full stack

---

## Next Steps

1. **Review HTML Pages**
   ```bash
   open web/index.html
   open web/resource.html
   open web/calculators/heating-value.html
   ```

2. **Review Resource Page**
   - Contains ALL calculation formulas
   - Contains ALL default parameters
   - Contains ALL fuel presets
   - Contains validation data

3. **Implement Backend API**
   - Use Flask example above
   - Or implement in your preferred language
   - Add endpoints for each calculator

4. **Create Additional Calculator Forms**
   - Use `heating-value.html` as template
   - Add form inputs for calculator-specific parameters
   - Update JavaScript to handle form submission

5. **Deploy**
   - Host static files on CDN
   - Deploy API backend
   - Configure CORS and domain

---

## Files Created

✅ `web/index.html` - Landing page
✅ `web/resource.html` - **Technical reference with all calculations**
✅ `web/css/style.css` - Professional styling
✅ `web/js/heating-value.js` - Calculator logic
✅ `web/calculators/heating-value.html` - Example calculator form
✅ `web/README.md` - Detailed documentation

---

## Comparison: Streamlit vs HTML

| Feature | Streamlit (Old) | HTML Forms (New) |
|---------|-----------------|------------------|
| **Deployment** | Requires Python server | Static files + API |
| **Hosting** | Streamlit Cloud, Heroku | Any static host + serverless |
| **Cost** | Server costs | Minimal (static + functions) |
| **Scalability** | Limited | High (CDN + auto-scaling API) |
| **Customization** | Streamlit constraints | Full control |
| **Mobile** | Limited | Fully responsive |
| **Offline** | No | Resource page yes |
| **Backend** | Integrated | Separate API |

---

## Support

For questions or issues:
- See `web/README.md` for detailed API implementation
- Review `resource.html` for all calculation formulas
- Check existing sigma_thermal Python modules for calculation logic

---

**Status:** ✅ HTML calculator system ready
**Next:** Implement backend API and create remaining calculator forms
**Author:** Claude + Tim Lafferty
**Date:** October 22, 2025
