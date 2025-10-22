# Sigma Thermal Engineering - HTML Calculators

HTML-based calculator interface for professional engineering calculations.

## Structure

```
web/
├── index.html              # Landing page
├── resource.html           # Technical reference (formulas, defaults, validation)
├── css/
│   └── style.css          # Professional styling
├── js/
│   └── heating-value.js   # Calculator logic
└── calculators/
    ├── heating-value.html # Heating value calculator form
    └── [other calculators...]
```

## Features

- **Professional Design**: Clean, minimal, modern interface using Inter font
- **Comprehensive Resources**: Complete technical documentation with all formulas and defaults
- **Form-Based Calculators**: Standard HTML forms that POST to API endpoints
- **Client-Side Validation**: Real-time input validation with JavaScript
- **Fuel Presets**: 9 pre-configured fuel compositions for quick calculations
- **Export Functionality**: Download results as JSON

## Resource Page Contents

The `resource.html` page documents:

1. **Combustion Calculations**
   - Heating values (HHV/LHV) - mass and volume basis
   - Air requirements (stoichiometric and actual)
   - Products of combustion (flue gas composition)
   - Flue gas enthalpy (sensible heat)
   - Combustion efficiency (stack loss method)

2. **Fluids & Steam Calculations**
   - Saturation properties (P↔T)
   - Enthalpy and quality
   - Water transport properties (ρ, μ, Cp, k, Pr)
   - Flash steam generation

3. **Default Parameters**
   - Standard conditions (60°F, 14.696 psia)
   - Reference states
   - Air composition
   - Typical operating parameters by application

4. **Fuel Composition Presets**
   - Pure Methane
   - Natural Gas (Typical, High BTU, Lean)
   - Landfill Gas
   - Digester Gas
   - Refinery Gas
   - Coke Oven Gas
   - Blast Furnace Gas

5. **Validation Data**
   - Excel VBA comparison results
   - ASME Steam Table validation
   - Acceptance criteria

## Using the Calculators

### Viewing Locally

Simply open `index.html` in a web browser to browse the static pages and resource documentation.

### Full Functionality (Requires Backend API)

The calculators submit forms to API endpoints (e.g., `/api/calculate/heating-value`). You need to set up a backend API to process these calculations.

#### Backend API Requirements

The API should accept POST requests with JSON payloads and return calculation results.

**Example: Heating Value Calculator**

Request:
```json
POST /api/calculate/heating-value
Content-Type: application/json

{
  "ch4": 85.0,
  "c2h6": 10.0,
  "c3h8": 3.0,
  "c4h10": 1.0,
  "h2": 0.0,
  "co": 0.0,
  "h2s": 0.0,
  "co2": 1.0,
  "n2": 0.0
}
```

Response:
```json
{
  "hhv_mass": 22487,
  "lhv_mass": 20256,
  "hhv_volume": 1035,
  "lhv_volume": 932,
  "excel_comparison": {
    "hhv": 22490,
    "deviation": 0.013
  }
}
```

### Backend Implementation Options

#### Option 1: Flask API (Python)

```python
from flask import Flask, request, jsonify
from sigma_thermal.combustion import GasComposition, hhv_mass_gas, lhv_mass_gas

app = Flask(__name__)

@app.route('/api/calculate/heating-value', methods=['POST'])
def calculate_heating_value():
    data = request.json

    fuel = GasComposition(
        methane_mass=data['ch4'],
        ethane_mass=data['c2h6'],
        propane_mass=data['c3h8'],
        butane_mass=data['c4h10'],
        hydrogen_mass=data['h2'],
        carbon_monoxide_mass=data['co'],
        hydrogen_sulfide_mass=data['h2s'],
        carbon_dioxide_mass=data['co2'],
        nitrogen_mass=data['n2']
    )

    return jsonify({
        'hhv_mass': hhv_mass_gas(fuel),
        'lhv_mass': lhv_mass_gas(fuel),
        'hhv_volume': hhv_volume_gas(fuel),
        'lhv_volume': lhv_volume_gas(fuel)
    })

if __name__ == '__main__':
    app.run(debug=True)
```

#### Option 2: FastAPI (Python)

```python
from fastapi import FastAPI
from pydantic import BaseModel
from sigma_thermal.combustion import GasComposition, hhv_mass_gas, lhv_mass_gas

app = FastAPI()

class FuelComposition(BaseModel):
    ch4: float
    c2h6: float
    c3h8: float
    c4h10: float
    h2: float
    co: float
    h2s: float
    co2: float
    n2: float

@app.post("/api/calculate/heating-value")
def calculate_heating_value(fuel_data: FuelComposition):
    fuel = GasComposition(
        methane_mass=fuel_data.ch4,
        ethane_mass=fuel_data.c2h6,
        # ... etc
    )

    return {
        'hhv_mass': hhv_mass_gas(fuel),
        'lhv_mass': lhv_mass_gas(fuel),
        # ... etc
    }
```

#### Option 3: Node.js/Express (if porting to JavaScript)

```javascript
const express = require('express');
const app = express();

app.use(express.json());

app.post('/api/calculate/heating-value', (req, res) => {
    const fuel = req.body;

    // Calculate heating values (requires porting Python functions to JS)
    const results = calculateHeatingValues(fuel);

    res.json(results);
});

app.listen(3000);
```

### CORS Configuration

If serving the HTML from a different domain than the API, enable CORS:

```python
# Flask
from flask_cors import CORS
CORS(app)

# FastAPI
from fastapi.middleware.cors import CORSMiddleware
app.add_middleware(CORSMiddleware, allow_origins=["*"])
```

## Deployment Options

### Static Hosting (Pages Only)

- GitHub Pages
- Netlify
- Vercel
- AWS S3 + CloudFront

### Full Stack (HTML + API)

- Heroku (Flask/FastAPI backend)
- AWS Elastic Beanstalk
- Google Cloud Run
- Azure App Service
- Docker container (Nginx + Python API)

## Extending the Calculators

To add a new calculator:

1. **Create HTML form** in `calculators/[calculator-name].html`
   - Use the heating-value.html as a template
   - Include form inputs for all required parameters
   - Add results display section

2. **Create JavaScript** in `js/[calculator-name].js`
   - Handle form submission
   - Validate inputs
   - Display results
   - Include preset management if applicable

3. **Update navigation** in all pages
   - Add link to new calculator in nav menu

4. **Implement API endpoint** in backend
   - Accept POST with calculator inputs
   - Call sigma_thermal functions
   - Return JSON results

5. **Update resource.html** with new calculation formulas

## Testing

### Static Pages
Open `index.html` directly in browser - no server required.

### With Backend
```bash
# Start Flask API
python api.py

# Or FastAPI
uvicorn api:app --reload

# Open in browser
http://localhost:8501/index.html
```

## Browser Compatibility

- Chrome/Edge (latest)
- Firefox (latest)
- Safari (latest)
- Mobile browsers supported via responsive design

## Performance

- No build step required
- Vanilla JavaScript (no frameworks)
- Minimal dependencies
- Fast page loads (<100KB total)

## Security Considerations

1. **Input Validation**: All inputs validated client-side and server-side
2. **CORS**: Configure appropriately for your domain
3. **Rate Limiting**: Consider adding to API endpoints
4. **HTTPS**: Use TLS for production deployments

## License

© 2025 GTS Energy Inc.

## Support

For issues or questions, contact the development team.
