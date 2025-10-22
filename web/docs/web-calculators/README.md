# Web Calculators Documentation

**Professional HTML-based thermal engineering calculators**

---

## Overview

The Sigma Thermal web calculators provide a modern, professional interface for thermal engineering calculations:

- ✅ **Clean, minimal design** - Professional UI with Inter font
- ✅ **HTML forms** - No framework dependencies
- ✅ **Real-time validation** - Input validation and total checking
- ✅ **Fuel presets** - 9 pre-configured fuel compositions
- ✅ **API integration** - Backend calculations via Azure Functions
- ✅ **Mobile responsive** - Works on desktop and mobile
- ✅ **Technical reference** - Comprehensive formulas and defaults

---

## Quick Start

### Access Live Calculators

Deploy to Azure Static Web Apps:
```bash
./deploy-azure.sh
```

Then access at:
```
https://[your-app].azurestaticapps.net
```

### Run Locally

```bash
# Install Azure Static Web Apps CLI
npm install -g @azure/static-web-apps-cli

# Start local server
swa start web --api-location api
```

Access at: `http://localhost:4280`

---

## Documentation Files

### [HTML Calculators](html-calculators.md)

**Complete documentation for the web interface**

Topics covered:
- Architecture and design
- Calculator pages
- Form handling
- API integration
- Styling and design system
- Fuel presets
- Technical reference page

**Full details on the web interface.**

---

## Available Calculators

### Heating Value Calculator

**URL:** `/calculators/heating-value.html`

**Features:**
- 9 fuel composition inputs (CH4, C2H6, C3H8, C4H10, H2, CO, H2S, CO2, N2)
- Real-time total validation (must sum to 100%)
- 9 fuel presets (natural gas, pure methane, pipeline gas, etc.)
- Calculates HHV/LHV on mass and volume basis
- Compares to Excel VBA results

**API Endpoint:** `POST /api/calculate/heating-value`

### More Calculators (Coming Soon)

- Air Requirement Calculator
- Products of Combustion Calculator
- Steam Properties Calculator
- Water Properties Calculator
- Flash Steam Calculator

---

## Page Structure

### Landing Page (`index.html`)

**Purpose:** Entry point for the application

**Features:**
- Professional header with navigation
- Quick stats (43 functions, 412 tests, <1% accuracy)
- Calculator cards organized by category
- Links to individual calculators
- Footer with technical info

### Resource Page (`resource.html`)

**Purpose:** Technical reference and documentation

**Features:**
- Table of contents with anchor links
- All calculation formulas with proper notation
- Component coefficient tables
- Default parameter values
- 9 fuel composition presets
- Validation results (Excel comparison, ASME validation)
- Smooth scrolling navigation

### Calculator Pages (`calculators/*.html`)

**Purpose:** Interactive calculation forms

**Features:**
- Form inputs with validation
- Fuel preset dropdown
- Real-time total checking
- Submit button (disabled until valid)
- Results display with formatted values
- Comparison to Excel VBA

---

## Design System

### Typography

```css
Font Family: Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif
Body Text: 16px, #212529
Headers: 600-700 weight, #212529
Code: 'SF Mono', 'Monaco', monospace
```

### Colors

```css
Primary Text: #212529
Secondary Text: #495057
Borders: #dee2e6
Background: #ffffff
Accent: #0066cc
Success: #28a745
Warning: #ffc107
Error: #dc3545
```

### Components

```css
Cards: 1px solid #dee2e6, 8px border-radius, 1.5rem padding
Buttons: Primary (#0066cc), padding 0.75rem 1.5rem
Inputs: border #ced4da, focus ring #0066cc
Tables: Striped rows, hover effect
```

**Full design details:** [HTML Calculators](html-calculators.md#design-system)

---

## Fuel Presets

### Available Presets

1. **Pure Methane** - 100% CH4
2. **Natural Gas (Typical)** - 85% CH4, 10% C2H6, 3% C3H8, 1% C4H10, 1% CO2
3. **Pipeline Gas** - 95% CH4, 3% C2H6, 0.5% C3H8, 0.2% C4H10, 0.8% CO2, 0.5% N2
4. **Wellhead Gas** - 80% CH4, 12% C2H6, 5% C3H8, 2% C4H10, 1% CO2
5. **Associated Gas** - 70% CH4, 15% C2H6, 8% C3H8, 4% C4H10, 2% CO2, 1% N2
6. **Lean Gas** - 98% CH4, 1% C2H6, 0.5% CO2, 0.5% N2
7. **Rich Gas** - 75% CH4, 15% C2H6, 6% C3H8, 3% C4H10, 1% CO2
8. **Biogas** - 60% CH4, 35% CO2, 5% N2
9. **Landfill Gas** - 50% CH4, 40% CO2, 10% N2

**Usage:**
```javascript
// Select preset from dropdown
document.getElementById('fuel-preset').value = 'natural-gas-typical';
// Trigger change event to load values
document.getElementById('fuel-preset').dispatchEvent(new Event('change'));
```

---

## API Integration

### Request Format

```javascript
const data = {
    ch4: parseFloat(document.getElementById('ch4').value) || 0,
    c2h6: parseFloat(document.getElementById('c2h6').value) || 0,
    c3h8: parseFloat(document.getElementById('c3h8').value) || 0,
    c4h10: parseFloat(document.getElementById('c4h10').value) || 0,
    h2: parseFloat(document.getElementById('h2').value) || 0,
    co: parseFloat(document.getElementById('co').value) || 0,
    h2s: parseFloat(document.getElementById('h2s').value) || 0,
    co2: parseFloat(document.getElementById('co2').value) || 0,
    n2: parseFloat(document.getElementById('n2').value) || 0
};

const response = await fetch('/api/calculate/heating-value', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(data)
});

const results = await response.json();
```

### Response Format

```json
{
  "hhv_mass": 23389.33,
  "lhv_mass": 20256.45,
  "hhv_volume": 1009.1,
  "lhv_volume": 907.8,
  "excel_comparison": {
    "hhv": 23875.0,
    "deviation": 0.0013
  }
}
```

---

## Form Validation

### Real-Time Total Validation

```javascript
function validateTotal() {
    const total = ch4 + c2h6 + c3h8 + c4h10 + h2 + co + h2s + co2 + n2;
    const validationDiv = document.getElementById('total-validation');
    const calculateBtn = document.getElementById('calculate-btn');

    if (Math.abs(total - 100) < 0.01) {
        validationDiv.className = 'status status-pass';
        validationDiv.innerHTML = '<strong>Total:</strong> ' + total.toFixed(1) + '% ✓ Valid';
        calculateBtn.disabled = false;
    } else {
        validationDiv.className = 'status status-fail';
        validationDiv.innerHTML = '<strong>Total:</strong> ' + total.toFixed(1) + '% ⚠ Must equal 100%';
        calculateBtn.disabled = true;
    }
}
```

### Input Constraints

```html
<input type="number"
       id="ch4"
       name="ch4"
       min="0"
       max="100"
       step="0.1"
       value="85.0"
       oninput="validateTotal()">
```

---

## File Structure

```
web/
├── index.html                    # Landing page
├── resource.html                 # Technical reference
├── css/
│   └── style.css                 # Global styles
├── js/
│   └── heating-value.js          # Calculator logic
└── calculators/
    └── heating-value.html        # Heating value calculator
```

---

## Deployment

### Azure Static Web Apps

Located in `web/` directory, configured via:
- `staticwebapp.config.json` - Routing and CORS
- `.github/workflows/azure-static-web-apps.yml` - CI/CD

**Deploy:**
```bash
./deploy-azure.sh
```

**Full guide:** [Azure Deployment](../azure-deployment/README.md)

---

## Local Development

### Run Web Server

```bash
# Using Python
cd web
python3 -m http.server 8000
```

Access at: `http://localhost:8000`

### Run with API

```bash
# Using Azure Static Web Apps CLI
npm install -g @azure/static-web-apps-cli
swa start web --api-location api
```

Access at: `http://localhost:4280`

---

## Styling Guidelines

### Professional Design Principles

1. **Minimal** - Clean, uncluttered layout
2. **Modern** - Inter font, subtle shadows
3. **Intuitive** - Clear labels, logical flow
4. **Professional** - No emojis in navigation or buttons
5. **Consistent** - Uniform spacing, colors, typography

### Component Examples

**Card:**
```html
<div class="card">
    <div class="card-header">Section Title</div>
    <p>Card content...</p>
</div>
```

**Button:**
```html
<button class="btn btn-primary" type="submit">
    Calculate Heating Values
</button>
```

**Status Indicator:**
```html
<div class="status status-pass">
    <strong>Total:</strong> 100.0% ✓ Valid
</div>
```

---

## Technical Reference Page

### Features

- **Comprehensive formulas** - All calculations with mathematical notation
- **Component tables** - HHV, LHV, Cp values for all components
- **Default values** - Operating conditions, environmental parameters
- **Fuel presets** - 9 pre-configured compositions
- **Validation data** - Comparison to Excel VBA and ASME standards
- **Table of contents** - Easy navigation with anchor links

### Navigation

Smooth scrolling to sections:
```javascript
html {
    scroll-behavior: smooth;
}

.section-anchor {
    scroll-margin-top: 100px;  /* Offset for fixed header */
}
```

---

## Browser Compatibility

### Supported Browsers

- ✅ Chrome 90+ (Windows, macOS, Linux)
- ✅ Firefox 88+ (Windows, macOS, Linux)
- ✅ Safari 14+ (macOS, iOS)
- ✅ Edge 90+ (Windows, macOS)

### Required Features

- ES6 JavaScript (async/await, fetch API)
- CSS Grid and Flexbox
- CSS Custom Properties (variables)

---

## Accessibility

### WCAG 2.1 Compliance

- ✅ Semantic HTML structure
- ✅ Proper heading hierarchy
- ✅ Alt text for all images
- ✅ Sufficient color contrast (4.5:1)
- ✅ Keyboard navigation support
- ✅ ARIA labels where appropriate

---

## Performance

### Optimization

- **Minimal dependencies** - No frameworks, vanilla JavaScript
- **Small assets** - < 50 KB total CSS + JS
- **Fast load time** - < 1 second on broadband
- **Efficient API calls** - Only on form submit

### Metrics

- **First Contentful Paint:** < 0.5s
- **Time to Interactive:** < 1.0s
- **Total Page Weight:** < 200 KB

---

## Next Steps

### Adding New Calculators

1. Create HTML form in `calculators/`
2. Add JavaScript handling in `js/`
3. Create Azure Function in `api/`
4. Update navigation in `index.html`
5. Add formulas to `resource.html`

### Planned Calculators

- Air Requirement Calculator
- Products of Combustion Calculator
- Steam Properties Calculator
- Water Properties Calculator
- Flash Steam Calculator

---

## Support

For help with web calculators:
1. Check [HTML Calculators](html-calculators.md) for details
2. Review [Azure Deployment](../azure-deployment/README.md) for hosting
3. Inspect browser console for JavaScript errors
4. Contact GTS Energy Inc.

---

**Back to:** [Main Documentation](../README.md)
