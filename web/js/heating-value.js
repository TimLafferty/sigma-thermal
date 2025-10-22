// Heating Value Calculator JavaScript

// Fuel composition presets (mass %)
const FUEL_PRESETS = {
    'custom': {
        description: 'Custom fuel composition',
        ch4: 0, c2h6: 0, c3h8: 0, c4h10: 0,
        h2: 0, co: 0, h2s: 0, co2: 0, n2: 0
    },
    'pure-methane': {
        description: '100% Methane (CH4) - Reference fuel',
        ch4: 100, c2h6: 0, c3h8: 0, c4h10: 0,
        h2: 0, co: 0, h2s: 0, co2: 0, n2: 0
    },
    'natural-gas-typical': {
        description: 'Typical natural gas composition (pipeline quality)',
        ch4: 85, c2h6: 10, c3h8: 3, c4h10: 1,
        h2: 0, co: 0, h2s: 0, co2: 1, n2: 0
    },
    'natural-gas-high': {
        description: 'High BTU natural gas (high methane content)',
        ch4: 92, c2h6: 6, c3h8: 1.5, c4h10: 0.5,
        h2: 0, co: 0, h2s: 0, co2: 0, n2: 0
    },
    'natural-gas-lean': {
        description: 'Lean natural gas (lower BTU, higher inerts)',
        ch4: 75, c2h6: 12, c3h8: 5, c4h10: 2,
        h2: 0, co: 0, h2s: 0, co2: 5, n2: 1
    },
    'landfill-gas': {
        description: 'Landfill gas (high CO2 content)',
        ch4: 50, c2h6: 0, c3h8: 0, c4h10: 0,
        h2: 0, co: 0, h2s: 0, co2: 45, n2: 5
    },
    'digester-gas': {
        description: 'Anaerobic digester gas',
        ch4: 60, c2h6: 0, c3h8: 0, c4h10: 0,
        h2: 0, co: 0, h2s: 0, co2: 35, n2: 5
    },
    'refinery-gas': {
        description: 'Refinery gas (mixed hydrocarbons)',
        ch4: 40, c2h6: 20, c3h8: 15, c4h10: 10,
        h2: 15, co: 0, h2s: 0, co2: 0, n2: 0
    },
    'coke-oven-gas': {
        description: 'Coke oven gas (high hydrogen content)',
        ch4: 25, c2h6: 2, c3h8: 1, c4h10: 0,
        h2: 55, co: 8, h2s: 0, co2: 3, n2: 6
    },
    'blast-furnace-gas': {
        description: 'Blast furnace gas (very low BTU, high N2)',
        ch4: 0, c2h6: 0, c3h8: 0, c4h10: 0,
        h2: 3, co: 25, h2s: 0, co2: 20, n2: 52
    }
};

// Load preset composition
function loadPreset() {
    const preset = document.getElementById('fuel-preset').value;
    const composition = FUEL_PRESETS[preset];

    if (composition) {
        document.getElementById('ch4').value = composition.ch4;
        document.getElementById('c2h6').value = composition.c2h6;
        document.getElementById('c3h8').value = composition.c3h8;
        document.getElementById('c4h10').value = composition.c4h10;
        document.getElementById('h2').value = composition.h2;
        document.getElementById('co').value = composition.co;
        document.getElementById('h2s').value = composition.h2s;
        document.getElementById('co2').value = composition.co2;
        document.getElementById('n2').value = composition.n2;

        // Update preset info
        const infoDiv = document.getElementById('preset-info');
        infoDiv.textContent = composition.description;

        // Validate total
        validateTotal();
    }
}

// Validate that total composition = 100%
function validateTotal() {
    const ch4 = parseFloat(document.getElementById('ch4').value) || 0;
    const c2h6 = parseFloat(document.getElementById('c2h6').value) || 0;
    const c3h8 = parseFloat(document.getElementById('c3h8').value) || 0;
    const c4h10 = parseFloat(document.getElementById('c4h10').value) || 0;
    const h2 = parseFloat(document.getElementById('h2').value) || 0;
    const co = parseFloat(document.getElementById('co').value) || 0;
    const h2s = parseFloat(document.getElementById('h2s').value) || 0;
    const co2 = parseFloat(document.getElementById('co2').value) || 0;
    const n2 = parseFloat(document.getElementById('n2').value) || 0;

    const total = ch4 + c2h6 + c3h8 + c4h10 + h2 + co + h2s + co2 + n2;

    // Update total display
    document.getElementById('total-value').textContent = total.toFixed(2);

    // Validation status
    const validationDiv = document.getElementById('total-validation');
    const messageSpan = document.getElementById('validation-message');
    const calculateBtn = document.getElementById('calculate-btn');

    if (Math.abs(total - 100) < 0.01) {
        validationDiv.className = 'status status-pass';
        messageSpan.textContent = ' (Valid)';
        calculateBtn.disabled = false;
    } else {
        validationDiv.className = 'status status-fail';
        messageSpan.textContent = ' (Must equal 100%)';
        calculateBtn.disabled = true;
    }

    // Check for high inerts
    const inerts = co2 + n2;
    const inertWarning = document.getElementById('inert-warning');
    if (inerts > 20) {
        inertWarning.style.display = 'block';
    } else {
        inertWarning.style.display = 'none';
    }
}

// Handle form submission
document.getElementById('heating-value-form').addEventListener('submit', async function(e) {
    e.preventDefault();

    const formData = new FormData(this);
    const data = Object.fromEntries(formData);

    try {
        // Make API call to calculate
        const response = await fetch('/api/calculate/heating-value', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        });

        if (!response.ok) {
            throw new Error('Calculation failed');
        }

        const results = await response.json();
        displayResults(results);

    } catch (error) {
        console.error('Error:', error);
        alert('Error calculating heating values. Please check your inputs and try again.');
    }
});

// Display calculation results
function displayResults(results) {
    // Show results section
    document.getElementById('results-section').style.display = 'block';

    // Format numbers with commas
    const format = (num) => num.toLocaleString('en-US', { maximumFractionDigits: 0 });

    // Mass basis
    document.getElementById('hhv-mass').textContent = format(results.hhv_mass);
    document.getElementById('lhv-mass').textContent = format(results.lhv_mass);
    document.getElementById('diff-mass').textContent = format(results.hhv_mass - results.lhv_mass);

    // Volume basis
    document.getElementById('hhv-volume').textContent = format(results.hhv_volume);
    document.getElementById('lhv-volume').textContent = format(results.lhv_volume);
    document.getElementById('diff-volume').textContent = format(results.hhv_volume - results.lhv_volume);

    // Excel comparison (if available)
    if (results.excel_comparison) {
        document.getElementById('excel-comparison').style.display = 'block';
        document.getElementById('python-hhv').textContent = format(results.hhv_mass);
        document.getElementById('excel-hhv').textContent = format(results.excel_comparison.hhv);

        const deviation = Math.abs(results.hhv_mass - results.excel_comparison.hhv) / results.excel_comparison.hhv * 100;
        document.getElementById('deviation').textContent = deviation.toFixed(4) + '%';

        // Status
        const statusDiv = document.getElementById('comparison-status');
        if (deviation <= 1.0) {
            statusDiv.className = 'status status-pass';
            statusDiv.innerHTML = '<strong>PASS</strong> - Deviation: ' + deviation.toFixed(4) + '% (Tolerance: 1.00%)';
        } else if (deviation <= 2.0) {
            statusDiv.className = 'status status-warning';
            statusDiv.innerHTML = '<strong>WARNING</strong> - Deviation: ' + deviation.toFixed(4) + '% (Tolerance: 1.00%)';
        } else {
            statusDiv.className = 'status status-fail';
            statusDiv.innerHTML = '<strong>FAIL</strong> - Deviation: ' + deviation.toFixed(4) + '% (Tolerance: 1.00%)';
        }
    }

    // Scroll to results
    document.getElementById('results-section').scrollIntoView({ behavior: 'smooth' });
}

// Export results as JSON
function exportJSON() {
    const data = {
        fuel_preset: document.getElementById('fuel-preset').value,
        composition: {
            ch4: parseFloat(document.getElementById('ch4').value),
            c2h6: parseFloat(document.getElementById('c2h6').value),
            c3h8: parseFloat(document.getElementById('c3h8').value),
            c4h10: parseFloat(document.getElementById('c4h10').value),
            h2: parseFloat(document.getElementById('h2').value),
            co: parseFloat(document.getElementById('co').value),
            h2s: parseFloat(document.getElementById('h2s').value),
            co2: parseFloat(document.getElementById('co2').value),
            n2: parseFloat(document.getElementById('n2').value)
        },
        results: {
            hhv_mass: document.getElementById('hhv-mass').textContent,
            lhv_mass: document.getElementById('lhv-mass').textContent,
            hhv_volume: document.getElementById('hhv-volume').textContent,
            lhv_volume: document.getElementById('lhv-volume').textContent
        }
    };

    const json = JSON.stringify(data, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);

    const a = document.createElement('a');
    a.href = url;
    a.download = 'heating-value-results.json';
    a.click();

    URL.revokeObjectURL(url);
}

// Initialize on page load
document.addEventListener('DOMContentLoaded', function() {
    loadPreset(); // Load default preset
    validateTotal(); // Initial validation
});
