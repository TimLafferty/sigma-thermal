#!/bin/bash
# Launch Sigma Thermal Calculators Web App

echo "🔥 Starting Sigma Thermal Engineering Calculators..."
echo ""
echo "The app will open in your browser at http://localhost:8501"
echo "Press Ctrl+C to stop the server"
echo ""

cd "$(dirname "$0")"

# Use venv if it exists, otherwise use system python
if [ -d ".venv" ]; then
    echo "Using virtual environment..."
    .venv/bin/streamlit run src/sigma_thermal/calculators/app.py
else
    streamlit run src/sigma_thermal/calculators/app.py
fi
