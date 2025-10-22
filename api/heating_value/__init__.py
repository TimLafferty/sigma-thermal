"""
Azure Function: Heating Value Calculator API
"""

import azure.functions as func
import json
import logging
import sys
from pathlib import Path

# Add parent directory to path to import sigma_thermal
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from sigma_thermal.combustion import (
    GasComposition,
    hhv_mass_gas,
    lhv_mass_gas,
    hhv_volume_gas,
    lhv_volume_gas
)


def main(req: func.HttpRequest) -> func.HttpResponse:
    """
    Calculate heating values for gaseous fuels.

    POST /api/heating-value
    Body: {
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
    """
    logging.info('Heating value calculation requested')

    # Enable CORS
    headers = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type',
        'Content-Type': 'application/json'
    }

    # Handle OPTIONS preflight request
    if req.method == 'OPTIONS':
        return func.HttpResponse(
            status_code=204,
            headers=headers
        )

    try:
        # Parse request body
        req_body = req.get_json()

        # Validate inputs
        required_fields = ['ch4', 'c2h6', 'c3h8', 'c4h10', 'h2', 'co', 'h2s', 'co2', 'n2']
        for field in required_fields:
            if field not in req_body:
                return func.HttpResponse(
                    json.dumps({'error': f'Missing required field: {field}'}),
                    status_code=400,
                    headers=headers
                )

        # Create fuel composition
        fuel = GasComposition(
            methane_mass=float(req_body['ch4']),
            ethane_mass=float(req_body['c2h6']),
            propane_mass=float(req_body['c3h8']),
            butane_mass=float(req_body['c4h10']),
            hydrogen_mass=float(req_body['h2']),
            carbon_monoxide_mass=float(req_body['co']),
            hydrogen_sulfide_mass=float(req_body['h2s']),
            carbon_dioxide_mass=float(req_body['co2']),
            nitrogen_mass=float(req_body['n2'])
        )

        # Calculate heating values
        results = {
            'hhv_mass': round(hhv_mass_gas(fuel), 2),
            'lhv_mass': round(lhv_mass_gas(fuel), 2),
            'hhv_volume': round(hhv_volume_gas(fuel), 2),
            'lhv_volume': round(lhv_volume_gas(fuel), 2)
        }

        # Add Excel comparison for Pure Methane
        if req_body['ch4'] >= 99.9 and sum([req_body[f] for f in required_fields if f != 'ch4']) < 0.1:
            results['excel_comparison'] = {
                'hhv': 23875.0,
                'deviation': abs(results['hhv_mass'] - 23875.0) / 23875.0
            }

        logging.info('Heating value calculation successful')

        return func.HttpResponse(
            json.dumps(results),
            status_code=200,
            headers=headers
        )

    except ValueError as e:
        logging.error(f'Invalid input: {str(e)}')
        return func.HttpResponse(
            json.dumps({'error': f'Invalid input: {str(e)}'}),
            status_code=400,
            headers=headers
        )

    except Exception as e:
        logging.error(f'Error calculating heating values: {str(e)}')
        return func.HttpResponse(
            json.dumps({'error': f'Calculation error: {str(e)}'}),
            status_code=500,
            headers=headers
        )
