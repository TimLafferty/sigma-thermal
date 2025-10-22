"""
Sigma Thermal - Industrial Heater Design and Calculation Library

This package provides comprehensive tools for thermal heating system design,
sizing, and analysis. It replaces the Excel VBA-based HC2 calculator system
with a modern, maintainable Python implementation.

Modules
-------
combustion : Combustion calculations, heating values, emissions
fluids : Fluid property calculations
heat_transfer : Radiant and convective heat transfer
engineering : Engineering utilities and equipment sizing
data : Lookup tables and reference data

Example
-------
>>> from sigma_thermal.combustion import GasComposition, hhv_mass_gas
>>> fuel = GasComposition(methane_mass=100)
>>> hhv = hhv_mass_gas(fuel)
>>> print(f"HHV: {hhv:.2f} BTU/lb")
"""

__version__ = "0.1.0"
__author__ = "GTS Energy Inc"

# Core imports for convenience
from . import (
    combustion,
    fluids,
    heat_transfer,
    engineering,
)

__all__ = [
    "combustion",
    "fluids",
    "heat_transfer",
    "engineering",
]
