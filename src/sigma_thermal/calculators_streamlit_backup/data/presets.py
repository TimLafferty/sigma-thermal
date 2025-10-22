"""
Example scenario presets for calculators.

Provides common fuel compositions and operating conditions
for quick calculator setup.
"""

# Fuel composition presets (mass %)
FUEL_PRESETS = {
    "Custom": {
        "description": "Custom fuel composition",
        "composition": {},
        "category": "custom"
    },
    "Pure Methane": {
        "description": "100% Methane (CH4) - Reference fuel",
        "composition": {
            "methane_mass": 100.0,
            "ethane_mass": 0.0,
            "propane_mass": 0.0,
            "butane_mass": 0.0,
            "hydrogen_mass": 0.0,
            "carbon_monoxide_mass": 0.0,
            "hydrogen_sulfide_mass": 0.0,
            "carbon_dioxide_mass": 0.0,
            "nitrogen_mass": 0.0
        },
        "category": "natural_gas"
    },
    "Natural Gas (Typical)": {
        "description": "Typical pipeline natural gas composition",
        "composition": {
            "methane_mass": 85.0,
            "ethane_mass": 10.0,
            "propane_mass": 3.0,
            "butane_mass": 1.0,
            "hydrogen_mass": 0.0,
            "carbon_monoxide_mass": 0.0,
            "hydrogen_sulfide_mass": 0.0,
            "carbon_dioxide_mass": 1.0,
            "nitrogen_mass": 0.0
        },
        "category": "natural_gas"
    },
    "Natural Gas (High BTU)": {
        "description": "High BTU natural gas with higher propane",
        "composition": {
            "methane_mass": 80.0,
            "ethane_mass": 12.0,
            "propane_mass": 6.0,
            "butane_mass": 1.5,
            "hydrogen_mass": 0.0,
            "carbon_monoxide_mass": 0.0,
            "hydrogen_sulfide_mass": 0.0,
            "carbon_dioxide_mass": 0.5,
            "nitrogen_mass": 0.0
        },
        "category": "natural_gas"
    },
    "Natural Gas (Lean)": {
        "description": "Lean natural gas with higher inerts",
        "composition": {
            "methane_mass": 88.0,
            "ethane_mass": 6.0,
            "propane_mass": 2.0,
            "butane_mass": 0.5,
            "hydrogen_mass": 0.0,
            "carbon_monoxide_mass": 0.0,
            "hydrogen_sulfide_mass": 0.0,
            "carbon_dioxide_mass": 2.0,
            "nitrogen_mass": 1.5
        },
        "category": "natural_gas"
    },
    "Landfill Gas": {
        "description": "Typical landfill gas (biogas)",
        "composition": {
            "methane_mass": 55.0,
            "ethane_mass": 0.0,
            "propane_mass": 0.0,
            "butane_mass": 0.0,
            "hydrogen_mass": 0.0,
            "carbon_monoxide_mass": 0.0,
            "hydrogen_sulfide_mass": 0.5,
            "carbon_dioxide_mass": 43.0,
            "nitrogen_mass": 1.5
        },
        "category": "renewable"
    },
    "Digester Gas": {
        "description": "Anaerobic digester gas",
        "composition": {
            "methane_mass": 60.0,
            "ethane_mass": 0.0,
            "propane_mass": 0.0,
            "butane_mass": 0.0,
            "hydrogen_mass": 0.0,
            "carbon_monoxide_mass": 0.0,
            "hydrogen_sulfide_mass": 1.0,
            "carbon_dioxide_mass": 38.0,
            "nitrogen_mass": 1.0
        },
        "category": "renewable"
    },
    "Refinery Gas": {
        "description": "Typical refinery fuel gas",
        "composition": {
            "methane_mass": 35.0,
            "ethane_mass": 20.0,
            "propane_mass": 15.0,
            "butane_mass": 10.0,
            "hydrogen_mass": 18.0,
            "carbon_monoxide_mass": 0.0,
            "hydrogen_sulfide_mass": 0.5,
            "carbon_dioxide_mass": 0.5,
            "nitrogen_mass": 1.0
        },
        "category": "industrial"
    },
    "Coke Oven Gas": {
        "description": "Coke oven gas (high hydrogen)",
        "composition": {
            "methane_mass": 25.0,
            "ethane_mass": 2.0,
            "propane_mass": 0.0,
            "butane_mass": 0.0,
            "hydrogen_mass": 60.0,
            "carbon_monoxide_mass": 8.0,
            "hydrogen_sulfide_mass": 0.5,
            "carbon_dioxide_mass": 2.0,
            "nitrogen_mass": 2.5
        },
        "category": "industrial"
    },
    "Blast Furnace Gas": {
        "description": "Blast furnace gas (low BTU)",
        "composition": {
            "methane_mass": 0.0,
            "ethane_mass": 0.0,
            "propane_mass": 0.0,
            "butane_mass": 0.0,
            "hydrogen_mass": 2.0,
            "carbon_monoxide_mass": 25.0,
            "hydrogen_sulfide_mass": 0.0,
            "carbon_dioxide_mass": 20.0,
            "nitrogen_mass": 53.0
        },
        "category": "industrial"
    }
}

# Operating condition presets
OPERATING_CONDITIONS = {
    "Standard": {
        "description": "Standard conditions - NIST",
        "ambient_temp": 77.0,  # °F
        "humidity": 0.013,  # lb H2O / lb dry air (60% RH at 77°F)
        "excess_air": 10.0,  # %
    },
    "Boiler (Low Excess Air)": {
        "description": "Well-tuned boiler operation",
        "ambient_temp": 77.0,
        "humidity": 0.013,
        "excess_air": 10.0,
    },
    "Boiler (Moderate Excess Air)": {
        "description": "Typical boiler operation",
        "ambient_temp": 77.0,
        "humidity": 0.013,
        "excess_air": 15.0,
    },
    "Furnace": {
        "description": "Industrial furnace",
        "ambient_temp": 77.0,
        "humidity": 0.013,
        "excess_air": 15.0,
    },
    "Heater": {
        "description": "Process heater",
        "ambient_temp": 77.0,
        "humidity": 0.013,
        "excess_air": 20.0,
    },
    "Incinerator": {
        "description": "Waste incinerator (high excess air)",
        "ambient_temp": 77.0,
        "humidity": 0.013,
        "excess_air": 50.0,
    },
    "Cold Weather": {
        "description": "Winter operation",
        "ambient_temp": 32.0,
        "humidity": 0.001,
        "excess_air": 15.0,
    },
    "Hot & Humid": {
        "description": "Summer operation",
        "ambient_temp": 95.0,
        "humidity": 0.020,
        "excess_air": 15.0,
    }
}

# Stack temperature presets (°F)
STACK_TEMPERATURE_PRESETS = {
    "Condensing Boiler": 150.0,
    "High Efficiency Boiler": 250.0,
    "Standard Boiler": 350.0,
    "Process Heater": 500.0,
    "Furnace (Low)": 800.0,
    "Furnace (Moderate)": 1200.0,
    "Furnace (High)": 1500.0,
    "Incinerator": 1800.0
}

# Steam pressure presets (psia)
STEAM_PRESSURE_PRESETS = {
    "Vacuum (Evaporator)": 2.0,
    "Low Pressure (HVAC)": 15.0,
    "Atmospheric": 14.7,
    "Low Steam (50 psig)": 64.7,
    "Medium Steam (100 psig)": 114.7,
    "High Steam (150 psig)": 164.7,
    "Process Steam (200 psig)": 214.7,
    "High Pressure (400 psig)": 414.7,
    "Utility Steam (600 psig)": 614.7
}


def get_fuel_preset_names(category: str = None):
    """
    Get list of fuel preset names, optionally filtered by category.

    Parameters
    ----------
    category : str, optional
        Category filter ('natural_gas', 'renewable', 'industrial', 'custom')

    Returns
    -------
    list
        List of preset names
    """
    if category is None:
        return list(FUEL_PRESETS.keys())
    else:
        return [name for name, preset in FUEL_PRESETS.items()
                if preset['category'] == category]


def get_fuel_composition(preset_name: str):
    """
    Get fuel composition for a preset.

    Parameters
    ----------
    preset_name : str
        Name of preset

    Returns
    -------
    dict
        Composition dictionary
    """
    if preset_name not in FUEL_PRESETS:
        return {}

    return FUEL_PRESETS[preset_name]['composition'].copy()


def get_operating_conditions(preset_name: str):
    """
    Get operating conditions for a preset.

    Parameters
    ----------
    preset_name : str
        Name of preset

    Returns
    -------
    dict
        Operating conditions
    """
    if preset_name not in OPERATING_CONDITIONS:
        return OPERATING_CONDITIONS['Standard'].copy()

    return OPERATING_CONDITIONS[preset_name].copy()
