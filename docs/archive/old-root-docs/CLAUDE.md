# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Overview

This repository contains Excel-based thermal heating system design and calculation tools for GTS Energy Inc. The primary files are VBA-enabled Excel workbooks (.xlsm) and add-ins (.xlam) that perform engineering calculations for industrial heating systems.

## Repository Structure

```
sources/
├── HC2-Calculators.xlsm    # Main calculation workbook (6.4MB)
└── Engineering-Functions.xlam    # VBA function library (526KB)
```

## Main Components

### HC2-Calculators.xlsm

The primary workbook containing 27 worksheets organized into functional areas:

**Input Sheets:**
- New Primary Inputs - Main user input interface
- Lookups - Reference tables and data
- Internal Datasheet - Internal calculations
- Customer Datasheet - Customer-facing output (printable)
- Hoja de Datos - Spanish language datasheet

**Calculation Sheets:**
- Heater Calcs - Core thermal calculations
- Qtion - Heat transfer calculations
- W-Beam Data 2.0 - W-beam specific calculations
- Expansion Calculation - Thermal expansion analysis
- Secondary Loop Balance - Secondary loop hydraulics

**Equipment Sizing:**
- Heater Equip Area
- Air and Exhaust Equip Area
- Burner and Controls Equip Area
- System Equip Area
- Fuel Train Equip Area
- Burner House Sizing

**Pricing/Budget:**
- New Budget - Cost estimation and pricing
- Item Lookup - Equipment item database
- Item Table - Component pricing tables
- _002___INDIRECT_COSTS - Indirect cost calculations

**Output/Documentation:**
- Heater Table - Equipment summary table (printable)
- Heater Nameplate - Nameplate specifications (printable)
- Control Panel STs - Control panel specifications
- Cv Tables - Control valve sizing tables

**P&ID Diagrams:**
- Heater P&ID
- Burner P&ID
- Pump P&ID
- Expansion P&ID

### Engineering-Functions.xlam

VBA add-in providing custom engineering functions used throughout the HC2 workbook. These functions likely include thermal property calculations, heat transfer correlations, and other engineering utilities.

## Key Named Ranges

The workbook uses 293+ named ranges for inputs, calculations, and outputs. Key functional areas include:

**System Configuration:**
- Heater_Type, Heater_Models, Configuration_Style
- Burner_Make, Burner_Models, Burner_Duty_Cycle
- Pump_Make, Pump_Configuration
- Installation_Environment, Area_Classification

**Thermal/Fluid Properties:**
- Fluid_Lookup, FluidType, ThermalFluid
- FluidFlow1-4, FluidTempIn1-4, FluidTempOut1-4
- OilTempIn, OilTempOut1-4, GasTempOut1-4
- Fuel gas composition (MethaneVol, EthaneVol, PropaneVol, etc.)

**Equipment Specifications:**
- FT_Assembly, FT_BMS_Code (Fuel train)
- Motor_Requirements, Motor_Size, Motor_Voltage
- Stack_Height, Stack_Diameter
- Standard_Tanks, Tank_Orientation, Tank_Tower_Height

**Engineering Standards:**
- Code_Stamp, Piping_Code, Electrical_Code
- Panel_Certification, NEMA
- Incoterms_2000, Incoterms_Location

**Calculations:**
- EfficiencyAssumed, EfficiencyCalc
- AirFlow, FuelFlowMass, FuelFlowVol
- FGRPercent (Flue Gas Recirculation)
- Capacity, Heater duty calculations (_Qg1-4, _QHE1-4, _Qo1-4)

## Working with Excel VBA Files

### Viewing VBA Code

Excel VBA files (.xlsm, .xlam) are binary formats containing VBA macros. To extract and view VBA code:

```bash
# Extract the entire workbook structure
unzip -d extracted/ sources/HC2-Calculators.xlsm

# The VBA code is stored in xl/vbaProject.bin (binary format)
# Use Excel or a VBA extraction tool to view the actual code
```

### File Format Notes

- `.xlsm` - Excel Macro-Enabled Workbook (Office Open XML with macros)
- `.xlam` - Excel Add-In (macro-enabled)
- Both formats are ZIP archives containing XML files and binary VBA projects
- VBA code is stored in `xl/vbaProject.bin` in a binary format that requires specialized tools to parse

### Modifying Excel Files

To modify these files programmatically, consider:

1. **Python + openpyxl**: For worksheet data and structure (cannot handle VBA)
2. **Python + xlwings**: For VBA-enabled workbooks (requires Excel installation)
3. **Python + oletools/olevba**: For VBA code extraction and analysis
4. **Excel application**: Direct editing via Excel (most reliable for VBA)

### Important Considerations

- The workbook contains extensive VBA code in `xl/vbaProject.bin`
- Making changes requires understanding the VBA macro dependencies
- Named ranges are critical to workbook function - do not modify without understanding dependencies
- Printable sheets have defined Print_Area ranges: Customer Datasheet, Heater Nameplate, Heater Table, Secondary Loop Balance

## Engineering Context

This is a thermal heating system design tool that:
- Sizes industrial heaters and burners
- Performs thermal calculations for heat exchangers
- Designs fuel trains and control systems
- Generates customer datasheets and P&ID diagrams
- Provides cost estimates and equipment lists
- Supports multiple unit systems and languages (English/Spanish)

The tool appears to follow industry standards including various piping codes, electrical codes, and equipment certifications (ASME, NEMA, etc.).
