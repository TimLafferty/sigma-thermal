# VBA Function Inventory

**Total Functions**: 576

## Functions by Module

### CombustionFunctions.bas (30 functions)

- **FlueGasEnthalpy**(H2O: Single, CO2: Single, N2: Single, O2: Single, GasTemp: Single, AmbientTemp: Single) → Single
- **NOxConv**(Source: Single, SourceUnit: String, TargetUnit: String) → Variant
- **COConv**(Source: Single, SourceUnit: String, TargetUnit: String) → Variant
- **HHVMass**(FuelType: String, AirMass: Single, ArgonMass: Single, MethaneMass: Single, EthaneMass: Single, PropaneMass: Single, ButaneMass: Single, PentaneMass: Single, HexaneMass: Single, CO2Mass: Single, COMass: Single, CMass: Single, N2Mass: Single, H2Mass: Single, O2Mass: Single, H2SMass: Single, H2OMass: Single) → Variant
- **Efficiency**(MFlueGas: Single, MAirFlow: Single, EnthalpyHigh: Single, EnthalpyLow: Single, MFuel: Single, LHV: Single) → Variant
- **FlameTemp**(FuelFlowMass: Single, LHVMass: Single, Humidity: Single, ExcessAirMass: Single, Tambient: Single, POC_CO2: Single, POC_H2O: Single, POC_N2: Single, POC_O2: Single, HeatLoss: Single) → Single
- **RecircTemp**(FlameTemp: Single, StackTemp: Single, AmbientTemp: Single, RecircRate: Single, CO2Flow: Single, H2OFlow: Single, N2Flow: Single, O2Flow: Single, MoistAir: Single) → Single
- **EnthalpyCO2**(GasTemp: Single, AmbientTemp: Single) → Variant
- **EnthalpyH2O**(GasTemp: Single, AmbientTemp: Single) → Variant
- **EnthalpyN2**(GasTemp: Single, AmbientTemp: Single) → Variant
- **EnthalpyO2**(GasTemp: Single, AmbientTemp: Single) → Variant
- **AirFuelRatioVol**(FuelType: String, AirMass: Single, FuelFlowVol: Single) → Variant
- **POC_H2OMass**(FuelType: String, FuelFlowMass: Single, Humidity: Single, AirFlowMass: Single, AirMass: Single, ArgonMass: Single, MethaneMass: Single, EthaneMass: Single, PropaneMass: Single, ButaneMass: Single, PentaneMass: Single, HexaneMass: Single, CO2Mass: Single, COMass: Single, CMass: Single, N2Mass: Single, H2Mass: Single, O2Mass: Single, H2SMass: Single, H2OMass: Single) → Variant
- **POC_CO2Vol**(FuelType: String, AirVol: Single, AmmoniaVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, IButeneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, SVol: Single, SO2Vol: Single, H2OVol: Single) → Variant
- **POC_H2OVol**(FuelType: String, AirVol: Single, AmmoniaVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, IButeneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, SVol: Single, SO2Vol: Single, H2OVol: Single) → Variant
- **POC_N2Vol**(FuelType: String, AirVol: Single, AmmoniaVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, IButeneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, SVol: Single, SO2Vol: Single, H2OVol: Single) → Variant
- **POC_SO2Vol**(FuelType: String, AirVol: Single, AmmoniaVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, IButeneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, SVol: Single, SO2Vol: Single, H2OVol: Single) → Variant
- **RequiredO2ForCombustion**(FuelType: String, AirVol: Single, AmmoniaVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, IButeneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, SVol: Single, SO2Vol: Single, H2OVol: Single) → Variant
- **POC_CO2Mass**(FuelType: String, FuelFlowMass: Single, AirFlowMass: Single, AirMass: Single, ArgonMass: Single, MethaneMass: Single, EthaneMass: Single, PropaneMass: Single, ButaneMass: Single, PentaneMass: Single, HexaneMass: Single, CO2Mass: Single, COMass: Single, CMass: Single, N2Mass: Single, H2Mass: Single, O2Mass: Single, H2SMass: Single, H2OMass: Single) → Variant
- **POC_N2Mass**(FuelType: String, FuelFlowMass: Single, ExcessAirMass: Single, AirFlowMass: Single, AirMass: Single, ArgonMass: Single, MethaneMass: Single, EthaneMass: Single, PropaneMass: Single, ButaneMass: Single, PentaneMass: Single, HexaneMass: Single, CO2Mass: Single, COMass: Single, CMass: Single, N2Mass: Single, H2Mass: Single, O2Mass: Single, H2SMass: Single, H2OMass: Single) → Variant
- **POC_O2Mass**(FuelFlowMass: Single, ExcessAirMass: Single, AirFlowMass: Single, O2Mass: Single) → Variant
- **AirMass**(GasFlowMass: Single, AirFuelRatio: Single) → Variant
- **LHVMass**(FuelType: String, AirMass: Single, ArgonMass: Single, MethaneMass: Single, EthaneMass: Single, PropaneMass: Single, ButaneMass: Single, PentaneMass: Single, HexaneMass: Single, CO2Mass: Single, COMass: Single, CMass: Single, N2Mass: Single, H2Mass: Single, O2Mass: Single, H2SMass: Single, H2OMass: Single) → Variant
- **AirFuelRatio**(FuelType: String, AirMass: Single, ArgonMass: Single, MethaneMass: Single, EthaneMass: Single, PropaneMass: Single, ButaneMass: Single, PentaneMass: Single, HexaneMass: Single, CO2Mass: Single, COMass: Single, CMass: Single, N2Mass: Single, H2Mass: Single, O2Mass: Single, H2SMass: Single, H2OMass: Single) → Variant
- **FuelFlowMass**(FuelType: String, FuelFlowVol: Single, AirVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, H2OVol: Single) → Variant
- **FuelFlowVol**(FiringRate: Single, LHV_Vol: Single) → Variant
- **HHV_Vol**(FuelType: String, AirVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, H2OVol: Single) → Variant
- **LHV_Vol**(FuelType: String, AirVol: Single, ArgonVol: Single, MethaneVol: Single, EthaneVol: Single, PropaneVol: Single, ButaneVol: Single, PentaneVol: Single, HexaneVol: Single, CO2Vol: Single, COVol: Single, CVol: Single, N2Vol: Single, H2Vol: Single, O2Vol: Single, H2SVol: Single, H2OVol: Single) → Variant
- **ComponentWeight**(Component: String, ComponentVol: Single, GasFlowVol: Single, GasFlowMass: Single) → Variant
- **ComponentWeightPercent**(Component: String, ComponentVol: Single, FluidMW: Single) → Variant

### ConvectionFunctions.bas (31 functions)

- **WorkbookIsOpen**(wbname: Variant) → Boolean
- **Ttubeout**(Itube%: Variant, q: Variant, Pflow: Variant, Cp: Variant, FlowConfig%: Variant) → Variant
- **Ttubeouttemp**(Itube%: Variant, NT: Variant, NR: Variant, Nsplits: Variant, FlowConfig%: Variant) → Variant
- **Cells**(Row: Variant, Col: Variant) → Variant
- **FuelTrainCells**(Row: Variant, Col: Variant) → Variant
- **APO**(OD: Variant, Nfins: Variant, fthick: Variant) → Single
- **Ao**(OD: Variant, Nfins: Variant, fthick: Variant, fheight: Variant, ws: Variant, ftype%: Variant) → Single
- **Corr1**(Re: Variant) → Single
- **Corr3**(fheight: Variant, sf: Variant, tubeconfig%: Variant, ftype%: Variant) → Single
- **Corr5**(NR: Variant, Pl: Variant, Pt: Variant, tubeconfig%: Variant) → Single
- **Jfactor**(CF1: Variant, CF3: Variant, CF5: Variant, df: Variant, OD: Variant, Tb: Variant, ts: Variant) → Variant
- **Hc**(Jfactor: Variant, g: Variant, cpg: Variant, kb: Variant, Visc: Variant, OD: Variant, TubeType%: Variant, tubeconfig%: Variant) → Single
- **Efactor**(ho: Variant, lfht: Variant, tf: Variant, kf: Variant, OD: Variant, df: Variant, ws: Variant, ftype%: Variant) → Single
- **Corr2**(Re: Variant) → Single
- **Corr4**(OD: Variant, Pt: Variant, lf_hfin: Variant, sf: Variant, tubeconfig%: Variant, ftype%: Variant) → Single
- **Corr6**(NR: Variant, Pl: Variant, Pt: Variant, tubeconfig%: Variant) → Single
- **ffactor**(CF2: Variant, CF4: Variant, CF6: Variant, df: Variant, OD: Variant, tubeconfig%: Variant) → Single
- **DPshell**(Gn: Variant, An: Variant, Ad: Variant, NR: Variant, dens1: Variant, dens2: Variant, ffl: Variant) → Single
- **MBL**(od_tube: Variant, Pitchl: Variant, Pitcht: Variant) → Single
- **hrad**(Tgas: Variant, Tinside: Variant, ppCO2: Variant, ppH2O: Variant, ab: Variant, Ao: Variant, OD: Variant, Pl: Variant, Pt: Variant) → Single
- **FinTemp**(e: Variant, Twall: Variant, Tb_shell: Variant) → Single
- **Ac**(OD: Variant, Nfins: Variant, fthick: Variant, fheight: Variant) → Single
- **An**(Ad: Variant, Ac: Variant, LTube: Variant, NT: Variant, A_Baffle: Variant) → Single
- **Prop**(Prop1: Variant, Prop2: Variant, t1: Variant, t2: Variant, T: Variant) → Single

### EngineeringFunctions.bas (153 functions)

- **Interpolate**(x1: Single, x: Single, x2: Single, y1: Single, y2: Single) → Variant
- **TubeWallTemp**(Flux: Single, OD: Single, ID: Single, Hi: Single, WallThickness: Single, k: Single, TempOut: Single) → Variant
- **DINFlangeRating**(PClass: String, temp: Single) → Variant
- **Hco_liquid_cross**(diameter: Single, k: Single, Pr: Single, Re: Single) → Variant
- **Hco_flat_plate**(l: Single, k: Single, Pr: Single, Re: Single) → Variant
- **TubeWallTemperature**(Flux: Single, OD: Single, ID: Single, Hi: Single, Tubewall: Single, k: Single, FluidTemp: Single) → Variant
- **FGRDuctSize**(FGRDensity: Single, FGRMassFlow: Single) → Variant
- **VentDischargeLineSize**(Duty: String) → Variant
- **CADamperSize**(CADuctSize: Single) → Variant
- **CADamperCv**(DamperSize: Single) → Variant
- **CADamperTorque**(DamperSize: Single) → Variant
- **ExpTankExpansionLineSize**(Duty: String) → Variant
- **ExpansionTankID**(TankSize: Single) → Variant
- **ExpansionTankStraightLength**(TankSize: Single) → Variant
- **ExpansionTankLevelSwitchHeight**(TankSize: Single) → Variant
- **ExpTankLineID**(VentLineSize: Single) → Variant
- **ExpVentDischargeID**(ExpLineSize: Single) → Variant
- **ExpTankC1**(TankSize: Single) → Variant
- **ExpTankC2**(TankSize: Single) → Variant
- **ExpTankC3**(TankSize: Single) → Variant
- **ExpTankC4**(TankSize: Single) → Variant
- **ExpTankC6**(TankSize: Single) → Variant
- **ExpTankC7**(TankSize: Single) → Variant
- **ExpTankC8**(TankSize: Single) → Variant
- **RecDrainTankSize**(SystemDesignVol: Single) → Variant
- **RecExpTankSize**(SafeExpVolume: Single) → Variant
- **SystemVolume**(InCoilDia: Single, OutCoilDia: Single, InCoilLength: Single, OutCoilLength: Single, numberTubes: Integer, PipingLineSize: Single, PipingLength: Single, UserVolume: Single, PipeSCH: String) → Variant
- **PipingVolume**(PipeSize: Single, length: Single, PipeSCH: String) → Variant
- **ExpansionVolume**(Duty: String) → Variant
- **GasViscosity**(GasComp: Variant, temp: Single, Press: Single) → Variant
- **GasDensity**(GasComp: Variant, temp: Single, Press: Single) → Variant
- **GasSpecificHeat**(GasComp: Variant, temp: Single, Press: Single) → Variant
- **GasThermalConductivity**(GasComp: Variant, temp: Single, Press: Single) → Variant
- **HelixLength**(PipeOD: Single, PitchDiameter: Single, n: Variant) → Variant
- **Concat**(rng: range) → String
- **CvRequiredGas**(Flow: Single, p1: Single, y: Single, M: Single, x: Single, t1: Single, z: Single) → Variant
- **CvRequiredLiquid**(MassFlow: Single, dP: Single, SpecWeight: Single) → Variant
- **ControlValveDP**(q: Single, p1: Single, y: Single, M: Single, Cv: Single, t1: Single, z: Single) → Variant
- **FluidMixture**(CO2: Single, H2O: Single, N2: Single, O2: Single, SO2: Single, Methane: Single, Ethane: Single, Propane: Single, Butane: Single, H2S: Single, H2: Single) → String
- **PartialVolumeHorzCyl**(radius: Single, height: Single, length: Single) → Variant
- **PartialVolumeVertCyl**(radius: Single, height: Single, length: Single) → Variant
- **Area**(diameter: Variant) → Variant
- **LinearInterp**(Group1High: Single, Group1Low: Single, Group2High: Single, Group2Low: Single, Target: Single) → Variant
- **EstFanHp**(CFM: Single, inWC: Single, Efficiency: Single) → Variant
- **EstPumpHp**(GPM: Single, PSI: Single, Efficiency: Single) → Variant
- **Talk**(Txt: String) → Variant
- **TestViscosity**(Gas: String, temp: Single, Press: Single) → Variant
- **TeeWeight**(TeeOD: Single, TeeSchedule: String) → Single
- **LineSize**(Flow: Single) → Variant
- **PipeDia**(Flow: Single, FlowUnits: String, VelocityDesired: Single, Schedule: String) → Variant
- **BalanceFlow**(DPInner: Single, DPOuter: Single, OuterFlow: Single) → Variant
- **LookUpHigh**(LookupValue: Variant, LookupRange: range, ResultRange: range) → Variant
- **LiquidOrificeSize**(PipeID: Single, PressureDrop: Single, FluidDensity: Single, Cd: Single, Flow: Single) → Single
- **LiquidOrificeDP**(PipeID: Single, OrificeID: Single, FluidDensity: Single, Cd: Single, Flow: Single) → Single
- **LiquidOrificeFlow**(PipeID: Single, OrificeID: Single, FluidDensity: Single, Cd: Single, PressureDrop: Single) → Single
- **CoilVolume**(diameter: Single, length: Single, numberTubes: Integer, NumberReturns: Integer, ReturnRadius: Single) → Single
- **hco**(Re: Single, k: Single, d: Single) → Single
- **LossCoefficientBend**(Reynolds: Single, BendRadius: Single, ID: Single, Angle: Single, frictionFactor: Single) → Single
- **LossCoefficientDivergingTee**(FlowBranch: Single, FlowHeader: Single, AreaBranch: Single, AreaHeader: Single, frictionFactor: Single) → Single
- **EquivalentDiameter**(diameter: Single, splits: Integer) → Variant
- **hci**(Di: Variant, Re: Variant, Pr: Variant, ki: Variant, visc_i: Variant, visc_w: Variant) → Single
- **Hci_nucleate_boiling**(Viscosity_liquid: Single, LatentHeat: Single, Density_liquid: Single, Density_vapor: Single, Cp_liquid: Single, Tsurface: Single, Tsat: Single, Pr: Single, SurfaceConstant: Single, SurfaceTension: Single) → Variant
- **Hci_film_boiling**(Re: Single, Density_vapor: Single, Density_liquid: Single, ThermCond_vapor: Single, Viscosity_vapor: Single) → Variant
- **IdealGasDensity**(MolecularWeight: Single, Pressure: Single, Temperature: Single) → Variant
- **LossCoefficientConvergingTee**(FlowBranch: Single, FlowHeader: Single, AreaBranch: Single, AreaHeader: Single, frictionFactor: Single) → Single
- **Overall_U**(hci: Single, hco: Single, TubeResistance: Single, FoulingFactor: Single, OD: Single, ID: Single) → Variant
- **PipeID**(PipeOD: Single, PipeWall: Single) → Variant
- **PipeOD**(PipeNom: Single) → Variant
- **PipeSurfaceArea**(PipeOD: Single, length: Single, Number: Single) → Variant
- **PipeThermalConductivity**(Material: String, Temperature: Single) → Variant
- **PipeWall**(PipeNom: Single, Schedule: String) → Variant
- **LMTD**(TSin: Variant, TSout: Variant, TTin: Variant, TTout: Variant, FlowConfig%: Variant) → Double
- **FlangeRating**(temp: Single, PClass: String, Material: String) → Variant
- **FlangeWeight**(PipeNPS: Single, FlangeType: String, FlangeRating: String) → Variant
- **CapWeight**(CapSize: Single, CapSchedule: String) → Single
- **DPPipe**(ID: Variant, Velocity: Variant, Re: Variant, dens1: Variant, dens2: Variant, Visc: Variant, L_eq: Variant, e: Variant) → Single
- **FuelTrainPipeDP**(PipeNomSize: Single, PipeSCH: String, FuelMW: Single, FuelTemp: Single, FuelPressure: Single, FuelFlowRate: Single, FuelTrainLength: Single) → Variant
- **PipeVelocity**(PipeNomSize: Single, PipeSCH: String, FuelMW: Single, FuelTemp: Single, FuelPressure: Single, FuelFlowRate: Single) → Variant
- **ElbowWeight**(ElbowOD: Single, Schedule: String, radius: String) → Single
- **PipeWeight**(OD: Single, ID: Single, length: Single, Density: Single) → Variant
- **LossCoefficientReturn**(Reynolds: Single, ReturnRadius: Single, ID: Single, Angle: Single, frictionFactor: Single) → Single
- **LossCoefficientMiteredReturn**(Reynolds: Single, ReturnRadius: Single, ID: Single, frictionFactor: Single) → Single
- **ReturnRadius**(PipeNom: Single, PipeOD: Single, ReturnSpacing: String) → Single
- **ReturnWeight**(OD: Single, ID: Single, ReturnRadius: Single, Number: Integer, Density: Single) → Variant
- **Reynolds**(Velocity: Single, diameter: Single, Viscosity: Single, Density: Single) → Variant
- **Prandtl2**(Viscosity: Single, SpecificHeat: Single, ThermalConductivity: Single) → Variant
- **API530RuptureWall**(Material: String, PipeOD: Single, Pressure: Single, Temperature: Single, CorrosionAllowance: Single) → Variant
- **API530ElasticWall**(Material: String, PipeID: Single, Pressure: Single, Temperature: Single, CorrosionAllowance: Single) → Variant
- **frictionFactor**(Re: Single, diameter: Single) → Variant
- **ReturnSurfaceArea**(PipeOD: Single, ReturnRadius: Single, Number: Integer) → Variant
- **SecVIIIMinWall**(Material: String, PipeID: Single, Pressure: Single, Temperature: Single, CorrosionAllowance: Single, Efficiency: Single) → Variant
- **SecVIIIMaxPressure**(Material: String, PipeID: Single, PipeWall: Single, Temperature: Single, CorrosionAllowance: Single, Efficiency: Single) → Variant
- **SecVIIIAllowableStress**(Material: String, Temperature: Single) → Single
- **TubeResistance**(OD: Single, ID: Single, TubeThermCond: Single) → Variant
- **LossCoefficientReducer**(alpha: Single, initialDiameter: Single, finalDiameter: Single, frictionFactor: Variant) → Single
- **oxygenConcentration**(measuredConcentration: Single, O2refValue: Single, measuredO2percent: Single) → Single
- **nitrogenConcentration**(measuredO2ppm: Single, measuredO2percent: Single) → Single
- **valveSize**(Cv: Single, range: String) → String
- **BallRotation**(Cv: Variant, valveSize: String) → Single
- **Reinforcement**(P: Double, HeaderNom: Single, BranchNom: Single, BranchSCH: String, CA: Double, Material: String, temp: Single) → Variant
- **HC2DesignDuty**(HC2Model: String) → Variant
- **InnerCoilDia**(HC2Model: String) → Variant
- **InnerCoilSplits**(HC2Model: String) → Variant
- **InnerCoilPitch**(HC2Model: String) → Variant
- **InnerCoilTurns**(HC2Model: String) → Variant
- **OuterCoilDia**(HC2Model: String) → Variant
- **OuterCoilSplits**(HC2Model: String) → Variant
- **OuterCoilPitch**(HC2Model: String) → Variant
- **OuterCoilTurns**(HC2Model: String) → Variant
- **HC2DesignFlow**(HC2Model: String) → Variant
- **ShellID**(HC2Model: String) → Variant
- **ShellThickness**(HC2Model: String) → Variant
- **ShellLength**(HC2Model: String) → Variant
- **ShellBottomPlateThickness**(HC2Model: String) → Variant
- **ShellTopPlateThickness**(HC2Model: String) → Variant
- **ShellSupportType**(HC2Model: String) → Variant
- **ShellSupportSize**(HC2Model: String) → Variant
- **ShellSupportFlatbarThickness**(HC2Model: String) → Variant
- **ShellSupportNumber**(HC2Model: String) → Variant
- **ShellSupportLength**(HC2Model: String) → Variant
- **ShellSupportWeldFactor**(HC2Model: String) → Variant
- **ShellBottomPlateFlatbarThickness**(HC2Model: String) → Variant
- **ShellBottomPlateFlatbarHeight**(HC2Model: String) → Variant
- **InnerCoilInletSleeve**(HC2Model: String) → Variant
- **InnerCoilOutletSleeve**(HC2Model: String) → Variant
- **OuterCoilInletSleeve**(HC2Model: String) → Variant
- **OuterCoilOutletSleeve**(HC2Model: String) → Variant
- **CoilType**(HC2Model: String) → Variant
- **InnerWrapperRingBolts**(HC2Model: String) → Variant
- **InnerLiftingStrapWidth**(HC2Model: String) → Variant
- **InnerLiftingStrapThickness**(HC2Model: String) → Variant
- **CoilSpacerType**(HC2Model: String) → Variant
- **CoilSpacerSize**(HC2Model: String) → Variant
- **CoilSpacerNumber**(HC2Model: String) → Variant
- **CoilSpacerLength**(HC2Model: String) → Variant
- **CoilSpacerFlatbarThickness**(HC2Model: String) → Variant
- **CoilSpacerFlatbarWidth**(HC2Model: String) → Variant
- **OuterLiftingStrapWidth**(HC2Model: String) → Variant
- **InnerWrapperLength**(HC2Model: String) → Variant
- **InnerWrapperThickness**(HC2Model: String) → Variant
- **OuterWrapperLength**(HC2Model: String) → Variant
- **OuterWrapperThickness**(HC2Model: String) → Variant
- **CoilLiftingLugs**(HC2Model: String) → Variant
- **CoilLiftingLugLength**(HC2Model: String) → Variant
- **CoilLiftingLugThickness**(HC2Model: String) → Variant
- **CrossPlateEndHeight**(HC2Model: String) → Variant
- **CrossPlateThickness**(HC2Model: String) → Variant
- **CrossPlateEndLength**(HC2Model: String) → Variant
- **By**(vaporFrac: Variant, W: Variant, liqDensity: Variant, vapDensity: Variant, tubeID: Variant, splits: Variant) → Variant
- **Bx**(vaporFrac: Variant, W: Variant, liqDensity: Variant, vapDensity: Variant, liqViscosity: Variant, vapViscosity: Variant, tension: Variant) → Variant
- **Regime**(By: Variant, Bx: Variant) → Variant
- **TwoPhasedP**(vaporFrac: Variant, W: Variant, liqDensity: Single, vapDensity: Single, liqViscosity: Single, vapViscosity: Single, tubeID: Single, splits: Variant, flowType: Variant, Leq: Variant) → Variant
- **frictionFactorCheng**(Re: Single, diameter: Single) → Variant

### FluidFunctions.bas (28 functions)

- **DowAVaporPressure**(Temperature: Single) → Variant
- **DowAVaporViscosity**(Temperature: Single) → Variant
- **DowALiquidEnthalpy**(Temperature: Single) → Variant
- **DowAVaporDensity**(Temperature: Single) → Variant
- **DowAVaporEnthalpy**(Temperature: Single) → Variant
- **DowJVaporPressure**(Temperature: Single) → Variant
- **DowJLiquidEnthalpy**(Temperature: Single) → Variant
- **DowJVaporDensity**(Temperature: Single) → Variant
- **DowJVaporEnthalpy**(Temperature: Single) → Variant
- **DowJVaporViscosity**(Temperature: Single) → Variant
- **FluidDensity**(Fluid: String, Temperature: Single) → Variant
- **FluidSpecificHeat**(Fluid: String, Temperature: Single) → Variant
- **FluidViscosity**(Fluid: String, Temperature: Single) → Variant
- **FluidThermalConductivity**(Fluid: String, Temperature: Single) → Variant
- **LiquidAmineCp**(Temperature: Single, Pressure: Single) → Variant
- **LiquidAmineViscosity**(Temperature: Single, Pressure: Single) → Variant
- **LiquidAmineDensity**(Temperature: Single, Pressure: Single) → Variant
- **AmineEnthalpy**(Temperature: Single, Pressure: Single) → Variant
- **LiquidAmineThermalConductivity**(Temperature: Single, Pressure: Single) → Variant
- **LiquidAmineSurfaceTension**(Temperature: Single, Pressure: Single) → Variant
- **LiquidAmineCriticalPressure**(Temperature: Single, Pressure: Single) → Variant
- **LiquidAmineCriticalTemperature**(Temperature: Single, Pressure: Single) → Variant
- **LiquidAmineMoleWeight**(Temperature: Single, Pressure: Single) → Variant
- **VaporAmineDensity**(Temperature: Single, Pressure: Single) → Variant
- **VaporAmineViscosity**(Temperature: Single, Pressure: Single) → Variant
- **VaporAmineCp**(Temperature: Single, Pressure: Single) → Variant
- **VaporAmineThermalConductivity**(Temperature: Single, Pressure: Single) → Variant
- **VaporAmineMoleWeight**(Temperature: Single, Pressure: Single) → Variant

### Module1.bas (2 functions)


### Module11.bas (38 functions)

- **convertedPressure**(SIunit: String, pressureValue: Single, EnglishUnit: String) → Single

### Module2.bas (1 functions)


### Module9.bas (1 functions)


### PricingFunctions.bas (149 functions)

- **FindPrice**(STNumber: String) → Variant
- **DLDPipingPrice**(LineSize: Single) → Variant
- **TankTowerPrice**(LineSize: Single) → Variant
- **TankTowerAssemblyPrice**(LineSize: Single) → Variant
- **TankTowerPaintPrice**(LineSize: Single) → Variant
- **SalesCommission**(Profit: Single) → Variant
- **RepCommission**(EquipmentType: String, SalesPrice: Single) → Variant
- **MainSSOVSize**(FlowRate: Single) → Variant
- **VentValveSize**(MainSSOVSize: Single) → Variant
- **BurnerPicLookup**(Series: String, ControlMethod: String, PackFan: String) → Variant
- **HeaterPicLookup**(HeaterConfig: String, FlowConfig: String) → Variant
- **FTPicLookup**(PilotPressure: Single, RatedPressure: Single) → Variant
- **SSOVPrice**(Size: String, AreaClassification: String, Actuation: String, Connection: String) → Variant
- **SystemBypassValve**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **SystemBypassActuator**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **SystemBypassYoke**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **SystemBypassFlanges**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **SystemBypassBlind**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **SystemBypassGaskets**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **SystemBypassPositioner**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **SystemBypassBracket**(Size: String, AreaClassification: String, MinTemp: Single) → Variant
- **MainFuelTrainDesignPressure**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **MainRegulatorRegistration**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainRegulator**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainSSOVs**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainVentValve**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindUpstreamPressureGauge**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainPressureGauges**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainPressureSwitches**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindUpstreamGaugeValve**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainGaugeValves**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainInletIsoValve**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainOutletIsoValve**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindFuelTrainStrainer**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindMainSnubbers**(FuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, RegulatorPressure: Single, FuelFlowRate: Single) → Variant
- **FindPilotRegulator**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **PilotSSOVSize**(AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotSSOVs**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotSSOVManufacturer**(AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotVent**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotPressureGauge**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotGaugeValve**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotInletIsoValve**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotOutletIsoValve**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindPilotNeedleValve**(PilotFuelPressure: Single, AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **MinPilotPressure**(AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **MaxPilotPressure**(AreaClass: String, Connect: String, MinTemp: Single, ControlVoltage: String, PilotRegulatorPressure: Single, PilotFuelFlowRate: Single, POC: String) → Variant
- **FindStackPrice**(diameter: Single) → Variant
- **BurnerPilotCapacity**(Manufacturer: String, Model: String, Series: String) → Variant
- **FindPumpModel**(Fluid: String, SystemFlowRate: Single, Manufacturer: String, PumpIsoValveType: String) → Variant
- **FindExpansionTank**(TankCapacity: Single, StandardAlt: String) → Variant
- **FindExpTankLevelSwitch**(TankCapacity: Single, StandardAlt: String) → Variant
- **FindLevelGauge**(TankCapacity: Single, StandardAlt: String) → Variant
- **FindLevelGaugeIsoValve**(TankCapacity: Single, StandardAlt: String) → Variant
- **FindExpTankGaugeVentDrainValve**(TankCapacity: Single, StandardAlt: String) → Variant
- **FindPumpInletIsoValve**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindPumpOutletIsoValve**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindPumpStrainer**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindPumpStrainerDrainValve**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindPumpCheckValve**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindExpansionBellows**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindSuctionPressureGauge**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindDischargePressureGauge**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindSiphons**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindPumpSkidDrainValves**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindPumpSkidGaugeValves**(SystemFlowRate: Single, PumpIsoValveType: String, PumpMan: String) → Variant
- **FindHeaterPSV**(HeaterModel: String, AreaClass: String, MinTemp: Single, StackTW: String) → Variant
- **FindCoilTCs**(HeaterModel: String, AreaClass: String, MinTemp: Single, StackTW: String) → Variant
- **FindCoilTWs**(HeaterModel: String, AreaClass: String, MinTemp: Single, StackTW: String) → Variant
- **FindFlowSwitchAssembly**(HeaterModel: String, AreaClass: String, MinTemp: Single, StackTW: String) → Variant
- **FindHCCoil**(HeaterModel: String) → Variant
- **CoilAssemblyPrice**(HeaterModel: String) → Variant
- **HeaterShellPrice**(HeaterModel: String) → Variant
- **HeaterLidPrice**(HeaterModel: String) → Variant
- **HeaterSubAssembly**(HeaterModel: String) → Variant
- **CoilMaterialAIQualityPrice**(HeaterModel: String) → Variant
- **LidMaterialPrice**(HeaterModel: String) → Variant
- **ShellMaterialPrice**(HeaterModel: String) → Variant
- **SkidLaborPrice**(HeaterModel: String, Orientation: String, SkidSaddle: String) → Variant
- **SkidMaterialPrice**(HeaterModel: String, Orientation: String, SkidSaddle: String) → Variant
- **StackBaseHeight**(StackDia: Single) → Variant
- **StackBasePrice**(StackDia: Single) → Variant
- **StackExtraPrice**(StackDia: Single) → Variant
- **PumpMotorRPM**(Manufacturer: String) → Variant
- **PumpMotorHP**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **PumpMotorFrameSize**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **FindBarePump**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **FindPumpMotor**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **FindPumpCoupling**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **FindPumpCouplingGuard**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **FindPumpBasePlate**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **FindPumpAssembly**(Manufacturer: String, SystemFlowRate: Single, Fluid: String) → Variant
- **PumpSkidFabPrice**(LineSize: Single, PumpQuantity: Single) → Variant
- **PumpSkidPipePrice**(LineSize: Single, PumpQuantity: Single) → Variant
- **PumpSkidAssemblyPrice**(LineSize: Single, PumpQuantity: Single) → Variant
- **PumpSkidPaintPrice**(LineSize: Single, PumpQuantity: Single) → Variant
- **SSFPipePrice**(LineSize: Single) → Variant
- **SSFPipeAssemblyPrice**(LineSize: Single) → Variant
- **SSFPipePaintPrice**(LineSize: Single) → Variant
- **PumpHeaderPipePrice**(LineSize: Single, PumpQuantity: Single) → Variant
- **PumpHeaderPaintPrice**(LineSize: Single, PumpQuantity: Single) → Variant
- **BurnerMaxFuelInletPressure**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerAirPressureReq**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerPilotInletPressure**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerTurndown**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerMaxCapacity**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerSTNumber**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerPackagedFan**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerPackagedFuelValve**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerPackagedAirValve**(Manufacturer: String, Model: String, Series: String) → Variant
- **BurnerSeries**(Manufacturer: String, NOx: Single, CO: Single, ControlMethod: String, FiringRate: Single) → Variant
- **BurnerModel**(Series: String, DesFiringRate: Single) → Variant
- **ControlValveZeroPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveTenPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveTwentyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveThirtyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveFortyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveFiftyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveSixtyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveSeventyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveEightyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveNinetyPerCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveFullOpenCv**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveTorque**(ControlValveSize: Single, VBallAngle: Single) → Variant
- **ControlValveSize**(Cv: Single) → Variant
- **FindVBallAngle**(ControlValveSize: Single, Cv: Single) → Variant
- **FindControlValveSize**(SSOVSize: Single) → Variant
- **FindFTOutletIsoValveCv**(valveSize: Single, Connect: String) → Variant
- **FindMainSSOVCv**(SSOVSize: Single) → Variant
- **FindSystemBypassValveSize**(CvRequired: Single) → Variant
- **FindPilotSSOVCv**(SSOVSize: Single, Manufacturer: String) → Variant
- **FindNeedleValveCV**(ALOSize: Single) → Variant
- **ValvePrice**(ValveType: String, Size: String, Material: String, PressureClass: String, Connection: String, Make: String, ServiceType: String) → Variant
- **WeldTime**(Material: String, PressureClass: String, Size: String) → Variant
- **BellowsPrice**(Size: String) → Variant
- **BurnerPrice**(Size: String) → Variant
- **HeaderWeldTime**(Material: String, PressureClass: String, Size: String, radius: String) → Variant
- **WeldPrepTime**(Material: String, PressureClass: String, Size: String, PrepType: String) → Variant
- **PipePrice**(Material: String, PressureClass: String, Size: String) → Variant
- **ReturnPrice**(Material: String, PressureClass: String, Size: String, radius: String) → Variant
- **CapPrice**(Material: String, PressureClass: String, Size: String) → Variant
- **FlangePrice**(Material: String, PressureClass: String, Size: String, Connection: String) → Variant
- **TubeStabTime**(length: String) → Variant
- **ShellPlateWelding**(length: String, diameter: String, thickness: String) → Variant
- **SaddleWelding**(diameter: String) → Variant
- **SkidWelding**(beamWidth: String, beamHeight: String, webThickness: String, supports: String) → Variant
- **SkidPainting**(beamWidth: String, beamHeight: String, webThickness: String, supports: String, skidWidth: String, skidLength: String) → Variant
- **materialHandling**(weight: String) → Variant
- **FindMargin**(ProductType: String, Application: String, CustType: String, Scope: String) → Variant

### RadiantFunctions.bas (3 functions)

- **ExchangeFactor**(ratio: Single, Emissivity: Single) → Variant
- **GasRadCoeff**(TubeWallTemp: Single, AvgGasTemp: Variant) → Variant
- **GasEmissivity**(Pl: Single, GasTemp: Single) → Variant

### RefpropCode.bas (108 functions)

- **Setup**(FluidName: Variant) → Variant
- **Temperature**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Pressure**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Density**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **CompressibilityFactor**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **LiquidDensity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **VaporDensity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **LiquidEnthalpy**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **VaporEnthalpy**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **LiquidEntropy**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **VaporEntropy**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **LiquidCp**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **VaporCp**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Volume**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Energy**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Enthalpy**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Entropy**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **IsochoricHeatCapacity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Cv**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **IsobaricHeatCapacity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Cp**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Csat**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **SpeedOfSound**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Sound**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **LatentHeat**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **HeatOfVaporization**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **HeatOfCombustion**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **GrossHeatingValue**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **NetHeatingValue**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **JouleThomson**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **IsentropicExpansionCoef**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **IsothermalCompressibility**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **VolumeExpansivity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **AdiabaticCompressibility**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **AdiabaticBulkModulus**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **IsothermalExpansionCoef**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **IsothermalBulkModulus**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **SpecificHeatInput**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **SecondVirial**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dPdrho**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **d2Pdrho2**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dPdT**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dPdTsat**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dHdT_D**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dHdT_P**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dHdD_T**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dHdD_P**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dHdP_T**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **dHdP_D**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **drhodT**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Cstar**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Quality**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **QualityMole**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **QualityMass**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **LiquidMoleFraction**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **VaporMoleFraction**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **LiquidMassFraction**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **VaporMassFraction**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **Fugacity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **FugacityCoefficient**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **ChemicalPotential**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **Activity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **ActivityCoefficient**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant, i: Variant) → Variant
- **Viscosity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **ThermalConductivity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **KinematicViscosity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **ThermalDiffusivity**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Prandtl**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **SurfaceTension**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **DielectricConstant**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **MolarMass**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **MoleFraction**(FluidName: Variant, i: Variant) → Variant
- **MassFraction**(FluidName: Variant, i: Variant) → Variant
- **ComponentName**(FluidName: Variant, i: Variant) → Variant
- **LiquidFluidString**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **VaporFluidString**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Mole2Mass**(FluidName: Variant, i: Variant, Prop1: Variant, Prop2: Variant, Prop3: Variant, Prop4: Variant, Prop5: Variant, Prop6: Variant, Prop7: Variant, Prop8: Variant, Prop9: Variant, Prop10: Variant, Prop11: Variant, Prop12: Variant, Prop13: Variant, Prop14: Variant, Prop15: Variant, Prop16: Variant, Prop17: Variant, Prop18: Variant, Prop19: Variant, Prop20: Variant) → Variant
- **Mass2Mole**(FluidName: Variant, i: Variant, Prop1: Variant, Prop2: Variant, Prop3: Variant, Prop4: Variant, Prop5: Variant, Prop6: Variant, Prop7: Variant, Prop8: Variant, Prop9: Variant, Prop10: Variant, Prop11: Variant, Prop12: Variant, Prop13: Variant, Prop14: Variant, Prop15: Variant, Prop16: Variant, Prop17: Variant, Prop18: Variant, Prop19: Variant, Prop20: Variant) → Variant
- **EOSMax**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **EOSMin**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **ErrorCode**(InputCell: Variant) → Variant
- **ErrorString**(InputCell: Variant) → Variant
- **Trim2**(a: Variant) → Variant
- **Viscosity_ETAK0**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Viscosity_ETAK1**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Viscosity_ETAKR**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Viscosity_ETAKB**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **Transport_Omega**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **ThermalConductivity_TCXK0**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **ThermalConductivity_TCXKB**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **ThermalConductivity_TCXKC**(FluidName: Variant, InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **UnitConvert**(InputValue: Variant, UnitType: String, OldUnits: String, NewUnits: String) → Variant
- **ConvertUnits**(InpCode: Variant, Units: Variant, Prop1: Variant, Prop2: Variant) → Variant
- **FluidString**(Nmes: Variant, Comps: Variant, massmole: String) → String
- **WorkBookName**() → Variant
- **WhereIsWorkbook**() → Variant
- **WhereAreREFPROPfunctions**() → Variant
- **SeeFileLinkSources**() → Variant
- **PropertyUnits**(InpCode: Variant, Units: Variant) → Variant
- **SelectedDefaultUnits**() → Variant
- **RefpropXLSVersionNumber**() → Variant
- **RefpropDLLVersionNumber**() → Variant

### Sheet20.cls (2 functions)


### Sheet26.cls (4 functions)


### Sheet9.cls (6 functions)


### ThisWorkbook.cls (2 functions)


### WaterBathFunctions.bas (11 functions)

- **Hco_film_boiling**(Kvapor: Single, VaporDensity: Single, LiquidDensity: Single, h: Single, CpVapor: Single, Tsurf: Single, Tsat: Single, OD: Single, VaporViscosity: Single) → Variant
- **BeamLengthCheck**(length: Variant, L1: Variant, L2: Variant, L3: Variant, L4: Variant, L5: Variant, S1: Variant, S2: Variant, S3: Variant, S4: Variant, S5: Variant) → Single
- **ExponentialDecayParameter**(AverageFluxrate: Single, MaxFluxRate: Single, length: Single) → Variant
- **hco_cond**(numberTubes: Integer, LiquidDensity: Single, VaporDensity: Single, LiquidConductivity: Single, LiquidViscosity: Single, diameter: Single, T_Sat: Single, T_surf: Single, LatentHeat: Single) → Variant
- **Grashoff**(beta: Single, LMTD: Single, PipeOD: Single, Density: Single, Viscosity: Single) → Variant
- **beta**(DensityAvg: Single, DensityHighTemp: Single, DensityLowTemp: Single, TempBathHigh: Single, TempBathLow: Single) → Variant
- **hco_nat_conv**(Ra: Single, k_bath: Single, OD: Single) → Variant
- **Hco_nucleate_boiling**(visc_bath: Single, h: Single, DensAvg: Single, DensVapor: Single, Cp_Bath: Single, Tsurf: Single, Tsat: Single, Pr: Single, Csf: Single) → Variant
- **ShellDia**(coilSize: Single, numberTubes: Integer, FireTubeSize: Single, numberPasses: Integer) → Single
- **shellDia2**(coilSize: Single, numberTubes: Integer, chamberSize: Single, returnSize: Single, numberPasses: Integer) → Single
- **shellDia3**(coilSize: Single, numberTubes: Integer, chamberSize: Single, returnSize: Single, numReturns: Integer) → Single

### WoodFunctions.bas (7 functions)

- **WoodLHV**(c: Single, h: Single, O: Single, n: Single, s: Single, Ash: Single, MC: Single) → Variant
- **WoodHHV**(LHV: Single, h: Single, MC: Single) → Variant
- **WoodAirFuelRatio**(c: Single, h: Single, O: Single, s: Single) → Variant
- **Wood_POC_CO2**(c: Single, MassFuel: Single) → Variant
- **Wood_POC_H2O**(h: Single, MC: Single, MassFuel: Single) → Variant
- **Wood_POC_N2**(c: Single, h: Single, O: Single, ExcessAir: Single, n: Single, MassFuel: Single) → Variant
- **Wood_POC_O2**(c: Single, h: Single, O: Single, ExcessAir: Single, MassFuel: Single) → Variant

