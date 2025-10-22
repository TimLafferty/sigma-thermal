olevba 0.60.2 on Python 3.11.4 - http://decalage.info/python/oletools
===============================================================================
FILE: sources/Engineering-Functions.xlam
Type: OpenXML
WARNING  For now, VBA stomping cannot be detected for files in memory
-------------------------------------------------------------------------------
VBA MACRO ThisWorkbook.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/ThisWorkbook'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub SaveMyXlam()
ThisWorkbook.Save
End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet1.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO WoodFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/WoodFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Function WoodLHV(c As Single, h As Single, O As Single, n As Single, s As Single, Ash As Single, MC As Single)
'lower heating value in BTU/lb
Dim C_Enthalpy, H_Enthalpy, O_Enthalpy, N_Enthalpy, S_Enthalpy, H2O_Enthalpy As Single

C_Enthalpy = 14991.8
H_Enthalpy = 40409
O_Enthalpy = -4652.64
N_Enthalpy = 2705.42
S_Enthalpy = 4506.17
H2O_Enthalpy = -1055.46


WoodLHV = c * C_Enthalpy + h * H_Enthalpy + O * O_Enthalpy + n * N_Enthalpy + s * S_Enthalpy + MC * H2O_Enthalpy

End Function
Function WoodHHV(LHV As Single, h As Single, MC As Single)
'Higher heating value in BTU/lb
Dim H2O_Enthalpy
H2O_Enthalpy = -1055.46

WoodHHV = LHV + (9 * h + MC) * H2O_Enthalpy * -1

End Function
Function WoodAirFuelRatio(c As Single, h As Single, O As Single, s As Single)
'Air/Fuel ratio:  ft3 of air/lb fuel
WoodAirFuelRatio = 151.4 * c + 454 * h + 56.8 * s - 56.8 * O
End Function
Function Wood_POC_CO2(c As Single, MassFuel As Single)
Wood_POC_CO2 = 31.5 * c * MassFuel * 0.116

End Function
Function Wood_POC_H2O(h As Single, MC As Single, MassFuel As Single)
Wood_POC_H2O = (188 * h + 21.04 * MC) * MassFuel * 0.047

End Function
Function Wood_POC_N2(c As Single, h As Single, O As Single, ExcessAir As Single, n As Single, MassFuel As Single)

Wood_POC_N2 = ((119.3 * c + 355.3 * h - 44.77 * O) * (1 + ExcessAir) + 13.53 * n) * MassFuel * 0.074

End Function
Function Wood_POC_O2(c As Single, h As Single, O As Single, ExcessAir As Single, MassFuel As Single)
Wood_POC_O2 = (31.5 * c + 94 * h - 11.84 * O) * ExcessAir * MassFuel * 0.084
End Function
-------------------------------------------------------------------------------
VBA MACRO CombustionFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/CombustionFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Public Function FlueGasEnthalpy(H2O As Single, CO2 As Single, N2 As Single, O2 As Single, GasTemp As Single, AmbientTemp As Single) As Single

If (H2O + CO2 + N2 + O2) >= 0.99 And (H2O + CO2 + N2 + O2) <= 1.01 Then
    Dim H2OEnthalpy, CO2Enthalpy, N2Enthalpy, O2Enthalpy As Single
    
    H2OEnthalpy = EnthalpyH2O(GasTemp, AmbientTemp)
    CO2Enthalpy = EnthalpyCO2(GasTemp, AmbientTemp)
    N2Enthalpy = EnthalpyN2(GasTemp, AmbientTemp)
    O2Enthalpy = EnthalpyO2(GasTemp, AmbientTemp)
    
    FlueGasEnthalpy = H2OEnthalpy * H2O + CO2Enthalpy * CO2 + N2Enthalpy * N2 + O2Enthalpy * O2
Else
    FlueGasEnthalpy = "Error, sum does not equal unity"
    Exit Function
End If

End Function
Function NOxConv(Source As Single, SourceUnit As String, TargetUnit As String)

Dim MW As Single

MW = 46.05

    Select Case SourceUnit
        Case "g/GJ"
            Select Case TargetUnit
                Case "mg/Nm3"
                    NOxConv = "To be added"
                Case "ppm"
                    NOxConv = Source * 1.907
                Case "mg/Nm3"
                    'NOxConv = Source / 5.07 * (17.95)
                    NOxConv = "To be added"
                    Exit Function
            End Select
        Case "lb/MMBTU"
            Select Case TargetUnit
                Case "mg/Nm3"
                    NOxConv = 832 * Source
                    NOxConv = NOxConv / 22.4 * MW
                Case "ppm"
                    NOxConv = 832 * Source
                    Exit Function
            End Select
        Case "mg/Nm3"
            Select Case TargetUnit
                Case "ppm"
                    NOxConv = Source * 22.4 / MW
                Case "lb/MMBTU"
                    NOxConv = Source * 22.4 / MW
                    NOxConv = NOxConv / 832
                Case "g/GJ"
                    'NOxConv = Source * 5.07 / (17.95)
                    NOxConv = "To be added"
                    Exit Function
            End Select
        Case "ppm"
            Select Case TargetUnit
                Case "mg/Nm3"
                    NOxConv = Source / 22.4 * MW
                Case "lb/MMBTU"
                    NOxConv = Source / 832
                Case "g/GJ"
                    NOxConv = Source / 1.907
                    Exit Function
            End Select
    
    End Select
    

End Function
Function COConv(Source As Single, SourceUnit As String, TargetUnit As String)

Dim MW As Single

MW = 28.01

    Select Case SourceUnit
        Case "g/GJ"
            Select Case TargetUnit
                Case "mg/Nm3"
                    COConv = "To be added"
                Case "ppm"
                    COConv = Source * 0.00220462 / 0.947817
                    COConv = 1286 * COConv
        End Select
        Case "lb/MMBTU"
            Select Case TargetUnit
                Case "mg/Nm3"
                    COConv = 1286 * Source
                    COConv = Source / 22.4 * MW
                Case "ppm"
                    COConv = 1286 * Source
                    Exit Function
            End Select
        Case "mg/Nm3"
            Select Case TargetUnit
                Case "ppm"
                    COConv = Source * 22.4 / MW
                Case "lb/MMBTU"
                    COConv = Source * 22.4 / MW
                    COConv = Source / 1286
                    Exit Function
            End Select
        Case "ppm"
            Select Case TargetUnit
                Case "mg/Nm3"
                    COConv = Source / 22.4 * MW
                Case "lb/MMBTU"
                    COConv = Source / 1286
                    Exit Function
            End Select
    
    End Select
    

End Function

Public Function HHVMass(FuelType As String, AirMass As Single, ArgonMass As Single, MethaneMass As Single, EthaneMass As Single, PropaneMass As Single, ButaneMass As Single, PentaneMass As Single, HexaneMass As Single, CO2Mass As Single, COMass As Single, CMass As Single, N2Mass As Single, H2Mass As Single, O2Mass As Single, H2SMass As Single, H2OMass As Single)

Dim HHVMassArray As Variant
Dim CompWeightArray As Variant
Dim i As Integer

Select Case FuelType

    Case "Gas"
        HHVMassArray = Array(0, 0, 23875, 22323, 21669, 21321, 21095, 20966, 0, 4347, 14093, 0, 61095, 0, 7097, 0)
        CompWeightArray = Array(AirMass, ArgonMass, MethaneMass, EthaneMass, PropaneMass, ButaneMass, PentaneMass, HexaneMass, CO2Mass, COMass, CMass, N2Mass, H2Mass, O2Mass, H2SMass, H2OMass)
        
        i = 0
        HHVMass = HHVMassArray(i) * CompWeightArray(i) / 100
        Do While i < 15
        
            HHVMass = HHVMass + HHVMassArray(i + 1) * CompWeightArray(i + 1) / 100
            i = i + 1
        Loop
    
    Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
        CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
        HHVMassArray = Array(9797, 20190, 19423, 18993, 18844, 18909, 18126)
        i = 0
        Do While i < 7
               
                If CompWeightArray(i) = FuelType Then
                    HHVMass = HHVMassArray(i)
                    Exit Do
                End If
        i = i + 1
        Loop
End Select

End Function
Function Efficiency(MFlueGas As Single, MAirFlow As Single, EnthalpyHigh As Single, EnthalpyLow As Single, MFuel As Single, LHV As Single)
Efficiency = ((MFlueGas * (EnthalpyHigh - EnthalpyLow) + MAirFlow * 0.013 * -971) / (MFuel * LHV)) 'Flue Gas Energy Change minus Latent Heat of Water Vapor from Humidity divided by LHV Firing Rate
End Function

Public Function FlameTemp(FuelFlowMass As Single, LHVMass As Single, Humidity As Single, ExcessAirMass As Single, Tambient As Single, POC_CO2 As Single, POC_H2O As Single, POC_N2 As Single, POC_O2 As Single, HeatLoss As Single) As Single

Dim H_CO2, H_H2O, H_N2, H_O2 As Single
Dim E_CO2, E_H2O, E_N2, E_O2 As Single
Dim E_in, E_out As Single
Dim tf, step As Single
Dim neg, i As Integer

FlameTemp = 220 'initialize flame temperature (°F)
E_out = (E_CO2 + E_H2O + E_N2 + E_O2) / 10 ^ 6 'total energy out
E_in = (FuelFlowMass * LHVMass) / 10 ^ 6 'total energy in

step = 10
neg = -1
i = 0

Do While Abs((E_out - E_in) / E_in) > 1E-05
    'enthalpy curves for products of combustion
    'H_CO2 = Enthalpy("CO2", "TP", "E", FlameTemp - Tambient, 14.7)
    'H_H2O = Enthalpy("water", "TP", "E", FlameTemp - Tambient, 14.7)
    'H_O2 = Enthalpy("oxygen", "TP", "E", FlameTemp - Tambient, 14.7)
    'H_N2 = Enthalpy("nitrogen", "TP", "E", FlameTemp - Tambient, 14.7)
    
    H_CO2 = EnthalpyCO2(FlameTemp, Tambient)
    H_H2O = EnthalpyH2O(FlameTemp, Tambient)
    H_O2 = EnthalpyO2(FlameTemp, Tambient)
    H_N2 = EnthalpyN2(FlameTemp, Tambient)
    
    
    'H_CO2 = 0.3066 * (FlameTemp - Tambient) + 27.532
    'H_H2O = 0.6074 * (FlameTemp - Tambient) + 94.553
    'H_O2 = 0.2687 * (FlameTemp - Tambient) + 73.209
    'H_N2 = 0.2907 * (FlameTemp - Tambient) + 86.657
    
    'energy for products of combustion
    E_CO2 = H_CO2 * POC_CO2
    E_H2O = H_H2O * POC_H2O + Humidity * ExcessAirMass * 970
    E_N2 = H_N2 * POC_N2
    E_O2 = H_O2 * POC_O2
       
    E_in = (FuelFlowMass * LHVMass) / 10 ^ 6  'total energy in
    E_out = (E_CO2 + E_H2O + E_N2 + E_O2) / 10 ^ 6 + E_in * HeatLoss  'total energy out
    
    FlameTemp = FlameTemp + step
    If ((E_out - E_in) / E_in) * neg < 0 Then
        step = step / -10
        neg = neg * -1
    End If

i = i + 1
If i > 10000 Then
    Exit Do
End If

Loop

      
End Function
Public Function RecircTemp(FlameTemp As Single, StackTemp As Single, AmbientTemp As Single, RecircRate As Single, CO2Flow As Single, H2OFlow As Single, N2Flow As Single, O2Flow As Single, MoistAir As Single) As Single

Dim H_CO2_Flame, H_H2O_Flame, H_N2_Flame, H_O2_Flame As Single
Dim H_CO2_Recirc, H_H2O_Recirc, H_N2_Recirc, H_O2_Recirc As Single
Dim H_CO2_Stack, H_H2O_Stack, H_N2_Stack, H_O2_Stack As Single

Dim E_CO2_Flame, E_H2O_Flame, E_N2_Flame, E_O2_Flame As Single
Dim E_CO2_Recirc, E_H2O_Recirc, E_N2_Recirc, E_O2_Recirc As Single
Dim E_CO2_Stack, E_H2O_Stack, E_N2_Stack, E_O2_Stack As Single

Dim CO2FlowRecirc, H2OFlowRecirc, N2FlowRecirc, O2FlowRecirc As Single
Dim CO2FlowStack, H2OFlowStack, N2FlowStack, O2FlowStack As Single

Dim EFlame, ERecirc, EStack As Single
Dim step, Check As Double
Dim neg, i As Integer

'Flame temp
H_CO2_Flame = EnthalpyCO2(FlameTemp, AmbientTemp)
H_H2O_Flame = EnthalpyH2O(FlameTemp, AmbientTemp)
H_N2_Flame = EnthalpyN2(FlameTemp, AmbientTemp)
H_O2_Flame = EnthalpyO2(FlameTemp, AmbientTemp)

E_CO2_Flame = H_CO2_Flame * CO2Flow
E_H2O_Flame = H_H2O_Flame * H2OFlow + MoistAir * 970
E_N2_Flame = H_N2_Flame * N2Flow
E_O2_Flame = H_O2_Flame * O2Flow

EFlame = E_CO2_Flame + E_H2O_Flame + E_N2_Flame + E_O2_Flame

'Recirc temp
CO2FlowStack = CO2Flow * RecircRate / (1 - RecircRate)
H2OFlowStack = H2OFlow * RecircRate / (1 - RecircRate)
N2FlowStack = N2Flow * RecircRate / (1 - RecircRate)
O2FlowStack = O2Flow * RecircRate / (1 - RecircRate)


'Stack temp
CO2FlowRecirc = CO2Flow + CO2FlowStack
H2OFlowRecirc = H2OFlow + H2OFlowStack
N2FlowRecirc = N2Flow + N2FlowStack
O2FlowRecirc = O2Flow + O2FlowStack

H_CO2_Stack = EnthalpyCO2(StackTemp, AmbientTemp)
H_H2O_Stack = EnthalpyH2O(StackTemp, AmbientTemp)
H_N2_Stack = EnthalpyN2(StackTemp, AmbientTemp)
H_O2_Stack = EnthalpyO2(StackTemp, AmbientTemp)

E_CO2_Stack = H_CO2_Stack * CO2FlowStack
E_H2O_Stack = H_H2O_Stack * H2OFlowStack
E_N2_Stack = H_N2_Stack * N2FlowStack
E_O2_Stack = H_O2_Stack * O2FlowStack
EStack = E_CO2_Stack + E_H2O_Stack + E_N2_Stack + E_O2_Stack


step = 10
neg = -1
i = 0
RecircTemp = 500
H_CO2_Recirc = EnthalpyCO2(RecircTemp, AmbientTemp)
H_H2O_Recirc = EnthalpyH2O(RecircTemp, AmbientTemp)
H_N2_Recirc = EnthalpyN2(RecircTemp, AmbientTemp)
H_O2_Recirc = EnthalpyO2(RecircTemp, AmbientTemp)

E_CO2_Recirc = H_CO2_Recirc * CO2FlowRecirc
E_H2O_Recirc = H_H2O_Recirc * H2OFlowRecirc
E_N2_Recirc = H_N2_Recirc * N2FlowRecirc
E_O2_Recirc = H_O2_Recirc * O2FlowRecirc
ERecirc = E_CO2_Recirc + E_H2O_Recirc + E_N2_Recirc + E_O2_Recirc
Check = ERecirc - EFlame - EStack
  
  Do While Abs(Check) > 100
              
    
    H_CO2_Recirc = EnthalpyCO2(RecircTemp, AmbientTemp)
    H_H2O_Recirc = EnthalpyH2O(RecircTemp, AmbientTemp)
    H_N2_Recirc = EnthalpyN2(RecircTemp, AmbientTemp)
    H_O2_Recirc = EnthalpyO2(RecircTemp, AmbientTemp)
    
    E_CO2_Recirc = H_CO2_Recirc * CO2FlowRecirc
    E_H2O_Recirc = H_H2O_Recirc * H2OFlowRecirc
    E_N2_Recirc = H_N2_Recirc * N2FlowRecirc
    E_O2_Recirc = H_O2_Recirc * O2FlowRecirc
      
    ERecirc = E_CO2_Recirc + E_H2O_Recirc + E_N2_Recirc + E_O2_Recirc
    
    Check = (ERecirc - EFlame - EStack)
    If (Check) * neg < 0 Then
        step = step / -2
        neg = neg * -1
    End If
    RecircTemp = RecircTemp + step
    i = i + 1
    If i > 10000 Then
        Exit Do
    End If
    
  Loop



End Function


Public Function EnthalpyCO2(GasTemp As Single, AmbientTemp As Single)

Dim EnthalpyHigh, EnthalpyLow As Single

EnthalpyHigh = 1.08941E-05 * GasTemp ^ 2 + 0.262597665 * GasTemp + 176.9479842
EnthalpyLow = 1.08941E-05 * AmbientTemp ^ 2 + 0.262597665 * AmbientTemp + 176.9479842
EnthalpyCO2 = EnthalpyHigh - EnthalpyLow

End Function

Public Function EnthalpyH2O(GasTemp As Single, AmbientTemp As Single)

Dim EnthalpyHigh, EnthalpyLow As Single

EnthalpyHigh = 3.65285E-05 * GasTemp ^ 2 + 0.452215911 * GasTemp + 1049.366151
EnthalpyLow = 3.65285E-05 * AmbientTemp ^ 2 + 0.452215911 * AmbientTemp + 1049.366151
EnthalpyH2O = EnthalpyHigh - EnthalpyLow

'EnthalpyH2O = 0.6074 * (GasTemp - AmbientTemp) + 94.553

End Function
Public Function EnthalpyN2(GasTemp As Single, AmbientTemp As Single)

Dim EnthalpyHigh, EnthalpyLow As Single


EnthalpyHigh = 8.46332E-06 * GasTemp ^ 2 + 0.255630011 * GasTemp + 107.2712456
EnthalpyLow = 8.46332E-06 * AmbientTemp ^ 2 + 0.255630011 * AmbientTemp + 107.2712456
EnthalpyN2 = EnthalpyHigh - EnthalpyLow
'EnthalpyN2 = 0.2907 * (GasTemp - AmbientTemp) + 86.657

End Function

Public Function EnthalpyO2(GasTemp As Single, AmbientTemp As Single)

Dim EnthalpyHigh, EnthalpyLow As Single

EnthalpyHigh = 7.53536E-06 * GasTemp ^ 2 + 0.23706691 * GasTemp + 92.56930357
EnthalpyLow = 7.53536E-06 * AmbientTemp ^ 2 + 0.23706691 * AmbientTemp + 92.56930357
EnthalpyO2 = EnthalpyHigh - EnthalpyLow

'EnthalpyO2 = 0.2687 * (GasTemp - AmbientTemp) + 73.209

End Function
Public Function AirFuelRatioVol(FuelType As String, AirMass As Single, FuelFlowVol As Single)
Dim SVAir As Single

'specific volume of air at standard conditions
SVAir = 13.1579 'ft3/lb
If FuelType = "Gas" Then
    AirFuelRatioVol = AirMass * SVAir / FuelFlowVol

Else
    AirFuelRatioVol = AirMass * SVAir / FuelFlowVol * 7.481
End If

End Function
Public Function POC_H2OMass(FuelType As String, FuelFlowMass As Single, Humidity As Single, AirFlowMass As Single, AirMass As Single, ArgonMass As Single, MethaneMass As Single, EthaneMass As Single, PropaneMass As Single, ButaneMass As Single, PentaneMass As Single, HexaneMass As Single, CO2Mass As Single, COMass As Single, CMass As Single, N2Mass As Single, H2Mass As Single, O2Mass As Single, H2SMass As Single, H2OMass As Single)

Dim POC_H2O As Variant
Dim CompWeightArray As Variant
Dim i As Integer

Select Case FuelType

    Case "Gas"
        POC_H2O = Array(0, 0, 2.246, 1.797, 1.634, 1.55, 1.5, 1.46, 0, 0, 0, 0, 8.937, 0, 0.529, 0)
        CompWeightArray = Array(AirMass, ArgonMass, MethaneMass, EthaneMass, PropaneMass, ButaneMass, PentaneMass, HexaneMass, CO2Mass, COMass, CMass, N2Mass, H2Mass, O2Mass, H2SMass, H2OMass)
        
        i = 0
        POC_H2OMass = FuelFlowMass * POC_H2O(i) * CompWeightArray(i) / 100 + Humidity * AirFlowMass + FuelFlowMass * CompWeightArray(15) / 100
        
        Do While i < 15
            POC_H2OMass = POC_H2OMass + FuelFlowMass * POC_H2O(i + 1) * CompWeightArray(i + 1) / 100
            i = i + 1
        Loop
        
    Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
    POC_H2O = Array(1.13, 1.3, 1.2, 1.12, 1.04, 0.97, 0.84)
    CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
    i = 0
    Do While i < 7
           
            If CompWeightArray(i) = FuelType Then
                POC_H2OMass = FuelFlowMass * POC_H2O(i) + Humidity * AirFlowMass
                Exit Do
            End If
    i = i + 1
    Loop
End Select

End Function

Public Function POC_CO2Vol(FuelType As String, AirVol As Single, AmmoniaVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, IButeneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, SVol As Single, SO2Vol As Single, H2OVol As Single)
Dim POC_CO2 As Single
Dim POC_CO2Array As Variant
Dim CompVolArray As Variant
Dim i As Integer

Select Case FuelType

Case "Gas"
    POC_CO2Array = Array(0, 0, 0, 1, 2, 3, 4, 4, 5, 6, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0)
    CompVolArray = Array(AirVol, AmmoniaVol, ArgonVol, MethaneVol, EthaneVol, PropaneVol, ButaneVol, IButeneVol, PentaneVol, HexaneVol, CO2Vol, COVol, CVol, N2Vol, H2Vol, O2Vol, H2SVol, SVol, SO2Vol, H2OVol)
    
    i = 0
    POC_CO2Vol = POC_CO2Array(i) * CompVolArray(i) / 100 + CompVolArray(11) / 100
    
    Do While i < 19
        POC_CO2Vol = POC_CO2Vol + POC_CO2Array(i + 1) * CompVolArray(i + 1) / 100
        i = i + 1
    Loop
    
'Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
    'POC_CO2 = Array(1.38, 3.14, 3.17, 3.2, 3.16, 3.24, 3.25)
    'CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
   ' i = 0
    'Do While i < 7
           
            'If CompWeightArray(i) = FuelType Then
               ' POC_CO2Mass = FuelFlowMass * POC_CO2(i)
               ' Exit Do
           'End If
    'i = i + 1
   ' Loop


End Select

End Function
Public Function POC_H2OVol(FuelType As String, AirVol As Single, AmmoniaVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, IButeneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, SVol As Single, SO2Vol As Single, H2OVol As Single)
Dim POC_H2O As Single
Dim POC_H2OArray As Variant
Dim CompVolArray As Variant
Dim i As Integer

Select Case FuelType

Case "Gas"
    POC_H2OArray = Array(0, 1.5, 0, 2, 3, 4, 5, 4, 6, 7, 0, 0, 0, 0, 1, 0, 1, 0, 0, 1)
    CompVolArray = Array(AirVol, AmmoniaVol, ArgonVol, MethaneVol, EthaneVol, PropaneVol, ButaneVol, IButeneVol, PentaneVol, HexaneVol, CO2Vol, COVol, CVol, N2Vol, H2Vol, O2Vol, H2SVol, SVol, SO2Vol, H2OVol)
    
    i = 0
    POC_H2OVol = POC_H2OArray(i) * CompVolArray(i) / 100
    
    Do While i < 19
        POC_H2OVol = POC_H2OVol + POC_H2OArray(i + 1) * CompVolArray(i + 1) / 100
        i = i + 1
    Loop
    
'Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
    'POC_CO2 = Array(1.38, 3.14, 3.17, 3.2, 3.16, 3.24, 3.25)
    'CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
   ' i = 0
    'Do While i < 7
           
            'If CompWeightArray(i) = FuelType Then
               ' POC_CO2Mass = FuelFlowMass * POC_CO2(i)
               ' Exit Do
           'End If
    'i = i + 1
   ' Loop


End Select

End Function
Public Function POC_N2Vol(FuelType As String, AirVol As Single, AmmoniaVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, IButeneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, SVol As Single, SO2Vol As Single, H2OVol As Single)
Dim POC_N2 As Single
Dim POC_N2Array As Variant
Dim CompVolArray As Variant
Dim i As Integer

Select Case FuelType

Case "Gas"
    POC_N2Array = Array(0, 3.32, 0, 7.53, 13.18, 18.82, 24.47, 22.59, 30.11, 35.76, 0, 1.88, 3.76, 0, 1.88, 3.76, 0, 1.88, 0, 5.65, 3.76, 0, 0)
    CompVolArray = Array(AirVol, AmmoniaVol, ArgonVol, MethaneVol, EthaneVol, PropaneVol, ButaneVol, IButeneVol, PentaneVol, HexaneVol, CO2Vol, COVol, CVol, N2Vol, H2Vol, O2Vol, H2SVol, SVol, SO2Vol, H2OVol)
    
    i = 0
    POC_N2Vol = POC_N2Array(i) * CompVolArray(i) / 100 + CompVolArray(14) / 100
    
    Do While i < 19
        POC_N2Vol = POC_N2Vol + POC_N2Array(i + 1) * CompVolArray(i + 1) / 100
        i = i + 1
    Loop
    
'Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
    'POC_CO2 = Array(1.38, 3.14, 3.17, 3.2, 3.16, 3.24, 3.25)
    'CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
   ' i = 0
    'Do While i < 7
           
            'If CompWeightArray(i) = FuelType Then
               ' POC_CO2Mass = FuelFlowMass * POC_CO2(i)
               ' Exit Do
           'End If
    'i = i + 1
   ' Loop

End Select

End Function
Public Function POC_SO2Vol(FuelType As String, AirVol As Single, AmmoniaVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, IButeneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, SVol As Single, SO2Vol As Single, H2OVol As Single)
Dim POC_SO2 As Single
Dim POC_SO2Array As Variant
Dim CompVolArray As Variant
Dim i As Integer

Select Case FuelType

Case "Gas"
    POC_SO2Array = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 0)
    CompVolArray = Array(AirVol, AmmoniaVol, ArgonVol, MethaneVol, EthaneVol, PropaneVol, ButaneVol, IButeneVol, PentaneVol, HexaneVol, CO2Vol, COVol, CVol, N2Vol, H2Vol, O2Vol, H2SVol, SVol, SO2Vol, H2OVol)
    
    i = 0
    POC_SO2Vol = POC_SO2Array(i) * CompVolArray(i) / 100 + CompVolArray(19) / 100
    
    Do While i < 19
        POC_SO2Vol = POC_SO2Vol + POC_SO2Array(i + 1) * CompVolArray(i + 1) / 100
        i = i + 1
    Loop
    
'Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
    'POC_CO2 = Array(1.38, 3.14, 3.17, 3.2, 3.16, 3.24, 3.25)
    'CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
   ' i = 0
    'Do While i < 7
           
            'If CompWeightArray(i) = FuelType Then
               ' POC_CO2Mass = FuelFlowMass * POC_CO2(i)
               ' Exit Do
           'End If
    'i = i + 1
   ' Loop

End Select

End Function
Public Function RequiredO2ForCombustion(FuelType As String, AirVol As Single, AmmoniaVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, IButeneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, SVol As Single, SO2Vol As Single, H2OVol As Single)
Dim RequiredOxygen As Single
Dim RequiredOxygenArray As Variant
Dim CompVolArray As Variant
Dim CompWeightArray As Variant
Dim i As Integer
Select Case FuelType

    Case "Gas"
        RequiredOxygenArray = Array(0, 0.75, 0, 2, 3.5, 5, 6.5, 6, 8, 9.5, 0, 0.5, 1, 0, 0.5, 0, 1.5, 1, 0, 0)
        CompVolArray = Array(AirVol, AmmoniaVol, ArgonVol, MethaneVol, EthaneVol, PropaneVol, ButaneVol, IButeneVol, PentaneVol, HexaneVol, CO2Vol, COVol, CVol, N2Vol, H2Vol, O2Vol, H2SVol, SVol, SO2Vol, H2OVol)
        
        i = 0
        
        RequiredOxygen = RequiredOxygenArray(i) * CompVolArray(i) / 100 + CompVolArray(16) / 100
       
        Do While i < 19
            RequiredOxygen = RequiredOxygen + RequiredOxygenArray(i + 1) * CompVolArray(i + 1) / 100
           i = i + 1
        Loop
        RequiredO2ForCombustion = RequiredOxygen
        
   ' Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
       ' CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
       ' AirFuelRatioArray = Array(xx, xx, xx, xx, xx, xx, xx)
       ' i = 0
       ' Do While i < 7
               
       '         If CompolArray(i) = FuelType Then
       '             AirFuelRatio = AirFuelRatioArray(i)
       '             Exit Do
       '         End If
      '  i = i + 1
      '  Loop

End Select
End Function

Public Function POC_CO2Mass(FuelType As String, FuelFlowMass As Single, AirFlowMass As Single, AirMass As Single, ArgonMass As Single, MethaneMass As Single, EthaneMass As Single, PropaneMass As Single, ButaneMass As Single, PentaneMass As Single, HexaneMass As Single, CO2Mass As Single, COMass As Single, CMass As Single, N2Mass As Single, H2Mass As Single, O2Mass As Single, H2SMass As Single, H2OMass As Single)

Dim POC_CO2 As Variant
Dim CompWeightArray As Variant
Dim i As Integer

Select Case FuelType

Case "Gas"
    POC_CO2 = Array(0, 0, 2.743, 2.927, 2.994, 3.029, 3.05, 3.06, 0, 1.571, 3.664, 0, 0, 0, 0, 0, 3.25)
    CompWeightArray = Array(AirMass, ArgonMass, MethaneMass, EthaneMass, PropaneMass, ButaneMass, PentaneMass, HexaneMass, CO2Mass, COMass, CMass, N2Mass, H2Mass, O2Mass, H2SMass, CO2Mass)
    
    i = 0
    POC_CO2Mass = FuelFlowMass * POC_CO2(i) * CompWeightArray(i) / 100 + FuelFlowMass * CompWeightArray(8) / 100
    
    Do While i < 15
        POC_CO2Mass = POC_CO2Mass + FuelFlowMass * POC_CO2(i + 1) * CompWeightArray(i + 1) / 100
        i = i + 1
    Loop
Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
    POC_CO2 = Array(1.38, 3.14, 3.17, 3.2, 3.16, 3.24, 3.25)
    CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
    i = 0
    Do While i < 7
           
            If CompWeightArray(i) = FuelType Then
                POC_CO2Mass = FuelFlowMass * POC_CO2(i)
                Exit Do
            End If
    i = i + 1
    Loop


End Select

End Function
Public Function POC_N2Mass(FuelType As String, FuelFlowMass As Single, ExcessAirMass As Single, AirFlowMass As Single, AirMass As Single, ArgonMass As Single, MethaneMass As Single, EthaneMass As Single, PropaneMass As Single, ButaneMass As Single, PentaneMass As Single, HexaneMass As Single, CO2Mass As Single, COMass As Single, CMass As Single, N2Mass As Single, H2Mass As Single, O2Mass As Single, H2SMass As Single, H2OMass As Single)

Dim POC_N2 As Variant
Dim CompWeightArray As Variant
Dim i As Integer

Select Case FuelType
    Case "Gas"
        POC_N2 = Array(0, 0, 13.246, 12.367, 12.047, 11.882, 11.81, 11.74, 0, 1.897, 8.846, 0, 26.353, 0, 4.682, 0, 10.25)
        CompWeightArray = Array(AirMass, ArgonMass, MethaneMass, EthaneMass, PropaneMass, ButaneMass, PentaneMass, HexaneMass, CO2Mass, COMass, CMass, N2Mass, H2Mass, O2Mass, H2SMass, N2Mass)
        
        i = 0
        POC_N2Mass = FuelFlowMass * POC_N2(i) * CompWeightArray(i) / 100 + (ExcessAirMass - AirFlowMass) * 0.7686 + N2Mass * FuelFlowMass / 100
        
        Do While i < 15
            POC_N2Mass = POC_N2Mass + FuelFlowMass * POC_N2(i + 1) * CompWeightArray(i + 1) / 100
            i = i + 1
        Loop
    
    Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
        POC_N2 = Array(4.97, 11.36, 11.1, 10.95, 10.68, 10.59, 10.25)
        CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
        i = 0
        Do While i < 7
               
                If CompWeightArray(i) = FuelType Then
                    POC_N2Mass = FuelFlowMass * POC_N2(i) + (ExcessAirMass - AirFlowMass) * 0.7686
                    Exit Do
                End If
        i = i + 1
        Loop
End Select

End Function

Public Function POC_O2Mass(FuelFlowMass As Single, ExcessAirMass As Single, AirFlowMass As Single, O2Mass As Single)

POC_O2Mass = (ExcessAirMass - AirFlowMass) * 0.2314 + FuelFlowMass * O2Mass / 100

End Function
Public Function AirMass(GasFlowMass As Single, AirFuelRatio As Single)

AirMass = GasFlowMass * AirFuelRatio

End Function
Public Function LHVMass(FuelType As String, AirMass As Single, ArgonMass As Single, MethaneMass As Single, EthaneMass As Single, PropaneMass As Single, ButaneMass As Single, PentaneMass As Single, HexaneMass As Single, CO2Mass As Single, COMass As Single, CMass As Single, N2Mass As Single, H2Mass As Single, O2Mass As Single, H2SMass As Single, H2OMass As Single)

Dim LHVMassArray As Variant
Dim CompWeightArray As Variant
Dim i As Integer

Select Case FuelType
    Case "Gas"
        LHVMassArray = Array(0, 0, 21495, 20418, 19937, 19678, 20485, 19403, 0, 4347, 14093, 0, 51623, 0, 6537, 0)
        CompWeightArray = Array(AirMass, ArgonMass, MethaneMass, EthaneMass, PropaneMass, ButaneMass, PentaneMass, HexaneMass, CO2Mass, COMass, CMass, N2Mass, H2Mass, O2Mass, H2SMass, H2OMass)
        
        i = 0
        LHVMass = LHVMassArray(i) * CompWeightArray(i) / 100
        Do While i < 15
        
            LHVMass = LHVMass + LHVMassArray(i + 1) * CompWeightArray(i + 1) / 100
            i = i + 1
        Loop
Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
        CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
        LHVMassArray = Array(8706, 18790, 18211, 17855, 17790, 17929, 17277)
        i = 0
        Do While i < 7
               
                If CompWeightArray(i) = FuelType Then
                    LHVMass = LHVMassArray(i)
                    Exit Do
                End If
        i = i + 1
        Loop
End Select


End Function

Public Function AirFuelRatio(FuelType As String, AirMass As Single, ArgonMass As Single, MethaneMass As Single, EthaneMass As Single, PropaneMass As Single, ButaneMass As Single, PentaneMass As Single, HexaneMass As Single, CO2Mass As Single, COMass As Single, CMass As Single, N2Mass As Single, H2Mass As Single, O2Mass As Single, H2SMass As Single, H2OMass As Single)

Dim AirFuelRatioArray As Variant
Dim CompWeightArray As Variant
Dim i As Integer
Select Case FuelType

    Case "Gas"
        AirFuelRatioArray = Array(0, 0, 17.195, 16.1, 15.7, 15.5, 15.32, 15.238, 0, 2.468, 11.51, 0, 34.29, 0, 6.093, 0)
        CompWeightArray = Array(AirMass, ArgonMass, MethaneMass, EthaneMass, PropaneMass, ButaneMass, PentaneMass, HexaneMass, CO2Mass, COMass, CMass, N2Mass, H2Mass, O2Mass, H2SMass, H2OMass)
        
        i = 0
        AirFuelRatio = AirFuelRatioArray(i) * CompWeightArray(i) / 100
        Do While i < 15
        
            AirFuelRatio = AirFuelRatio + AirFuelRatioArray(i + 1) * CompWeightArray(i + 1) / 100
            i = i + 1
        Loop
    Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
        CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
        AirFuelRatioArray = Array(6.47, 14.8, 14.55, 14.35, 13.99, 13.88, 13.44)
        i = 0
        Do While i < 7
               
                If CompWeightArray(i) = FuelType Then
                    AirFuelRatio = AirFuelRatioArray(i)
                    Exit Do
                End If
        i = i + 1
        Loop
End Select
    
    
End Function
Public Function FuelFlowMass(FuelType As String, FuelFlowVol As Single, AirVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, H2OVol As Single)

Dim SpecificVolume, densityArray, CompWeightArray As Variant
Dim Density, ComponentWeightTotal As Single
Dim Components(16) As Single
Dim i, step, neg As Integer

'load component percentages into array
Components(0) = AirVol / 100
Components(1) = ArgonVol / 100
Components(2) = MethaneVol / 100
Components(3) = EthaneVol / 100
Components(4) = PropaneVol / 100
Components(5) = ButaneVol / 100
Components(6) = PentaneVol / 100
Components(7) = HexaneVol / 100
Components(8) = CO2Vol / 100
Components(9) = COVol / 100
Components(10) = CVol / 100
Components(11) = N2Vol / 100
Components(12) = H2Vol / 100
Components(13) = O2Vol / 100
Components(14) = H2SVol / 100
Components(15) = H2OVol / 100

'specific volume at standard conditions in ft3/lb
SpecificVolume = Array(13.063, 24.017, 23.574, 12.455, 8.361, 6.321, 5.252, 4.398, 8.547, 13.506, 31.517, 13.372, 187.97, 11.819, 10.978, 21.017)

'initialize mass flow
FuelFlowMass = 1

Select Case FuelType
    Case "Gas"
        step = 10
        neg = 1
        i = 0
          Do While Abs(ComponentWeightTotal - 1) > 0.0001
              i = 0
              ComponentWeightTotal = FuelFlowVol * Components(i) / (SpecificVolume(i) * FuelFlowMass)
              Do While i < 15
                  ComponentWeightTotal = ComponentWeightTotal + FuelFlowVol * Components(i + 1) / (SpecificVolume(i + 1) * FuelFlowMass)
                  i = i + 1
              Loop
              FuelFlowMass = FuelFlowMass + step
              If (ComponentWeightTotal - 1) * neg < 0 Then
                  step = step / -10
                  neg = neg * -1
              End If
          Loop
          
     Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
        CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
        densityArray = Array(6.607, 6.073, 6.877, 7.303, 7.585, 7.785, 8.219)
        i = 0
        Do While i < 7
               
                If CompWeightArray(i) = FuelType Then
                    Density = densityArray(i)
                    Exit Do
                End If
        i = i + 1
        Loop
        FuelFlowMass = Density * FuelFlowVol

End Select



End Function
 

Public Function FuelFlowVol(FiringRate As Single, LHV_Vol As Single)
'calculates volumetric gas flow given firing rate(LHV basis) and lower heating value
FuelFlowVol = FiringRate * 10 ^ 6 / LHV_Vol

End Function
Public Function HHV_Vol(FuelType As String, AirVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, H2OVol As Single)
 
Dim Components(16) As Single
Dim HHV, CompWeightArray As Variant
Dim i As Integer
Dim CompSum As Single

Select Case FuelType

    Case "Gas"
        'load component percentages into array
        Components(0) = AirVol / 100
        Components(1) = ArgonVol / 100
        Components(2) = MethaneVol / 100
        Components(3) = EthaneVol / 100
        Components(4) = PropaneVol / 100
        Components(5) = ButaneVol / 100
        Components(6) = PentaneVol / 100
        Components(7) = HexaneVol / 100
        Components(8) = CO2Vol / 100
        Components(9) = COVol / 100
        Components(10) = CVol / 100
        Components(11) = N2Vol / 100
        Components(12) = H2Vol / 100
        Components(13) = O2Vol / 100
        Components(14) = H2SVol / 100
        Components(15) = H2OVol / 100
        
        i = 0
        
        'CompSum = Components(0)
        
        For i = 0 To 15
            CompSum = CompSum + Components(i)
        Next i
        
        
        'Higher heating values for components in BTU/SCF
        HHV = Array(0, 0, 1013, 1792, 2592, 3373, 4008.7, 4755.9, 0, 321.9, 0, 0, 325, 0, 646, 0)
        
        i = 0
        HHV_Vol = HHV(i) * Components(i)
        Do While i < 15
            HHV_Vol = HHV_Vol + HHV(i + 1) * Components(i + 1)
            i = i + 1
        Loop
        
        If Abs(CompSum - 1) > 0.001 Then
            HHV_Vol = "Sum of components does not equal 100%"
        End If
    
    Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
        HHV = Array(64728.779, 122613.87, 133571.971, 138705.879, 142931.74, 147206.565, 148977.594)
        CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
        i = 0
        Do While i < 7
               
                If CompWeightArray(i) = FuelType Then
                    HHV_Vol = HHV(i)
                    Exit Do
                End If
        i = i + 1
        Loop


End Select
    

End Function

Public Function LHV_Vol(FuelType As String, AirVol As Single, ArgonVol As Single, MethaneVol As Single, EthaneVol As Single, PropaneVol As Single, ButaneVol As Single, PentaneVol As Single, HexaneVol As Single, CO2Vol As Single, COVol As Single, CVol As Single, N2Vol As Single, H2Vol As Single, O2Vol As Single, H2SVol As Single, H2OVol As Single)
 
Dim Components(16) As String
Dim LHV, CompWeightArray As Variant
Dim i As Integer
Dim CompSum As Single

Select Case FuelType
    Case "Gas"
        'load component percentages into array
        Components(0) = AirVol / 100
        Components(1) = ArgonVol / 100
        Components(2) = MethaneVol / 100
        Components(3) = EthaneVol / 100
        Components(4) = PropaneVol / 100
        Components(5) = ButaneVol / 100
        Components(6) = PentaneVol / 100
        Components(7) = HexaneVol / 100
        Components(8) = CO2Vol / 100
        Components(9) = COVol / 100
        Components(10) = CVol / 100
        Components(11) = N2Vol / 100
        Components(12) = H2Vol / 100
        Components(13) = O2Vol / 100
        Components(14) = H2SVol / 100
        Components(15) = H2OVol / 100
        
        i = 0
        
        'CompSum = Components(0)
        
        For i = 0 To 15
            CompSum = CompSum + Components(i)
        Next i
        
        
        'Higher heating values for components in BTU/SCF
        LHV = Array(0, 0, 911, 1622, 2322, 3018, 3900, 4412, 0, 321, 0, 0, 275, 0, 595, 0)
        
        i = 0
        LHV_Vol = LHV(i) * Components(i)
        Do While i < 15
            LHV_Vol = LHV_Vol + LHV(i + 1) * Components(i + 1)
            i = i + 1
        Loop
        
        If Abs(CompSum - 1) > 0.001 Then
            LHV_Vol = "Sum of components does not equal 100%"
        End If

Case "methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil"
        LHV = Array(57520.542, 114111.67, 125237.047, 130395.065, 134937.15, 139577.265, 141999.663)
        CompWeightArray = Array("methanol", "gasoline", "#1 oil", "#2 oil", "#4 oil", "#5 oil", "#6 oil")
        i = 0
        Do While i < 7
               
                If CompWeightArray(i) = FuelType Then
                    LHV_Vol = LHV(i)
                    Exit Do
                End If
        i = i + 1
        Loop


End Select


End Function

Public Function ComponentWeight(Component As String, ComponentVol As Single, GasFlowVol As Single, GasFlowMass As Single)

Dim SpecificVolumeArray As Variant
Dim ComponentArray As Variant
Dim SpecificVolume As Single
Dim i As Integer

ComponentArray = Array("Air", "Argon", "Methane", "Ethane", "Propane", "Butane", "Pentane", "Hexane", "Carbon Dioxide", "Carbon Monoxide", "Carbon", "Nitrogen", "Hydrogen", "Oxygen", "Hydrogen Sulfide", "Water", "#2 Oil")

SpecificVolumeArray = Array(13.1545, 29.4977, 23.6489, 12.6168, 8.6036, 6.5273, 5.2581, 4.4113, 8.6202, 13.5443, 31.6146, 13.5491, 188.1822, 11.8555, 11.1332, 21.0577, 0.14)

For i = 0 To 16
    If Component = ComponentArray(i) Then
        SpecificVolume = SpecificVolumeArray(i)
    End If
Next i

ComponentWeight = ComponentVol / 100 * GasFlowVol / (SpecificVolume * GasFlowMass) * 100

End Function


Public Function ComponentWeightPercent(Component As String, ComponentVol As Single, FluidMW As Single)
Dim MolecularWeightArray As Variant
Dim ComponentArray As Variant
Dim i As Integer
Dim MolecularWeight As Single

ComponentArray = Array("Air", "Ammonia", "Argon", "Methane", "Ethane", "Propane", "Butane", "IButene", "Pentane", "Hexane", "Carbon Dioxide", "Carbon Monoxide", "Carbon", "Nitrogen", "Hydrogen", "Oxygen", "Hydrogen Sulfide", "sulfur.fld", "SO2", "water")
MolecularWeightArray = Array(28.96, 17.03, 39.95, 16.04, 30.07, 44.1, 58.12, 56.11, 72.15, 86.18, 44.01, 28.01, 16.04, 28.01, 2.02, 32, 34.08, 30.07, 64.06, 18.02)

For i = 0 To 19

If Component = ComponentArray(i) Then
MolecularWeight = MolecularWeightArray(i)
End If
Next i
ComponentWeightPercent = ComponentVol * MolecularWeight / FluidMW

End Function
-------------------------------------------------------------------------------
VBA MACRO ConvectionFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/ConvectionFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Public Function WorkbookIsOpen(wbname) As Boolean
'   Returns TRUE if the workbook is open
    Dim x As Workbook
    On Error Resume Next
    Set x = Workbooks(wbname)
    If Err = 0 Then WorkbookIsOpen = True _
        Else WorkbookIsOpen = False
End Function
Function Ttubeout(Itube%, q, Pflow, Cp, FlowConfig%)
'
'  Itube = current tube row
'  Q = heat flow, mm Btu/hr
'  Pflow = process flow lb/hr
'  Cp = Specific heat Btu/lb-F
'  flowconfig = flow configuration 1= counterflow, 2=parallel flow
'
'
If FlowConfig = 1 Then
'  Counterflow
         If Itube = 1 Then
            'Initial outlet temperature for counterflow
            Ttubeout = Worksheets("data").range("b14")
         ElseIf Itube > 1 Then
            Ttubeout = ActiveCell.offset(-1, -1)
            End If
End If
End Function

Function Ttubeouttemp(Itube%, NT, NR, Nsplits, FlowConfig%)
'   Ignore this subroutine for now
'  Itube = current tube row
'
Dim Ntubes As Integer
Dim NP As Integer       ' Number of parallel tube zones
Dim Npass As Integer    'Number of current parallel pass
Dim NrowP_s
'
Ntubes = NT * NR
NP = Int(Ntubes / Nsplits)
NrowP_s = Int(NR / NP)
Npass = Int(Itube / NrowP_s + 0.99)
'
If FlowConfig = 1 Then
'  Counterflow
         If Itube = 1 Then
            'Initial outlet temperature for counterflow
            Ttubeouttemp = Worksheets("data").range("b14")
         ElseIf Npass > 1 Then
            '
            End If
End If
'
End Function

Function Cells(Row, Col)

Cells = "=SUM(" & Row & ":" & Worksheets("Row Calcs").Cells(Row, Col + 3).Address & ")"

End Function
Function FuelTrainCells(Row, Col)

FuelTrainCells = Worksheets("Fuel Train Equip Area").Cells(i, 3)

End Function

'
'   Declare PI as public variable for general use within project
'
'Public Const pi As Single = 3.1415927
'Public message As String
'
'
Public Function APO(OD, Nfins, fthick) As Single
'
'   Calculate primary outside surface area)
'
' od - Outside tube diameter d,o
' Nfins - Number of fins per ft tube
' fthick - fin thickness
'
APO = pi * OD * (1 - Nfins * fthick) / 12
'
End Function
'
Public Function Ao(OD, Nfins, fthick, fheight, ws, ftype%) As Single
'
'Calculate total outside area
'
' od - Outside tube diameter d,o, in.
' Nfins - Number of fins per ft tube
' fthick - fin thickness, in.
' fheight - fin height, in.
' ws - fin segment width, in.
' ftype - fin type:  1=solid, 2=segmented
'
Dim df As Single    'local intermediate value OD of fins
Dim da As Single    'local intermediate value avg dia.
'
If ftype = 1 Then
    df = OD + 2 * fheight
    da = OD + fheight
    Ao = pi * (OD * (1 - Nfins * fthick) + Nfins * (2 * fheight * da + fthick * df)) / 12
ElseIf ftype = 2 Then
    da = ws + fthick
    Ao = pi * (OD * (1 - Nfins * fthick) + OD * Nfins * (2 * fheight * da + fthick * ws) / ws) / 12
ElseIf ftype < 1 Or ftype > 2 Then
    MsgBox "Error in fin type specification)"
    Ao = "  "
End If
'
End Function


'
Public Function Corr1(Re) As Single
'
'   Calculate Reynold correction factor C1
'
Corr1 = 0.25 * Re ^ -0.35
'
End Function
'
Function Corr3(fheight, sf, tubeconfig%, ftype%) As Single
'
'   Calculate geometry Correction factor C3
'
'   fheight - fin height, in.
'   sf - fin spacing, in.
'   tubeconfig - tube configuration:  1=staggered, 2= in-line
'   ftype - fin type:  1=solid, 2=segmented
'
Dim ci As Single
'
If ftype = 1 Then
'for solid fins:
    ci = 0.65 * Exp(-0.25 * fheight / sf)
    If tubeconfig = 1 Then
'       Staggered
        Corr3 = 0.35 + ci
    ElseIf tubeconfig = 2 Then
'       Inline
        Corr3 = 0.2 + ci
    ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
        MsgBox "Error in configuration specification For C3"
        Corr3 = "   "
    End If
ElseIf ftype = 2 Then
'for segmented fins
    ci = Exp(-0.35 * fheight / sf)
    If tubeconfig = 1 Then
'       Staggered
        Corr3 = 0.55 + 0.45 * ci
    ElseIf tubeconfig = 2 Then
'       Inline
        Corr3 = 0.35 + 0.5 * ci
    ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
        MsgBox "Error in configuration specification for C3"
        Corr3 = "   "
    End If
ElseIf ftype < 1 Or ftype > 2 Then
    MsgBox "Error in fin type specification for C3"
    Corr3 = "   "
    End If
'
End Function
'
'
Function Corr5(NR, Pl, Pt, tubeconfig%) As Single
'
'   Calculate Non-equilateral & row Correction factor C5
'
'   Nr - Number of Tube rows in direction of flow
'   Pl - longitudinal tube pitch
'   Pt - transverse tube pitch
'   tubeconfig - tube configuration:  1=staggered, 2= in-line
'
If tubeconfig = 1 Then
'
    Corr5 = 0.7 + (0.7 - 0.8 * Exp(-0.15 * (NR ^ 2))) * Exp(-1 * Pl / Pt)
    '
ElseIf tubeconfig = 2 Then
'
    Corr5 = 1.1 - (0.75 - 1.5 * Exp(-0.7 * NR)) * Exp(-2 * Pl / Pt)
'
ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
    MsgBox "Error in configuration specification for C5"
    Corr5 = "   "
End If
End Function
'
Function Jfactor(CF1, CF3, CF5, df, OD, Tb, ts)
'
'   Calulates Colburn number
'
'   CF1-CF5 are correctuion factors as per Corr1-Corr5
'   df - outside dia of fins, in.
'   od - outside tube dia, in.
'   Tb - Avg outside fluid temperature, F
'   ts - average fin temperature estimated as, F
'   Ti+0.3(Tb-Ti) where Ti=Avg. fluid temp inside tubes
'
Jfactor = CF1 * CF3 * CF5 * (df / OD) ^ 0.5 * ((Tb + 460) / (ts + 460)) ^ 0.25
End Function
'
Function Hc(Jfactor, g, cpg, kb, Visc, OD, TubeType%, tubeconfig%) As Single
'
'   Calulate outside convection coefficient (Shell side)in Btu/hr-ft2-F
'
'   Jfactor - Colburn number
'   G - Mass velocity, lb/hr-ft2
'   cpg - fluid specific heat at bulk temp, Btu/lb-F
'   kb - fluid thermal conductivity, Btu/ft-F
'   Visc -fluid viscosity
'   od = outside tube diameter
'   Tubetype - 1=finned, 2=bare
'   tubeconfig - tube configuration:  1=staggered, 2= in-line
'
Dim Pr As Single        'Prandtl Number
Dim hci As Single       'intermediate calc.
'
Pr = Visc * cpg / kb
'
If TubeType = 1 Then
    'hc for finned tube
    Hc = Jfactor * g * cpg * (1 / Pr) ^ 0.67
ElseIf TubeType = 2 Then
    'hc for bare tube
    hci = kb * 12 / OD * (Pr ^ 0.333 * ((OD / 12) * (g / Visc))) ^ 0.6
    If tubeconfig = 1 Then
    '   for staggered
        Hc = 0.33 * hci
    ElseIf tubeconfig = 2 Then
    '   for in-line
        Hc = 0.26 * hci
    ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
        MsgBox "Error in configuration specification for heat transfer coeff. calulation"
        Hc = "   "
    End If
ElseIf TubeType < 1 Or TubeType > 2 Then
    MsgBox "Error in tube type (finned/bare) specification for heat transfer coeff. calulation"
    Hc = "   "
End If
'
End Function
'
Function Efactor(ho, lfht, tf, kf, OD, df, ws, ftype%) As Single
'
'   Calulates Fin efficiency
'
'   ho - outside convective coefficient
'   lfht - fin height, in.
'   tf - fin thickness, in.
'   kf - fin thermal conductivity
'   od - tube diameter, outside, in.
'   df - fin outside diameter, in.
'   ws - fin segment width
'   ftype - fin type:  1=solid, 2=segmented
'
Dim b As Single
Dim M As Single
Dim x As Single
Dim y As Single
Dim lf
'
lf = lfht
b = lf + tf / 2
'
' check if fin thickness 0 for valid calulation before starting calulation.
' If thicness is 0 then set E=0.
'
If tf = 0 Then
    Efactor = 0
    GoTo Invalid_fin
End If
'
'If fin thickness not 0, proceed with calulation
If ftype = 1 Then
'
    M = Sqr(ho / (6 * kf * tf))
    x = (Exp(2 * M * b) - 1) / (Exp(2 * M * b) + 1) 'calculates tanh(mb)
    x = x / (M * b)                                 'calculate X
    y = x * (0.7 + 0.3 * x)
    Efactor = y * (0.45 * Log(df / OD) * (y - 1) + 1)
ElseIf ftype = 2 Then
    M = Sqr(ho * (tf + ws) / (6 * kf * tf * ws))
    x = (Exp(2 * M * b) - 1) / (Exp(2 * M * b) + 1) 'calculates tanh(mb)
    x = x / (M * b)                                 'calculate X
    Efactor = x * (0.9 + 0.1 * x)
ElseIf ftype < 1 Or ftype > 2 Then
    MsgBox "Error in fin type specification For fin efficiency calculation"
    Efactor = "   "
End If
'
Invalid_fin:
'
End Function


'  Module contains pressure drop calulation for the shell side
'
Function Corr2(Re) As Single
'
Corr2 = (0.07 + 8 * Re ^ -0.45)
'
End Function
'
Function Corr4(OD, Pt, lf_hfin, sf, tubeconfig%, ftype%) As Single
'   Calculate Correction factor C4
'
'   od - outside tube dia., in.
'   Pt - transverse tube pitch
'   lf_hfin - fin height, in.
'   sf - fin spacing, in.
'   tubeconfig - tube configuration:  1=staggered, 2= in-line
'   ftype - fin type:  1=solid, 2=segmented
'
Dim P As Single
'
If ftype = 1 Then
'for solid fins:
    If tubeconfig = 1 Then
        'staggered tubes
        P = -0.7 * (lf_hfin / sf) ^ 0.2
        Corr4 = 0.11 * (0.05 * Pt / OD) ^ P
    ElseIf tubeconfig = 2 Then
        'inline tubes
        P = -1.1 * (lf_hfin / sf) ^ 0.15
        Corr4 = 0.08 * (0.15 * (Pt / OD)) ^ P
    ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
        MsgBox "Error in configuration specification For C4"
        Corr4 = "   "
    End If
ElseIf ftype = 2 Then
'for segmented fins
    If tubeconfig = 1 Then
        'staggered tubes
        P = -0.7 * (lf_hfin / sf) ^ 0.23
        Corr4 = 0.11 * (0.05 * Pt / OD) ^ P
    ElseIf tubeconfig = 2 Then
        'inline tubes
        P = -1.1 * (lf_hfin / sf) ^ 0.2
        Corr4 = 0.08 * (0.15 * (Pt / OD)) ^ P
    ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
        MsgBox "Error in configuration specification For C4"
        Corr4 = "   "
ElseIf ftype < 1 Or ftype > 2 Then
    MsgBox "Error in fin type specification for C3"
    Corr4 = "   "
    End If
End If
End Function
'
'
Function Corr6(NR, Pl, Pt, tubeconfig%) As Single
'
'
'   Calculate Non-equilateral & row Correction factor C5
'
'   Nr - Number of Tube rows in direction of flow.  If value of 0 is passed through
'        then 4 rows will be assumed
'   Pl - longitudinal tube pitch
'   Pt - transverse tube pitch
'   tubeconfig - tube configuration:  1=staggered, 2= in-line
'
'
If NR = 0 Then NR = 4
'
If tubeconfig = 1 Or NR = 1 Then
'
    Corr6 = 1.1 + (1.8 - 2.1 * Exp(-0.15 * (NR ^ 2))) * Exp(-2 * Pl / Pt) _
            - (0.7 - 0.8 * Exp(-0.15 * (NR ^ 2))) * Exp(-0.6 * Pl / Pt)
    '
ElseIf tubeconfig = 2 And NR > 1 Then
'
    Corr6 = 1.6 - (0.75 - 1.5 * Exp(-0.7 * NR)) * Exp(-0.2 * (Pl / Pt) ^ 2)
'
ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
    MsgBox "Error in configuration specification for C6"
    Corr6 = "   "
End If
End Function
'
Function ffactor(CF2, CF4, CF6, df, OD, tubeconfig%) As Single
'
'Calulate flow factor for pressure drop equation
'
'   CF1-CF5 are correctuion factors as per Corr1-Corr5
'   df - outside dia of fins, in.
'   od - outside tube dia, in.
'
If tubeconfig = 1 Then
    '  staggered
    ffactor = CF2 * CF4 * CF6 * Sqr(df / OD)
ElseIf tubeconfig = 2 Then
    '  in-line
    ffactor = CF2 * CF4 * CF6 * df / OD
ElseIf tubeconfig < 1 Or tubeconfig > 2 Then
    MsgBox "Error in tube configuration specification in calulating f-factor)"
    ffactor = "  "
End If
'
End Function
'
Function DPshell(Gn, An, Ad, NR, dens1, dens2, ffl) As Single
'
'   Calculate pressure drop shell side, in. water
'
'   Gn - mass velocity, lb/hrft2
'   An - net free area in tube row, ft2
'   Ad - crossectional area of duct enclosing bundle, ft2
'   Nr - number of rows,
'   dens1 - density at inlet, lb/ft3
'   dens2 - density at outlet, lb/ft3
'   ffactor - flow factor f
'
Dim beta As Single, a As Single, densb As Single
'
'Calc Beta
beta = (An / Ad) ^ 2
'
'Avg density
densb = (dens1 + dens2) / 2
'
'a=factor
a = ((1 + beta) / (4 * NR)) * densb * ((1 / dens2) - (1 / dens1))
'
DPshell = (ffl + a) * Gn ^ 2 * NR / (densb * 1083000000)
'
'
'Message "Beta=" & Beta
'Message = Message & vbCrLf & "Ffactor" & ffactor
'Message = Message & vbCrLf & "A=" & A
'MsgBox Message
'
End Function

'
'

'
Function MBL(od_tube, Pitchl, Pitcht) As Single
'
'   Calculate mean beam radiating length in inches
'   Source:  ESCOA Manual page E-5, graph
'
'   od - outside tube diameter, in.
'   Pl - longitudinal tube pitch, in.
'   Pt - transvers tube pitch, in.
'
Dim Plratio As Single
Dim Ptratio As Single
Dim LOD1 As Single, LOD2 As Single, LOD As Single
Dim plfrac As Single
'
Plratio = Pitchl / od_tube
Ptratio = Pitcht / od_tube
'
'  Calculate L/od, lower limit
Select Case Plratio
    Case 0 To 0.9999: LOD1 = 0
    Case 1 To 1.9999: LOD1 = 0.08 * Ptratio - 0.05
    Case 2 To 2.9999: LOD1 = 0.177 * Ptratio - 0.06
    Case 3 To 3.9999: LOD1 = 0.27 * Ptratio - 0.09
    Case 4 To 4.9999: LOD1 = 0.36 * Ptratio - 0.09
    Case Is >= 5: LOD1 = 0.44 * Ptratio - 0.09
End Select
'
'  Calculate L/od, Upper limit
Select Case Plratio
    Case 0 To 0.9999: LOD2 = 0.08 * Ptratio - 0.05
    Case 1 To 1.9999: LOD2 = 0.177 * Ptratio - 0.06
    Case 2 To 2.9999: LOD2 = 0.27 * Ptratio - 0.09
    Case 3 To 3.9999: LOD2 = 0.36 * Ptratio - 0.09
    Case 4 To 4.9999: LOD2 = 0.44 * Ptratio - 0.09
    Case Is >= 5: LOD2 = 0.44 * Ptratio - 0.09
End Select
'
'Interpolate for actual L/od
plfrac = Plratio - Int(Plratio)
LOD = (LOD2 - LOD1) * plfrac + LOD1
'
MBL = LOD * od_tube
'
End Function
'
Function hrad(Tgas, Tinside, ppCO2, ppH2O, ab, Ao, OD, Pl, Pt) As Single
'
'
'   Tgas - Avg. gas temperature, shell side, F
'   Tinside - fluid temp inside, F
'   ppCO2=partial pressure CO2
'   ppH2O=partial pressure H2O
'   Ab = bare tube outside surface area, ft2/ft
'   Ao = Total Outside Surface area
'   MBL = Mean Beam Length, ft.
'
Dim ts As Single
Dim offset As Single  'intermediate value
Dim ppgas As Single
Dim MBeamL As Single
Dim tsfrac As Single
Dim gamma As Single
Dim TsrangeErr As Integer
'
'  Calculate Gamma radiation factor first
'
' Estimate avg. fin temperature
ts = Tinside + 0.3 * (Tgas - Tinside)
'
'Range Check:
If ts > 1400 Then
    tsfrac = 1
    TsrangeErr = 2
ElseIf ts < 400 Then
    tsfrac = 1
    TsrangeErr = 1
End If
'
offset = 0.00222 * ts - 1.58
gamma = 0.0022 * Tgas + offset
'
'Calculate radiative heat transfer coefficient
'
ppgas = ppCO2 + ppH2O
MBeamL = MBL(OD, Pl, Pt)
hrad = 2.2 * gamma * Sqr(ppgas * MBeamL) * (ab / Ao) ^ 0.75
'
'  At low Temperatures, calulation is not valid.  Trap error and set=0
'
If hrad < 0 Then hrad = 0
'
'Message = "Ts=" & Ts
'Message = Message & vbCrLf & "Gamma=" & Gamma
'Message = Message & vbCrLf & "ppgas=" & ppgas
'MsgBox Message
'
End Function


Sub ShowAddress()
Dim MC As range
'
Set MC = Worksheets("Data").range("TubeInput_wt")
MsgBox MC.Address()                              ' $A$1
MsgBox MC.Address(RowAbsolute:=False)            ' $A1
MsgBox MC.Address(ReferenceStyle:=xlR1C1)        ' R1C1
MsgBox MC.Address(ReferenceStyle:=xlR1C1, _
    RowAbsolute:=False, _
    ColumnAbsolute:=False, _
    RelativeTo:=Worksheets(1).Cells(3, 3))        ' R[-2]C[-2]

End Sub

Sub CopyFormat()
'
' CopyFormat Macro
' Macro recorded 2/18/2003 by Tom & Frances Wechsler
'

'
    Selection.Copy
    range("E19:I19").Select
    Selection.PasteSpecial Paste:=xlFormats, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
End Sub
Sub LiquidSort()
'
' LiquidSort Macro
' Macro recorded 2/19/2003 by Tom & Frances Wechsler
'

'
    range("A81:M150").Select
    Selection.Sort Key1:=range("A81"), Order1:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    range("A81").Select





End Sub

Function FinTemp(e, Twall, Tb_shell) As Single
'
'   E - fin efficiency
'   Twall - Tubewall Temperature
'   Tb_shell - Bulk outside fluid temperature
'
Dim Theta As Single
'
' Calc theta for fin temp calc.
Theta = 1.36 - 1.337 * e
'
FinTemp = Twall + Theta * (Tb_shell - Twall)
'
End Function



Sub CellColor()
'
' cellcolor Macro
' Macro recorded 3/5/2003 by Tom Wechsler
'

'
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 5
    End With
End Sub

Sub CellBold()
'
' cellbold Macro
' Macro recorded 3/10/2003 by Tom Wechsler
'

'
    range("D19:D33").Select
    ActiveSheet.Unprotect
    Selection.Font.Bold = False
End Sub
Sub copyvalue()
'
' copyvalue Macro
' Macro recorded 3/10/2003 by Tom Wechsler
'

'
    range("C19:C33").Select
    Selection.Copy
    range("C19").Select
    Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=False
    Application.CutCopyMode = False
End Sub
Sub Cellformula()
'
' Cellformula Macro
' Macro recorded 3/10/2003 by Tom Wechsler
'

'
    range("C19").Select
    ActiveCell.FormulaR1C1 = "=Properties!R[33]C[-1]"
    range("C20").Select
End Sub

'
Public Function Ac(OD, Nfins, fthick, fheight) As Single
'
'Calculate projected finned tube crossectional area
'
' od - Outside tube diameter d,o
' Nfins - Number of fins per ft tube
' fthick - fin thickness
' fheight - fin height
'
Ac = (OD + 2 * fheight * fthick * Nfins) / 12
'
End Function
'
Public Function An(Ad, Ac, LTube, NT, A_Baffle) As Single
'
'   Calculate net free area in tube row
'
'   Ad - Crossectional area of duct
'   Ac - finned tube projectd crossect. area
'   Ltube - effective finned length
'   Nt - number of tubes per row
'   A_baffle - crossectional area of flow obstructions other than
'   finned tubes, eg. bafflesbends, bare tubes
'
An = Ad - Ac * LTube * NT - A_Baffle
'
End Function
'
Function Prop(Prop1, Prop2, t1, t2, T) As Single
'
'   interpolation function returns value for property at T given properties
'   at T1 and T2
'
Prop = (Prop2 - Prop1) / (t2 - t1) * (T - t1) + Prop1
'
End Function



-------------------------------------------------------------------------------
VBA MACRO EngineeringFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/EngineeringFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
'
Public Const pi = 3.14159265658979
Public Const g = 32.17
Public Const R = 1545.33
Function Interpolate(x1 As Single, x As Single, x2 As Single, y1 As Single, y2 As Single)
Dim y As Single
If (x1 > x And x > x2) Or (x1 < x And x < x2) Then
Interpolate = ((y2 - y1) / (x2 - x1) * (x - x1)) + y1
Else
Interpolate = "bad x"
End If
End Function
Function TubeWallTemp(Flux As Single, OD As Single, ID As Single, Hi As Single, WallThickness As Single, k As Single, TempOut As Single)
TubeWallTemp = Flux * (OD / ID) * (1 / Hi) + Flux * OD / (OD - WallThickness) * (WallThickness / (k * 12)) + TempOut
End Function
Function DINFlangeRating(PClass As String, temp As Single)
Dim i As Integer
Dim Pressure, TempArray As Variant
i = 0

TempArray = Array(57, 122, 212, 302, 392, 482, 572, 662, 752)

Select Case PClass
    Case "PN16"
        Pressure = Array(232, 225, 219, 199, 184, 173, 159, 152, 148)
    Case "PN25"
        Pressure = Array(362, 352, 342, 312, 287, 270, 249, 239, 232)
    Case "PN40"
        Pressure = Array(580, 564, 550, 499, 461, 434, 400, 383, 373)
    Case "PN63"
        Pressure = Array(913, 889, 866, 787, 726, 683, 631, 605, 580)
    Case "PN100"
        Pressure = Array(1450, 1411, 1373, 1248, 1153, 1083, 1000, 958, 931)
End Select

Do While temp >= TempArray(i)

    DINFlangeRating = (temp - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (Pressure(i + 1) - Pressure(i)) + Pressure(i)
    i = i + 1
Loop

End Function
Function Hco_liquid_cross(diameter As Single, k As Single, Pr As Single, Re As Single)
Dim Cl, n As Single

If Re > 40 And Re <= 4000 Then
        Cl = 0.683
        n = 0.466
    ElseIf Re > 4000 And Re <= 40000 Then
        Cl = 0.193
        n = 0.618
    ElseIf Re > 40000 Then
        Cl = 0.0266
        n = 0.805
End If

Hco_liquid_cross = Cl * Pr ^ (1 / 3) * Re ^ n

End Function
Function Hco_flat_plate(l As Single, k As Single, Pr As Single, Re As Single)

'Hco_flat_plate = Pr ^ (1 / 3) * (0.036 * Re ^ 0.8 - 836) * k / L
Hco_flat_plate = 0.0288 * Re ^ 0.8 * Pr * k / (1 + 0.849 * Re ^ -0.1 * ((Pr - 1) + Log((5 * Pr + 1) / 6)))

End Function
Function TubeWallTemperature(Flux As Single, OD As Single, ID As Single, Hi As Single, Tubewall As Single, k As Single, FluidTemp As Single)

TubeWallTemperature = Flux * (OD / ID) * (1 / Hi) + Flux * (OD / (OD - Tubewall)) * (Tubewall / (k * 12)) + FluidTemp

End Function

Function FGRDuctSize(FGRDensity As Single, FGRMassFlow As Single)

Dim DuctSizeArray As Variant
Dim k As Single
Dim FGRVolFlow As Single
Dim DuctID As Single
Dim DuctArea As Single
Dim DuctVelocity As Single
Dim DuctSize As Single
Dim DuctSCH As String

k = 0
DuctVelocity = 1000
DuctSizeArray = Array(2, 2.5, 3, 3.5, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36)
FGRVolFlow = FGRMassFlow / FGRDensity
FGRVolFlow = FGRVolFlow / 3600

If FGRMassFlow = 0 Then
    FGRDuctSize = "N/A"
    Exit Function
End If

Do While k < 28
    If DuctVelocity > 40 Then
        DuctSize = DuctSizeArray(k)
            If DuctSize < 26 Then
                DuctSCH = "SCH10S"
            Else: DuctSCH = "SCH10"
            End If
        DuctID = PipeID(PipeOD(DuctSize), PipeWall(DuctSize, DuctSCH))
        DuctArea = Area(DuctID) / 144
        DuctVelocity = FGRVolFlow / DuctArea
    Else: Exit Do
    End If
k = k + 1
Loop

FGRDuctSize = DuctSize

End Function
Function VentDischargeLineSize(Duty As String)

'expansion volume ft3/h
'Assumed Expansion Line Size, in
'Expansion line ID-Calculated, in
'Selected Discharge and Vent Line Size, NPS inches
    'Below is Information Only
    'DutyMMBtuhr = Array(0.5, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10, 12.5, 17.5, 20, 25, 30, 35, 40, 50, 60)
    'EXPVOLArray = Array(10.87, 21.73, 32.6, 43.47, 54.34, 65.2, 86.94, 108.67, 130.41, 173.88, 217.34, 271.68, 380.35, 434.69, 543.36, 652.03, 760.7, 869.38, 1086.72, 1304.06)
    'AssumedExpansionLineSizeArray = Array(1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4)

    'ExpansionlineinreferencefromDIN4754 = Array(0.75, 0.75, 0.75, 1, 1.25, 1.25, 1.5, 1.5, 1.5, 2, 2, 2, 2, 2.5, 2.5, 2.5, 2.5, 3, 3, 4)
    'NominalSizeofDischargeandVentPipesreferencefromDIN4754 = Array(1, 1, 1, 1.25, 1.5, 1.5, 2, 2, 2, 2.5, 2.5, 2.5, 2.5, 3, 3, 3, 3, 4, 4, 6)

Dim SelectedVentLineSizeArray As Variant
Dim DutyArray As Variant
Dim i As Integer

DutyArray = Array(0.5, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10, 12.5, 15, 17.5, 20, 25, 30, 35, 40, 50, 60)
SelectedVentLineSizeArray = Array(1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 3, 4, 4, 6)
i = 0

Do While i < 20

 If Duty = DutyArray(i) Then
     VentDischargeLineSize = SelectedVentLineSizeArray(i)
     
    Exit Do
    
  Else: i = i + 1
 End If
Loop

End Function
Function CADamperSize(CADuctSize As Single)
Dim DamperSizeArray As Variant
Dim i As Single

DamperSizeArray = Array(2, 2.5, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 24, 30, 36)
i = 0

Do While i < 16
    If CADuctSize = DamperSizeArray(i) Then
        CADamperSize = DamperSizeArray(i)
        Exit Do
    Else: i = i + 1
    End If
Loop

End Function
Function CADamperCv(DamperSize As Single)
Dim DamperSizeArray As Variant
Dim DamperCVArray As Variant
Dim i As Single

DamperSizeArray = Array(2, 2.5, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 24, 30, 36)
DamperCVArray = Array(115, 155, 250, 420, 660, 980, 1720, 2700, 3880, 4700, 6200, 7900, 9800, 14300, 23000, 33700)
i = 0

Do While i < 16
    If DamperSize = DamperSizeArray(i) Then
        CADamperCv = DamperCVArray(i)
        Exit Do
    Else: i = i + 1
    End If
Loop

End Function
Function CADamperTorque(DamperSize As Single)
Dim DamperSizeArray As Variant
Dim DamperTorqueArray As Variant
Dim i As Single

DamperSizeArray = Array(2, 2.5, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 24, 30, 36)
DamperTorqueArray = Array(11, 11, 11, 11, 11, 11, 11, 18, 34, 42, 55, 94, 112, 213, 336, 594)
i = 0

Do While i < 16
    If DamperSize = DamperSizeArray(i) Then
        CADamperTorque = DamperTorqueArray(i)
        Exit Do
    Else: i = i + 1
    End If
Loop

End Function
Function ExpTankExpansionLineSize(Duty As String)

'expansion volume ft3/h
'Assumed Expansion Line Size, in
'Expansion line ID-Calculated, in
'Selected Discharge and Vent Line Size, NPS inches
    'Below is Information Only
    'DutyMMBtuhr = Array(0.5, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10, 12.5, 17.5, 20, 25, 30, 35, 40, 50, 60)
    'EXPVOLArray = Array(10.87, 21.73, 32.6, 43.47, 54.34, 65.2, 86.94, 108.67, 130.41, 173.88, 217.34, 271.68, 380.35, 434.69, 543.36, 652.03, 760.7, 869.38, 1086.72, 1304.06)
    'ExpansionlineinreferencefromDIN4754 = Array(0.75, 0.75, 0.75, 1, 1.25, 1.25, 1.5, 1.5, 1.5, 2, 2, 2, 2, 2.5, 2.5, 2.5, 2.5, 3, 3, 4)
    'NominalSizeofDischargeandVentPipesreferencefromDIN4754 = Array(1, 1, 1, 1.25, 1.5, 1.5, 2, 2, 2, 2.5, 2.5, 2.5, 2.5, 3, 3, 3, 3, 4, 4, 6)
    'AssumedExpansionLineSizeArray = Array(1, 1, 1, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 4)
    'SelectedDischargeandVentLineSizeArray = Array(1, 1, 1, 2, 2, 2, 2, 2, 2, 3, 3, 3, 3, 3, 3, 3, 3, 4, 4, 6)
    'If expansion line calculation needed use below equations
    'Dim TempArray, FluidThermalConductivityArray As Variant
    'ExpVol = ExpansionVolume(Duty)
    'ExpVollbperh = ExpVol * FluidDensity(Fluid, TempDesign)
    'LinearVelocity = ExpVollbperh / 3600 / FluidDensity(Fluid, TempDesign) / (Area((PipeID(PipeOD(AssumedExpLineSize), PipeWall(AssumedExpLineSize, Schedule)) / 12)))
    'ExplineCalc = Sqrt(4 * (ExpVol / 3600) / (pi() * LinearVelocity)) * 12


Dim ExpansionlineIDCalculatedArray As Variant
Dim DutyArray As Variant
Dim i As Integer

    DutyArray = Array(0.5, 1, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10, 12.5, 15, 17.5, 20, 25, 30, 35, 40, 50, 60)
    ExpansionlineIDCalculatedArray = Array(1.049, 1.049, 1.049, 2.067, 2.067, 2.067, 2.067, 2.067, 2.067, 2.067, 2.067, 2.067, 2.067, 2.067, 3.068, 3.068, 3.068, 3.068, 3.068, 3.068, 4.026)
    i = 0

Do While i < 20

 If Duty = DutyArray(i) Then
    ExpTankExpansionLineSize = ExpansionlineIDCalculatedArray(i)
    Exit Do
    
  Else: i = i + 1
 End If
Loop

End Function
Function ExpansionTankID(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[ID]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpansionTankID = Recordset(0)
    Else: ExpansionTankID = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpansionTankStraightLength(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[Length]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpansionTankStraightLength = Recordset(0)
    Else: ExpansionTankStraightLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpansionTankLevelSwitchHeight(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[Height of LS]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpansionTankLevelSwitchHeight = Recordset(0)
    Else: ExpansionTankLevelSwitchHeight = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankLineID(VentLineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Piping Lookup].[Selected Expansion line Size, in]"
'define table
SQL = SQL & "FROM [Expansion Tank Piping Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Piping Lookup].[Expansion line(in) reference from DIN4754]=" & VentLineSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankLineID = Recordset(0)
    Else: ExpTankLineID = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpVentDischargeID(ExpLineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Piping Lookup].[Selected Discharge and Vent Line Size, in]"
'define table
SQL = SQL & "FROM [Expansion Tank Piping Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Piping Lookup].[Selected Expansion Line Size, in]=" & ExpLineSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpVentDischargeID = Recordset(0)
    Else: ExpVentDischargeID = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankC1(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[C1, Inspection Port/Hand Hole]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankC1 = Recordset(0)
    Else: ExpTankC1 = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankC2(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[C2, Exp/Vent PSV Connection]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankC2 = Recordset(0)
    Else: ExpTankC2 = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankC3(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[C3, Inert Gas Blanket Inlet]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankC3 = Recordset(0)
    Else: ExpTankC3 = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankC4(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[C4, Exp Tank Liquid Connection]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankC4 = Recordset(0)
    Else: ExpTankC4 = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankC6(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[C6, Low Low Liquid level Switch Connection]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankC6 = Recordset(0)
    Else: ExpTankC6 = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankC7(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[C7, Level Gauge Connections]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankC7 = Recordset(0)
    Else: ExpTankC7 = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ExpTankC8(TankSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[C8, Man Way]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]=" & TankSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ExpTankC8 = Recordset(0)
    Else: ExpTankC8 = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function RecDrainTankSize(SystemDesignVol As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[Tank Size]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]>" & SystemDesignVol

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    RecDrainTankSize = Recordset(0)
    Else: RecDrainTankSize = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function RecExpTankSize(SafeExpVolume As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank Lookup].[Tank Size]"
'define table
SQL = SQL & "FROM [Expansion Tank Lookup]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank Lookup].[Tank Size]>" & SafeExpVolume

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    RecExpTankSize = Recordset(0)
    Else: RecExpTankSize = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemVolume(InCoilDia As Single, OutCoilDia As Single, InCoilLength As Single, OutCoilLength As Single, numberTubes As Integer, PipingLineSize As Single, PipingLength As Single, UserVolume As Single, PipeSCH As String)

Dim TotalPipingVolume As Single, CoilVol1 As Single, CoilVol2 As Single
Dim CoilVol As Single, SystemDesignVolume As Single

Dim NumberReturns As Integer, ReturnRadius As Single

'MsgBox SystemDesignVolume
NumberReturns = 0
ReturnRadius = 0

'Coil Volume in Gallon
CoilVol1 = CoilVolume(InCoilDia, InCoilLength, numberTubes, NumberReturns, ReturnRadius)
CoilVol2 = CoilVolume(OutCoilDia, OutCoilLength, numberTubes, NumberReturns, ReturnRadius)
CoilVol = (CoilVol1 + CoilVol2) * 7.48052

TotalPipingVolume = PipingVolume(PipingLineSize, PipingLength, PipeSCH)
SystemVolume = TotalPipingVolume + CoilVol + UserVolume

End Function
Function PipingVolume(PipeSize As Single, length As Single, PipeSCH As String)
 
Dim PipeDi As Single, Areapipe As Single
 
PipeDi = PipeID(PipeOD(PipeSize), PipeWall((PipeSize), PipeSCH))
Areapipe = Area(PipeDi)
PipingVolume = Areapipe * (length * 12 / 1728) * 7.48
 
End Function
Function ExpansionVolume(Duty As String)
'Calculates expansion volume based on absorbed duty, in ft3/hr
Dim k As Single
k = 21.734
ExpansionVolume = k * Duty
End Function

Function GasViscosity(GasComp As Variant, temp As Single, Press As Single)
Dim arrVals  As Variant
Dim GasViscosityTemp, GasPerc, Visc As Single
Dim GasName As String
Dim i As Integer

For i = 0 To 100
    arrVals = Split(GasComp, ";")
    GasName = arrVals(i)
    On Error GoTo Err1
    If GasName = "" Then
        Exit Function
    End If
    GasPerc = CSng(arrVals(i + 1))
    Visc = Viscosity(GasName, "TP", "E", temp, Press) * 1488
    GasViscosityTemp = Visc * GasPerc
    If i = 0 Then
        GasViscosity = GasViscosityTemp
    End If
    If i > 1 Then
        GasViscosity = GasViscosity + GasViscosityTemp
    End If
    
i = i + 1
Next i
Err1:

End Function

Function GasDensity(GasComp As Variant, temp As Single, Press As Single)
Dim arrVals  As Variant
Dim GasDensityTemp, GasPerc, Dens As Single
Dim GasName As String
Dim i As Integer

For i = 0 To 100
    arrVals = Split(GasComp, ";")
    GasName = arrVals(i)
    On Error GoTo Err1
    If GasName = "" Then
        Exit Function
    End If
    GasPerc = CSng(arrVals(i + 1))
    Dens = Density(GasName, "TP", "E", temp, Press)
    GasDensityTemp = Dens * GasPerc
    If i = 0 Then
        GasDensity = GasDensityTemp
    End If
    If i > 1 Then
        GasDensity = GasDensity + GasDensityTemp
    End If
    
i = i + 1
Next i
Err1:

End Function
Function GasSpecificHeat(GasComp As Variant, temp As Single, Press As Single)
Dim arrVals  As Variant
Dim GasSpecificHeatTemp, GasPerc, SH As Single
Dim GasName As String
Dim i As Integer

For i = 0 To 100
    arrVals = Split(GasComp, ";")
    GasName = arrVals(i)
    On Error GoTo Err1
    If GasName = "" Then
        Exit Function
    End If
    GasPerc = CSng(arrVals(i + 1))
    SH = IsobaricHeatCapacity(GasName, "TP", "E", temp, Press)
    GasSpecificHeatTemp = SH * GasPerc
    If i = 0 Then
        GasSpecificHeat = GasSpecificHeatTemp
    End If
    If i > 1 Then
        GasSpecificHeat = GasSpecificHeat + GasSpecificHeatTemp
    End If
    
i = i + 1
Next i
Err1:

End Function
Function GasThermalConductivity(GasComp As Variant, temp As Single, Press As Single)
Dim arrVals  As Variant
Dim GasThermalConductivityTemp, GasPerc, tc As Single
Dim GasName As String
Dim i As Integer

For i = 0 To 100
    arrVals = Split(GasComp, ";")
    GasName = arrVals(i)
    On Error GoTo Err1
    If GasName = "" Then
        Exit Function
    End If
    GasPerc = CSng(arrVals(i + 1))
    tc = ThermalConductivity(GasName, "TP", "E", temp, Press)
    GasThermalConductivityTemp = tc * GasPerc
    If i = 0 Then
        GasThermalConductivity = GasThermalConductivityTemp
    End If
    If i > 1 Then
        GasThermalConductivity = GasThermalConductivity + GasThermalConductivityTemp
    End If
    
i = i + 1
Next i
Err1:

End Function
Function HelixLength(PipeOD As Single, PitchDiameter As Single, n)

HelixLength = ((PipeOD / 12) ^ 2 + (pi * PitchDiameter / 12) ^ 2) ^ 0.5 * n


End Function

Function Concat(rng As range) As String

Dim i As Integer
       
    For i = 1 To rng.Cells.Count / 2
        If rng(i, 2) <> 0 Then
            Concat = Concat & rng(i, 1).Value & ";" & rng(i, 2).Value / 100 & ";"
        End If
    Next i

End Function
Function CvRequiredGas(Flow As Single, p1 As Single, y As Single, M As Single, x As Single, t1 As Single, z As Single)
Dim N7 As Single
N7 = 7320
p1 = p1 + 14.7 'change to psia
t1 = t1 + 459.67 'change to Rankine


CvRequiredGas = Flow / (N7 * p1 * y * (x / (M * t1 * z)) ^ 0.5)



End Function
Function CvRequiredLiquid(MassFlow As Single, dP As Single, SpecWeight As Single)
Dim N6 As Single
N6 = 63.3

CvRequiredLiquid = MassFlow / (N6 * (dP * SpecWeight) ^ 0.5)

End Function
Function ControlValveDP(q As Single, p1 As Single, y As Single, M As Single, Cv As Single, t1 As Single, z As Single)
Dim N7 As Single
N7 = 1360
p1 = p1 + 14.7
t1 = t1 + 459.67

ControlValveDP = ((q / (Cv * N7 * p1 * y)) ^ 2) * M * t1 * z * p1

End Function
Function FluidMixture(CO2 As Single, H2O As Single, N2 As Single, O2 As Single, SO2 As Single, Methane As Single, Ethane As Single, Propane As Single, Butane As Single, H2S As Single, H2 As Single) As String

End Function
Function PartialVolumeHorzCyl(radius As Single, height As Single, length As Single)

PartialVolumeHorzCyl = (radius ^ 2 * WorksheetFunction.Acos((radius - height) / radius) - (radius - height) * (2 * radius * height - height ^ 2) ^ 0.5) * length + 2 * (pi * (height ^ 2 * (1.5 * (2 * radius) - height)) / 12)

PartialVolumeHorzCyl = PartialVolumeHorzCyl / 12 ^ 3 * 7.481 'convert to gallons
End Function
Function PartialVolumeVertCyl(radius As Single, height As Single, length As Single)
Dim EllipseRadius, StraightLength, EndVolume, StraightVolume, TopVolume As Single


EllipseRadius = radius / 2
StraightLength = length - 2 * EllipseRadius

EndVolume = pi * (radius) ^ 2 * (2 * EllipseRadius) / 3
StraightVolume = pi * (radius) ^ 2 * (height - 2 * EllipseRadius)
TopVolume = (pi * (radius) ^ 2 * (3 * (height - EllipseRadius - StraightLength) - EllipseRadius) * (height - EllipseRadius - StraightLength) ^ 2) / (3 * EllipseRadius ^ 2)


If height <= EllipseRadius Then
        PartialVolumeVertCyl = (pi * (radius) ^ 2 * (3 * EllipseRadius - height) * height ^ 2) / (3 * EllipseRadius ^ 2)
    ElseIf height >= EllipseRadius And height <= (EllipseRadius + StraightLength) Then
        PartialVolumeVertCyl = EndVolume + (pi * (radius) ^ 2 * (height - EllipseRadius))
    ElseIf height > (EllipseRadius + StraightLength) Then
        PartialVolumeVertCyl = EndVolume + StraightVolume + TopVolume
End If

End Function
Function Area(diameter)
Area = pi * diameter ^ 2 / 4
End Function
Function LinearInterp(Group1High As Single, Group1Low As Single, Group2High As Single, Group2Low As Single, Target As Single)
LinearInterp = (Target - Group1High) / (Group1Low - Group1High) * (Group2Low - Group2High) + Group2High

End Function
Function EstFanHp(CFM As Single, inWC As Single, Efficiency As Single)
EstFanHp = CFM * inWC / (6356 * Efficiency)

End Function

Function EstPumpHp(GPM As Single, PSI As Single, Efficiency As Single)
EstPumpHp = GPM * PSI / (1713 * Efficiency)

End Function
Function Talk(Txt As String)

Application.Speech.Speak (Txt)
Talk = Txt
End Function

Function TestViscosity(Gas As String, temp As Single, Press As Single)
TestViscosity = Viscosity(Gas, "TP", "e", temp, Press + 14.7)

End Function



Function TeeWeight(TeeOD As Single, TeeSchedule As String) As Single
Dim TeeODArray, TeeWeightArray As Variant
Dim TeeScheduleArray As Variant
Dim i, n As Integer
Dim SCH40, SCH80, SCH160, NA As Single


TeeODArray = Array(1.9, 2.375, 2.875, 3.5, 4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 24)
TeeScheduleArray = Array(SCH40, SCH80, SCH160)

Select Case TeeSchedule

    Case "SCH40"
        TeeWeightArray = Array(1.9, 3.2, 5.8, 7.2, 8.5, 9.5, 12.7, 20.8, 33.1, 56.5, 90.9, 165, 210, 275, 355, 510)
    Case "SCH80"
        TeeWeightArray = Array(2.3, 3.9, 6.5, 9, 12.2, 16.2, 26.6, 41.8, 76.2, 120, 175, 240, 290, 360, 450, 630)
    Case "SCH160"
        TeeWeightArray = Array(2.88, 4.88, 8.13, 11.25, NA, 20.25, 33.25, 52.25, 95.25, 143.75, 211.25, 296.25, 355, 441.25, 552.5, 781.25)
End Select
n = 0
Do While i < 15
    If TeeOD = TeeODArray(i) Then
       TeeWeight = TeeWeightArray(i)
        
    End If
    i = i + 1
    TeeSchedule = TeeScheduleArray(n)
Loop
n = n + 1

End Function


Function LineSize(Flow As Single)
Dim FlowRangeLow, FlowRangeHigh, LineSizeArray As Variant
Dim i As Integer

FlowRangeLow = Array(0, 9.001, 17.001, 27.001, 47.001, 63.001, 105.001, 150.001, 230.001, 400.001, 900.001, 1550.001, 2500.001, 3500.001, 4250.001, 5500.001, 7000.001)
FlowRangeHigh = Array(9, 17, 27, 47, 63, 105, 150, 230, 400, 900, 1550, 2500, 3500, 4250, 5500, 7000, 8650)
LineSizeArray = Array(0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 4, 6, 8, 10, 12, 14, 16, 18, 20)

For i = 0 To 16
    If Flow >= FlowRangeLow(i) And Flow <= FlowRangeHigh(i) Then
        LineSize = LineSizeArray(i)
        Exit Function
    End If
Next i
    
End Function
Function PipeDia(Flow As Single, FlowUnits As String, VelocityDesired As Single, Schedule As String)
Dim PipeNomArray, PipeODArray, PipeWallArray As Variant
Dim Area, RequiredDia, PipeNom, PipeOD As Single
Dim i As Integer
PipeNomArray = Array(0.125, 0.25, 0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 42)
PipeODArray = Array(0.405, 0.54, 0.675, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 2.875, 3.5, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 42)


Select Case Schedule
        Case 0.25
            PipeWallArray = Array(0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25)
        Case "SCH5S"
            PipeWallArray = Array("NA", "NA", "NA", 0.065, 0.065, 0.065, 0.065, 0.065, 0.065, 0.083, 0.083, 0.083, 0.083, 0.109, 0.109, 0.109, 0.134, 0.156, 0.156, 0.165, 0.165, 0.188, 0.188, 0.218, "NA", "NA", 0.25, "NA", "NA", "NA", "NA")
        Case "SCH10S"
            PipeWallArray = Array(0.049, 0.065, 0.065, 0.083, 0.083, 0.109, 0.109, 0.109, 0.109, 0.12, 0.12, 0.12, 0.12, 0.134, 0.134, 0.148, 0.165, 0.18, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, "NA", "NA", 0.312, "NA", "NA", "NA", "NA")
        Case "SCH10"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.188, 0.188, 0.188, 0.218, 0.218, 0.25, 0.312, 0.312, 0.312, 0.312, 0.312, 0.312, "NA")
        Case "SCH20"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.25, 0.25, 0.25, 0.312, 0.312, 0.312, 0.375, 0.375, 0.375, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.375)
        Case "SCH30"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.277, 0.307, 0.33, 0.375, 0.375, 0.438, 0.5, 0.5, 0.562, , 0.625, 0.625, 0.625, 0.625, 0.625, "NA")
        Case "SCH40"
            PipeWallArray = Array(0.068, 0.088, 0.091, 0.109, 0.113, 0.133, 0.14, 0.145, 0.154, 0.203, 0.216, 0.226, 0.237, 0.258, 0.28, 0.322, 0.365, 0.406, 0.438, 0.5, 0.562, 0.594, "NA", 0.688, "NA", "NA", 0.688, 0.688, 0.688, 0.75, "NA")
        Case "STD"
            PipeWallArray = Array(0.068, 0.088, 0.091, 0.109, 0.113, 0.133, 0.145, 0.145, 0.154, 0.203, 0.216, 0.226, 0.237, 0.258, 0.28, 0.322, 0.365, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375)
        Case "SCH60"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.406, 0.5, 0.562, 0.594, 0.656, 0.75, 0.812, 0.875, 0.969, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCHXS"
            PipeWallArray = Array(0.095, 0.119, 0.126, 0.147, 0.154, 0.179, 0.191, 0.2, 0.276, 0.276, 0.3, 0.318, 0.337, 0.375, 0.432, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5)
        Case "SCH80"
            PipeWallArray = Array(0.095, 0.119, 0.126, 0.147, 0.154, 0.179, 0.191, 0.2, 0.218, 0.276, 0.3, 0.318, 0.337, 0.375, 0.432, 0.5, 0.594, 0.688, 0.75, 0.844, 0.938, 1.031, 1.125, 0.218, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH100"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.594, 0.719, 0.844, 0.938, 1.031, 1.156, 1.281, 1.375, 1.531, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH120"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.438, 0.5, 0.562, 0.719, 0.844, 1, 1.094, 1.219, 1.375, 1.5, 1.625, 1.812, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH140"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.812, 1, 1.125, 1.25, 1.438, 1.562, 1.75, 1.875, 2.062, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH160"
            PipeWallArray = Array("NA", "NA", "NA", 0.188, 0.219, 0.25, 0.25, 0.281, 0.344, 0.375, 0.438, , 0.531, 0.625, 0.719, 0.906, 1.125, 1.312, 1.406, 1.594, 1.781, 1.969, 2.125, 2.344, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCHXXS"
            PipeWallArray = Array("NA", "NA", "NA", 0.294, 0.308, 0.358, 0.382, 0.4, 0.436, 0.552, 0.6, , 0.674, 0.75, 0.864, 0.875, 1, 1, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA")
End Select


Select Case FlowUnits
    Case "CFH"
        Area = Flow / 3600 / VelocityDesired * 144 'in^2
    Case "CFM"
        Area = Flow / 60 / VelocityDesired * 144 'in^2
    Case "GPM"
        Area = Flow * 0.133681 / 60 / VelocityDesired * 144 'in^2
End Select

RequiredDia = (4 * Area / pi) ^ 0.5 'inches

i = 0
Do While i < 29
    If RequiredDia < (PipeODArray(i) - 2 * PipeWallArray(i)) Then
        PipeNom = PipeNomArray(i - 1)
        Exit Do
    End If
    i = i + 1
Loop

PipeDia = PipeNom

End Function
Function BalanceFlow(DPInner As Single, DPOuter As Single, OuterFlow As Single)
Dim InnerFlow As Single
Dim Diff As Single

'InnerFlow = 50000

Diff = (DPOuter - DPInner)

If Abs(Diff) > 0.01 Then
       
    If DPOuter > DPInner Then
            InnerFlow = InnerFlow + 0.1
        ElseIf DPOuter < DPInner Then
            OuterFlow = OuterFlow + 0.1
    End If
Else: Exit Function
End If
BalanceFlow = InnerFlow

End Function
Option Explicit
 
Function LookUpHigh(LookupValue As Variant, LookupRange As range, ResultRange As range) As Variant
    Dim tmpVal As Variant
    Dim i As Long, myMax As Long
    On Error Resume Next
    tmpVal = WorksheetFunction.Match(LookupValue, LookupRange, 0)
    If Err <> 0 Then
        On Error GoTo 0
        tmpVal = WorksheetFunction.Match(LookupValue, LookupRange, 1)
        myMax = WorksheetFunction.CountA(LookupRange)
        For i = tmpVal To myMax
            If LookupRange(i).Value > LookupValue Then
                LookUpHigh = ResultRange(i)
                Exit Function
            End If
        Next i
    Else
        LookUpHigh = ResultRange(tmpVal)
    End If
End Function


Function LiquidOrificeSize(PipeID As Single, PressureDrop As Single, FluidDensity As Single, Cd As Single, Flow As Single) As Single
Dim a As Single

Flow = Flow * 0.00222801 'convert flow from GPM to ft3/sec
PressureDrop = PressureDrop * 144 'convert from psi to lb/ft2
PipeID = PipeID / 12 ' convert to ft

a = Flow * 4 / (pi * Cd * (2 * g * PressureDrop / FluidDensity) ^ 0.5)

LiquidOrificeSize = (a ^ 2 * PipeID ^ 4 / (PipeID ^ 4 + a ^ 2)) ^ 0.25 'ft
LiquidOrificeSize = LiquidOrificeSize * 12 'inches

End Function

Function LiquidOrificeDP(PipeID As Single, OrificeID As Single, FluidDensity As Single, Cd As Single, Flow As Single) As Single

Dim beta, Ao As Single

Flow = Flow * 0.00222801 'convert flow from GPM to ft3/sec
beta = OrificeID / PipeID 'ratio of diameters
Ao = pi * OrificeID ^ 2 / 4 / 144 'orifice area in ft2

LiquidOrificeDP = FluidDensity / (2 * g) * (Flow * (1 - beta ^ 4) ^ 0.5 / (Ao * Cd)) ^ 2 / 144

End Function
Function LiquidOrificeFlow(PipeID As Single, OrificeID As Single, FluidDensity As Single, Cd As Single, PressureDrop As Single) As Single

Dim beta, Ao As Single

beta = OrificeID / PipeID 'ratio of diameters
Ao = pi * OrificeID ^ 2 / 4 / 144 'orifice area in ft2
PressureDrop = PressureDrop * 144 'convert from psi to lb/ft2

LiquidOrificeFlow = Ao * Cd / (1 - beta ^ 4) ^ 0.5 * (2 * g * PressureDrop / FluidDensity) ^ 0.5  'flow in ft3/sec

LiquidOrificeFlow = LiquidOrificeFlow / 0.00222801 ' converted to GPM

End Function
Function CoilVolume(diameter As Single, length As Single, numberTubes As Integer, NumberReturns As Integer, ReturnRadius As Single) As Single

Dim PipeArea, PipeVolume, ReturnVolume, ReturnLength As Single

PipeArea = Area(diameter)

ReturnLength = ReturnRadius * pi / 12

PipeVolume = PipeArea / 144 * length * numberTubes
ReturnVolume = PipeArea / 144 * ReturnLength * NumberReturns

CoilVolume = PipeVolume + ReturnVolume

End Function

Public Function hco(Re As Single, k As Single, d As Single) As Single
Dim c, n As Single

If Re > 40 And Re <= 4000 Then
        c = 0.615
        n = 0.466
    ElseIf Re > 4000 And Re <= 40000 Then
        c = 0.174
        n = 0.618
    ElseIf Re > 40000 Then
        c = 0.0239
        n = 0.805
End If

hco = k / (d / 12) * c * Re ^ n
    
End Function

Function LossCoefficientBend(Reynolds As Single, BendRadius As Single, ID As Single, Angle As Single, frictionFactor As Single) As Single

Dim lambda, x As Single

x = Reynolds * Sqr(ID / (2 * BendRadius))

If x > 50 And x < 600 Then
        
        lambda = 20 / Reynolds ^ 0.65 * (ID / (2 * BendRadius)) ^ 0.175
    
    ElseIf x > 600 And x < 1400 Then
        
        lambda = 10.4 / Reynolds ^ 0.55 * (ID / (2 * BendRadius)) ^ 0.225
    
    ElseIf x > 1400 Then
        
        lambda = 5 / Reynolds ^ 0.45 * (ID / (2 * BendRadius)) ^ 0.275
    
End If

LossCoefficientBend = 0.0175 * lambda * BendRadius / ID * Angle / frictionFactor


End Function

Function LossCoefficientDivergingTee(FlowBranch As Single, FlowHeader As Single, AreaBranch As Single, AreaHeader As Single, frictionFactor As Single) As Single
Dim a As Single
Dim GammaCS, gamma As Single
Dim VelocityBranch, VelocityHeader As Single

a = 0.9 * (1 - FlowBranch / FlowHeader)
VelocityBranch = FlowBranch / AreaBranch
VelocityHeader = FlowHeader / AreaHeader

GammaCS = a * (1 + 0.3 * (VelocityBranch / VelocityHeader) ^ 2)
gamma = GammaCS / (VelocityBranch / VelocityHeader) ^ 2

LossCoefficientDivergingTee = gamma

End Function

Function EquivalentDiameter(diameter As Single, splits As Integer)

EquivalentDiameter = Sqr(splits * diameter ^ 2)

End Function

Function hci(Di, Re, Pr, ki, visc_i, visc_w) As Single
'
'   Di - inside tube diameter, inches
'   Gi - inside mass flow velocity, lb/hr-ft2
'   cpi - specific heat of inside fluid, Btu/lb-F
'   ki - thermal conductivity of inside fluid, Btu/hr-ft-F
'   visc_i - viscosity of inside fluid, lb/hr-ft
'   visc_w - viscosity evaluated at wall Temperature
'
Dim k As Single
Dim f As Single

'Viscosity correction factor for high Prandtl number fluids
'or laminar/transition flow
'
'k = (visc_i / visc_w) ^ (0.14)
k = (visc_i / visc_w) ^ (0.11)
'
If Re > 10000 And Pr < 100 Then
'   turbulent flow
    hci = 0.027 * (Re ^ 0.8 * Pr ^ (1 / 3)) * ki / (Di / 12)
ElseIf Re > 10000 And Pr > 100 Then
'   high Prandtl No. fluid
    hci = 0.027 * Re ^ 0.8 * Pr ^ (1 / 3) * ki / (Di / 12) * k
ElseIf Re < 10000 And Re > 2100 Then
'   transition flow
    hci = 0.027 * Re ^ 0.8 * Pr ^ (1 / 3) * ki / (Di / 12) * k
    'f = (0.79 * 2.3 * Log(Re) - 1.64) ^ -2
    'hci = ((f / 8) * (Re - 1000) * Pr / (1 + 12.7 * (f / 8) ^ 0.5 * (Pr ^ (2 / 3) - 1))) * ki / (Di / 12) * k
    
ElseIf Re < 2100 Then
'   laminar flow
    hci = 1.86 * Re ^ (1 / 3) * Pr ^ (1 / 3) * k * ki / (Di / 12)
End If
'
End Function
Function Hci_nucleate_boiling(Viscosity_liquid As Single, LatentHeat As Single, Density_liquid As Single, Density_vapor As Single, Cp_liquid As Single, Tsurface As Single, Tsat As Single, Pr As Single, SurfaceConstant As Single, SurfaceTension As Single)
Hci_nucleate_boiling = Viscosity_liquid * LatentHeat * (g * (Density_liquid - Density_vapor) / SurfaceTension) ^ 0.5 * (Cp_liquid * (Tsurface - Tsat) / (LatentHeat * Pr * SurfaceConstant)) ^ 3
End Function
Function Hci_film_boiling(Re As Single, Density_vapor As Single, Density_liquid As Single, ThermCond_vapor As Single, Viscosity_vapor As Single)

Hci_film_boiling = 0.002 * Re ^ 0.6 * ((g * 3600 * Density_vapor * (Density_liquid - Density_vapor) * ThermCond_vapor ^ 3 / Viscosity_vapor ^ 2) ^ (1 / 3))

End Function
Function IdealGasDensity(MolecularWeight As Single, Pressure As Single, Temperature As Single)

Dim Patm, TempRankine As Single

Patm = 14.7
TempRankine = 459.67

IdealGasDensity = ((Patm + Pressure) * MolecularWeight * 144) / (R * (Temperature + TempRankine))

End Function

Function LossCoefficientConvergingTee(FlowBranch As Single, FlowHeader As Single, AreaBranch As Single, AreaHeader As Single, frictionFactor As Single) As Single
Dim a As Single
Dim GammaCS, gamma As Single

a = 0.9 * (1 - FlowBranch / FlowHeader)

GammaCS = a * (1 + ((FlowBranch / FlowHeader) * (AreaHeader / AreaBranch)) ^ 2 - 2 * (1 - FlowBranch / FlowHeader) ^ 2)
gamma = GammaCS / ((FlowBranch * AreaHeader) / (FlowHeader * AreaBranch)) ^ 2

LossCoefficientConvergingTee = gamma
End Function

Function Overall_U(hci As Single, hco As Single, TubeResistance As Single, FoulingFactor As Single, OD As Single, ID As Single)

Dim OutsideRadius, InsideRadius As Single
Dim InsideResistance, OutsideResistance As Single

'Radii must be in feet
OutsideRadius = (OD / 2) / 12
InsideRadius = (ID / 2) / 12

OutsideResistance = 1 / hco
InsideResistance = (OutsideRadius / InsideRadius) * (1 / hci)

Overall_U = 1 / (OutsideResistance + InsideResistance + TubeResistance + FoulingFactor)

End Function

Function PipeID(PipeOD As Single, PipeWall As Single)

PipeID = PipeOD - 2 * PipeWall

End Function

Function PipeOD(PipeNom As Single)
Dim PipeNomArray, PipeODArray As Variant
Dim i As Integer

PipeNomArray = Array(0.125, 0.25, 0.375, 0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 42)
PipeODArray = Array(0.405, 0.54, 0.675, 0.84, 1.05, 1.315, 1.66, 1.9, 2.375, 2.875, 3.5, 4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 42)

Do While i < 30
    If PipeNom = PipeNomArray(i) Then
        PipeOD = PipeODArray(i)
        Exit Do
    End If
    i = i + 1
Loop

End Function

Function PipeSurfaceArea(PipeOD As Single, length As Single, Number As Single)

PipeSurfaceArea = pi * PipeOD / 12 * length * Number

End Function
Function PipeThermalConductivity(Material As String, Temperature As Single)
Dim MaterialArray, mArray, bArray As Variant
Dim M, b As Single
Dim i As Integer
MaterialArray = Array("SA-106-B", "SA-333-6", "SA-335-P11", "SA-335-P22", "SA-335-P5", "SA-312-TP304", "SA-312-TP316")
mArray = Array(-0.008, -0.008, 0.0025, 0.0025, 0.0025, 0.0048, 0.0052)
bArray = Array(31, 31, 13.8, 13.8, 13.8, 8.24, 6.96)

i = 0
Do While i < 7
    If Material = MaterialArray(i) Then
        M = mArray(i)
        b = bArray(i)
        Exit Do
    End If
    i = i + 1
Loop
       
PipeThermalConductivity = M * Temperature + b

End Function

Function PipeWall(PipeNom As Single, Schedule As String)

Dim PipeNomArray, PipeWallArray As Variant
Dim i As Integer

PipeNomArray = Array(0.125, 0.25, 0.375, 0.5, 0.75, 1, 1.25, 1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36, 42)
Select Case Schedule
        Case "SCH5S"
            PipeWallArray = Array("NA", "NA", "NA", 0.065, 0.065, 0.065, 0.065, 0.065, 0.065, 0.083, 0.083, 0.083, 0.083, 0.109, 0.109, 0.109, 0.134, 0.156, 0.156, 0.165, 0.165, 0.188, 0.188, 0.218, "NA", "NA", 0.25, "NA", "NA", "NA", "NA")
        Case "SCH10S"
            PipeWallArray = Array(0.049, 0.065, 0.065, 0.083, 0.083, 0.109, 0.109, 0.109, 0.109, 0.12, 0.12, 0.12, 0.12, 0.134, 0.134, 0.148, 0.165, 0.18, 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, "NA", "NA", 0.312, "NA", "NA", "NA", "NA")
        Case "SCH10"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.25, 0.25, 0.25, 0.25, 0.25, 0.25, 0.312, 0.312, 0.312, 0.312, 0.312, 0.312, "NA")
        Case "SCH20"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.25, 0.25, 0.25, 0.312, 0.312, 0.312, 0.375, 0.375, 0.375, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.375)
        Case "SCH30"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.277, 0.307, 0.33, 0.375, 0.375, 0.438, 0.5, 0.5, 0.562, "NA", 0.625, 0.625, 0.625, 0.625, 0.625, "NA")
        Case "SCH40"
            PipeWallArray = Array(0.068, 0.088, 0.091, 0.109, 0.113, 0.133, 0.14, 0.145, 0.154, 0.203, 0.216, 0.226, 0.237, 0.258, 0.28, 0.322, 0.365, 0.406, 0.438, 0.5, 0.562, 0.594, "NA", 0.688, "NA", "NA", "NA", 0.688, 0.688, 0.75, "NA")
        Case "STD"
            PipeWallArray = Array(0.068, 0.088, 0.091, 0.109, 0.113, 0.133, 0.14, 0.145, 0.154, 0.203, 0.216, 0.226, 0.237, 0.258, 0.28, 0.322, 0.365, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, 0.375, "N/A")
        Case "SCH60"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.406, 0.5, 0.562, 0.594, 0.656, 0.75, 0.812, 0.875, 0.969, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCHXS"
            PipeWallArray = Array(0.095, 0.119, 0.126, 0.147, 0.154, 0.179, 0.191, 0.2, 0.218, 0.276, 0.3, 0.318, 0.337, 0.375, 0.432, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 0.5)
        Case "SCH80"
            PipeWallArray = Array(0.095, 0.119, 0.126, 0.147, 0.154, 0.179, 0.191, 0.2, 0.218, 0.276, 0.3, 0.318, 0.337, 0.375, 0.432, 0.5, 0.594, 0.688, 0.75, 0.844, 0.938, 1.031, 1.125, 1.218, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH100"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.594, 0.719, 0.844, 0.938, 1.031, 1.156, 1.281, 1.375, 1.531, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH120"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.438, 0.5, 0.562, 0.719, 0.844, 1, 1.094, 1.219, 1.375, 1.5, 1.625, 1.812, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH140"
            PipeWallArray = Array("NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", 0.812, 1, 1.125, 1.25, 1.438, 1.562, 1.75, 1.875, 2.062, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCH160"
            PipeWallArray = Array("NA", "NA", "NA", 0.188, 0.219, 0.25, 0.25, 0.281, 0.344, 0.375, 0.438, "NA", 0.531, 0.625, 0.719, 0.906, 1.125, 1.312, 1.406, 1.594, 1.781, 1.969, 2.125, 2.344, "NA", "NA", "NA", "NA", "NA", "NA", "NA")
        Case "SCHXXS"
            PipeWallArray = Array("NA", "NA", "NA", 0.294, 0.308, 0.358, 0.382, 0.4, 0.436, 0.552, 0.6, "NA", 0.674, 0.75, 0.864, 0.875, 1, 1, "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA", "NA")
            
End Select

Do While i < 30
    If PipeNom = PipeNomArray(i) Then
        PipeWall = PipeWallArray(i)
        Exit Do
    End If
    i = i + 1
Loop

End Function



Function LMTD(TSin, TSout, TTin, TTout, FlowConfig%) As Double
'
'   Function to calclulate Log Mean Temperature Difference
'   TSin - Temperature Shell side in
'   TSout -Temperature Shell side out
'   TTin - Temperature tube side in
'   TTout -Temperature tube side out
'
Dim delt1 As Single            'intermediate variable
Dim delt2 As Single            'intermediate variable
Dim ratio As Single
'
If FlowConfig = 1 Then
    delt2 = TSout - TTin
    delt1 = TSin - TTout
ElseIf FlowConfig = 2 Then
    delt2 = TSin - TTin
    delt1 = TSout - TTout
End If
'
ratio = delt2 / delt1
If ratio <= 0 Or delt1 = 0 Or delt2 = 0 Then
    'Error in temperature specification
 '   MsgBox "ERROR: invalid inlet/outlet temperature specification!" _
  ' & vbCrLf & "Check specified temperatures."
   LMTD = 1
Else
    LMTD = (delt2 - delt1) / Log(delt2 / delt1)
End If
'
If FlowConfig < 1 Or FlowConfig > 2 Then
    MsgBox "Error in flow configuration specification for LMTD calculation"
    LMTD = "  "
End If

End Function
Function FlangeRating(temp As Single, PClass As String, Material As String)
Dim k, i As Integer
Dim Mat, TempArray As Variant

Select Case (Material)
    Case "A105", "A515-70", "A516-70", "SA350-LF2", "SA-105"
        k = 1
    Case "1.2 - Carbon Steel - A350 LF3"
        k = 2
    Case "1.3 - Carbon Steel - A515 65,A515 65 (BS 1501 151/161 430)"
        k = 3
    Case "1.4 - Carbon Steel - A515 60, A516 60, A350 LF1"
        k = 4
    Case "1.5 - Alloy Steel - A182 F1"
        k = 5
    Case "1.7 - Alloy Steel - A182 F2"
        k = 6
    Case "1.9 - Alloy Steel - A182 F11, A182 F12 (BS 1501 621, BS 1503 620 440 / 621 460)", "SA-335-P11", "SA-182-F11"
        k = 7
    Case "1.11 - 1.10 in ANSI B16.5 - Alloy Steel - A182 F22 (BS 1501 622-515, BS 1503 622 490)", "SA-335-P22", "SA-182-F22"
        k = 8
    Case "1.13 - Alloy Steel - A182 F5, A182 F5a (BS 1503 625 590)"
        k = 9
    Case "1.14 - Alloy Steel - A182 F9", "SA-182-F9"
        k = 10
    Case "2.1 - Austenitic Steel - 304 (304 S31/S51)", "304SS", "SA-182-304"
        k = 11
    Case "2.2 - Austenitic Steel - 316 (316 S31/S51)", "316SS", "SA-182-316"
        k = 12
    Case "2.3 - Austenitic Steel - 304L / 316L (304 S11/316 S11)", "304LSS"
        k = 13
    Case "2.4 - Austenitic Steel - 321 (321 S31)", "SA-182-F321"
        k = 14
    Case "2.5 - Austenitic Steel - 347, 348 (347 S31/S51)", "347SS", "SA-182-347"
        k = 15
    Case "2.6 - Austenitic Steel - 309S"
        k = 16
    Case "2.7 - Austenitic Steel - 310 (310 S31)"
        k = 17
    Case "2.8 - Ferritic/Austenitic Steel - 318", "318SS"
        k = 18
End Select



Select Case (PClass)
    Case "150#"
        Select Case (k)
            Case 1
                Mat = Array(285, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 2
                Mat = Array(290, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 3
                Mat = Array(265, 250, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 4
                Mat = Array(235, 215, 210, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 5
                Mat = Array(265, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 6
                Mat = Array(290, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 7
                Mat = Array(290, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 8
                Mat = Array(290, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 9
                Mat = Array(290, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 10
                Mat = Array(290, 260, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 11
                Mat = Array(275, 230, 205, 190, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 12
                Mat = Array(275, 235, 215, 195, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 13
                Mat = Array(230, 195, 175, 160, 145, 140, 125, 110, 95, 80, 65, 0, 0, 0)
            Case 14
                Mat = Array(275, 245, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 15
                Mat = Array(275, 255, 230, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 16
                Mat = Array(260, 230, 220, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 17
                Mat = Array(260, 235, 220, 200, 170, 140, 125, 110, 95, 80, 65, 50, 35, 20)
            Case 18
                Mat = Array(290, 260, 230, 200, 170, 140, 125, 110, 95, 0, 0, 0, 0, 0)
        End Select
   
    Case "300#"
        
        Select Case (k)
            Case 1
                Mat = Array(740, 675, 655, 635, 600, 550, 535, 535, 505, 410, 270, 170, 105, 50, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 2
                Mat = Array(750, 750, 730, 705, 665, 605, 590, 570, 505, 410, 270, 170, 105, 50, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 3
                Mat = Array(695, 655, 640, 620, 585, 535, 525, 520, 475, 390, 270, 170, 105, 50, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 4
                Mat = Array(620, 560, 550, 530, 500, 455, 450, 450, 445, 370, 270, 170, 105, 50, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 5
                Mat = Array(695, 680, 655, 640, 620, 605, 590, 570, 530, 510, 485, 450, 280, 165, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 6
                Mat = Array(750, 750, 720, 695, 665, 605, 590, 570, 530, 510, 485, 450, 315, 200, 160, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 7
                Mat = Array(750, 750, 720, 695, 665, 605, 590, 570, 530, 510, 485, 450, 320, 215, 145, 95, 60, 40, 0, 0, 0, 0, 0, 0)
            Case 8
                Mat = Array(750, 750, 730, 705, 665, 605, 590, 570, 530, 510, 485, 450, 375, 260, 175, 110, 70, 40, 0, 0, 0, 0, 0, 0)
            Case 9
                Mat = Array(750, 745, 715, 705, 665, 605, 590, 570, 530, 510, 485, 370, 275, 200, 145, 100, 60, 35, 0, 0, 0, 0, 0, 0)
            Case 10
                Mat = Array(750, 750, 730, 705, 665, 605, 590, 570, 530, 510, 485, 450, 375, 255, 170, 115, 75, 50, 0, 0, 0, 0, 0, 0)
            Case 11
                Mat = Array(720, 600, 540, 495, 465, 435, 430, 425, 415, 405, 395, 390, 380, 320, 310, 255, 200, 155, 115, 85, 60, 50, 35, 25)
            Case 12
                Mat = Array(720, 620, 560, 515, 480, 450, 445, 430, 425, 420, 420, 415, 385, 350, 345, 305, 235, 185, 145, 115, 95, 75, 60, 40)
            Case 13
                Mat = Array(600, 505, 455, 415, 380, 360, 350, 345, 335, 330, 320, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 14
                Mat = Array(720, 645, 595, 550, 515, 485, 480, 465, 460, 450, 445, 440, 385, 355, 315, 270, 235, 185, 140, 110, 85, 65, 50, 40)
            Case 15
                Mat = Array(720, 660, 615, 575, 540, 515, 505, 495, 490, 485, 485, 450, 385, 365, 360, 325, 275, 170, 125, 95, 70, 55, 40, 35)
            Case 16
                Mat = Array(670, 605, 570, 535, 505, 480, 465, 455, 445, 435, 425, 415, 385, 335, 290, 225, 170, 130, 100, 80, 60, 45, 30, 25)
            Case 17
                Mat = Array(670, 605, 570, 535, 505, 480, 470, 455, 450, 435, 425, 420, 385, 345, 335, 260, 190, 135, 105, 75, 60, 45, 35, 25)
            Case 18
                Mat = Array(750, 720, 665, 615, 575, 555, 550, 540, 530, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
        End Select
    
    Case "400#"
        Select Case (k)
            Case 1
                Mat = Array(990, 900, 875, 845, 800, 730, 715, 710, 670, 550, 355, 230, 140, 70, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 2
                Mat = Array(1000, 1000, 970, 940, 885, 805, 785, 755, 670, 550, 355, 230, 140, 70, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 3
                Mat = Array(925, 875, 850, 825, 775, 710, 695, 690, 630, 520, 355, 230, 140, 70, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 4
                Mat = Array(825, 750, 730, 705, 665, 610, 600, 600, 590, 495, 355, 230, 140, 70, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 5
                Mat = Array(925, 905, 870, 855, 830, 805, 785, 755, 710, 675, 650, 600, 375, 220, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 6
                Mat = Array(1000, 1000, 965, 925, 885, 805, 785, 755, 710, 675, 650, 600, 420, 270, 210, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 7
                Mat = Array(1000, 1000, 965, 925, 885, 805, 785, 755, 710, 675, 650, 600, 425, 290, 190, 130, 80, 50, 0, 0, 0, 0, 0, 0)
            Case 8
                Mat = Array(1000, 1000, 970, 940, 885, 805, 785, 755, 710, 675, 650, 600, 505, 345, 235, 145, 90, 55, 0, 0, 0, 0, 0, 0)
            Case 9
                Mat = Array(1000, 995, 955, 940, 885, 805, 785, 755, 705, 675, 645, 495, 365, 265, 190, 135, 80, 45, 0, 0, 0, 0, 0, 0)
            Case 10
                Mat = Array(1000, 1000, 970, 940, 885, 805, 785, 755, 710, 675, 650, 600, 505, 340, 230, 150, 100, 70, 0, 0, 0, 0, 0, 0)
            Case 11
                Mat = Array(960, 800, 720, 660, 620, 580, 575, 565, 555, 540, 530, 520, 510, 430, 410, 345, 265, 205, 150, 115, 80, 65, 45, 35)
            Case 12
                Mat = Array(960, 825, 745, 685, 635, 600, 590, 580, 570, 565, 555, 555, 515, 465, 460, 405, 315, 245, 195, 155, 130, 100, 80, 55)
            Case 13
                Mat = Array(800, 675, 605, 550, 510, 480, 470, 460, 450, 440, 430, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 14
                Mat = Array(960, 860, 795, 735, 685, 650, 635, 620, 610, 600, 595, 590, 515, 475, 415, 360, 315, 245, 185, 145, 115, 85, 70, 50)
            Case 15
                Mat = Array(960, 880, 820, 765, 720, 685, 670, 660, 655, 650, 645, 600, 515, 485, 480, 430, 365, 230, 165, 125, 90, 75, 55, 45)
            Case 16
                Mat = Array(895, 805, 760, 710, 670, 635, 620, 610, 595, 580, 565, 555, 515, 450, 390, 300, 230, 175, 135, 105, 80, 60, 40, 30)
            Case 17
                Mat = Array(895, 810, 760, 715, 675, 640, 625, 610, 600, 580, 575, 555, 515, 460, 450, 345, 250, 185, 135, 100, 80, 60, 45, 35)
            Case 18
                Mat = Array(1000, 960, 885, 820, 770, 740, 735, 725, 710, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        End Select
    
    Case "600#"
        Select Case (k)
            Case 1
                Mat = Array(1480, 1350, 1315, 1270, 1200, 1095, 1075, 1065, 1010, 825, 535, 345, 205, 105, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 2
                Mat = Array(1500, 1500, 1455, 1410, 1330, 1210, 1175, 1135, 1010, 825, 535, 345, 205, 105, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 3
                Mat = Array(1390, 1315, 1275, 1235, 1165, 1065, 1045, 1035, 945, 780, 535, 345, 205, 105, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 4
                Mat = Array(1235, 1125, 1095, 1060, 995, 915, 895, 895, 885, 740, 535, 345, 205, 105, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 5
                Mat = Array(1390, 1360, 1305, 1280, 1245, 1210, 1175, 1135, 1065, 1015, 975, 900, 560, 330, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 6
                Mat = Array(1500, 1500, 1445, 1385, 1330, 1210, 1175, 1135, 1065, 1015, 975, 900, 630, 405, 315, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 7
                Mat = Array(1500, 1500, 1445, 1385, 1330, 1210, 1175, 1135, 1065, 1015, 975, 900, 640, 430, 290, 190, 125, 75, 0, 0, 0, 0, 0, 0)
            Case 8
                Mat = Array(1500, 1500, 1455, 1410, 1330, 1210, 1175, 1135, 1065, 1015, 975, 900, 755, 520, 350, 220, 135, 80, 0, 0, 0, 0, 0, 0)
            Case 9
                Mat = Array(1500, 1490, 1430, 1410, 1330, 1210, 1175, 1135, 1055, 1015, 965, 740, 550, 400, 290, 200, 125, 70, 0, 0, 0, 0, 0, 0)
            Case 10
                Mat = Array(1500, 1500, 1455, 1410, 1330, 1210, 1175, 1135, 1065, 1015, 975, 900, 755, 505, 345, 225, 150, 105, 0, 0, 0, 0, 0, 0)
            Case 11
                Mat = Array(1440, 1200, 1080, 995, 930, 875, 860, 850, 830, 805, 790, 780, 765, 640, 615, 515, 400, 310, 225, 170, 125, 95, 70, 55)
            Case 12
                Mat = Array(1440, 1240, 1120, 1025, 955, 900, 890, 870, 855, 845, 835, 830, 775, 700, 685, 610, 475, 370, 295, 235, 190, 150, 115, 85)
            Case 13
                Mat = Array(1200, 1015, 910, 825, 765, 720, 700, 685, 670, 660, 645, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 14
                Mat = Array(1440, 1290, 1190, 1105, 1030, 975, 955, 930, 915, 900, 895, 885, 775, 715, 625, 545, 465, 370, 280, 220, 170, 130, 105, 75)
            Case 15
                Mat = Array(1440, 1320, 1230, 1145, 1080, 1025, 1010, 990, 985, 975, 970, 900, 775, 725, 720, 645, 550, 345, 245, 185, 135, 110, 80, 70)
            Case 16
                Mat = Array(1345, 1210, 1140, 1065, 1010, 955, 930, 910, 895, 870, 850, 830, 775, 670, 585, 445, 345, 260, 200, 160, 115, 90, 60, 50)
            Case 17
                Mat = Array(1345, 1215, 1140, 1070, 1015, 960, 935, 910, 900, 875, 855, 835, 775, 685, 670, 520, 375, 275, 205, 150, 115, 90, 65, 50)
            Case 18
                Mat = Array(1500, 1440, 1330, 1230, 1150, 1115, 1100, 1085, 1065, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
                 
        End Select
    
    Case "900#"
        Select Case (k)
            Case 1
                Mat = Array(2220, 2025, 1970, 1900, 1795, 1640, 1610, 1600, 1510, 1235, 805, 515, 310, 155, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 2
                Mat = Array(2250, 2250, 2185, 2115, 1995, 1815, 1765, 1705, 1510, 1235, 805, 515, 310, 155, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 3
                Mat = Array(2085, 1970, 1915, 1850, 1745, 1600, 1570, 1555, 1420, 1175, 805, 515, 310, 155, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 4
                Mat = Array(1850, 1685, 1640, 1585, 1495, 1370, 1345, 1345, 1325, 1110, 805, 515, 310, 155, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 5
                Mat = Array(2085, 2035, 1955, 1920, 1865, 1815, 1765, 1705, 1595, 1525, 1460, 1350, 845, 495, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 6
                Mat = Array(2250, 2250, 2165, 2080, 1995, 1815, 1765, 1705, 1595, 1525, 1460, 1350, 945, 605, 475, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 7
                Mat = Array(2250, 2250, 2165, 2080, 1995, 1815, 1765, 1705, 1595, 1525, 1460, 1350, 955, 650, 430, 290, 185, 115, 0, 0, 0, 0, 0, 0)
            Case 8
                Mat = Array(2250, 2250, 2185, 2115, 1995, 1815, 1765, 1705, 1595, 1525, 1460, 1350, 1130, 780, 525, 330, 205, 125, 0, 0, 0, 0, 0, 0)
            Case 9
                Mat = Array(2250, 2235, 2150, 2115, 1995, 1815, 1765, 1705, 1585, 1525, 1450, 1110, 825, 595, 430, 300, 185, 105, 0, 0, 0, 0, 0, 0)
            Case 10
                Mat = Array(2250, 2250, 2185, 2115, 1995, 1815, 1765, 1705, 1595, 1525, 1460, 1350, 1130, 760, 515, 340, 225, 155, 0, 0, 0, 0, 0, 0)
            Case 11
                Mat = Array(2160, 1800, 1620, 1490, 1395, 1310, 1290, 1275, 1245, 1210, 1190, 1165, 1145, 965, 925, 770, 595, 465, 340, 255, 185, 145, 105, 80)
            Case 12
                Mat = Array(2160, 1860, 1680, 1540, 1435, 1355, 1330, 1305, 1280, 1265, 1255, 1245, 1160, 1050, 1030, 915, 710, 555, 440, 350, 290, 225, 175, 125)
            Case 13
                Mat = Array(1800, 1520, 1360, 1240, 1145, 1080, 1050, 1030, 1010, 985, 965, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 14
                Mat = Array(2160, 1935, 1785, 1655, 1545, 1460, 1435, 1395, 1375, 1355, 1340, 1325, 1160, 1070, 940, 815, 710, 555, 420, 330, 255, 195, 155, 115)
            Case 15
                Mat = Array(2160, 1980, 1845, 1720, 1620, 1540, 1510, 1485, 1475, 1460, 1455, 1350, 1160, 1090, 1080, 965, 825, 515, 370, 280, 205, 165, 125, 105)
            Case 16
                Mat = Array(2015, 1815, 1705, 1600, 1510, 1435, 1395, 1370, 1340, 1305, 1275, 1245, 1160, 1010, 875, 670, 515, 390, 300, 235, 175, 135, 95, 70)
            Case 17
                Mat = Array(2015, 1820, 1705, 1605, 1520, 1440, 1405, 1370, 1345, 1310, 1280, 1255, 1160, 1030, 1010, 780, 565, 410, 310, 225, 175, 135, 100, 75)
            Case 18
                Mat = Array(2250, 2160, 1995, 1845, 1730, 1670, 1650, 1625, 1595, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        End Select
    
    Case "1500#"
        Select Case (k)
            Case 1
                Mat = Array(3705, 3375, 3280, 3170, 2995, 2735, 2685, 2665, 2520, 2060, 1340, 860, 515, 260, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 2
                Mat = Array(3750, 3750, 3640, 3530, 3325, 3025, 2940, 2840, 2520, 2060, 1340, 860, 515, 260, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 3
                Mat = Array(3470, 3280, 3190, 3085, 2910, 2665, 2615, 2590, 2365, 1955, 1340, 860, 515, 260, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 4
                Mat = Array(3085, 2810, 2735, 2645, 2490, 2285, 2245, 2245, 2210, 1850, 1340, 860, 515, 260, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 5
                Mat = Array(3470, 3395, 3260, 3200, 3105, 3025, 2940, 2840, 2660, 2540, 2435, 2245, 1405, 825, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 6
                Mat = Array(3750, 3750, 3610, 3465, 3325, 3025, 2940, 2840, 2660, 2540, 2435, 2245, 1575, 1010, 790, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 7
                Mat = Array(3750, 3750, 3610, 3465, 3325, 3025, 2940, 2840, 2660, 2540, 2435, 2245, 1595, 1080, 720, 480, 310, 190, 0, 0, 0, 0, 0, 0)
            Case 8
                Mat = Array(3750, 3750, 3640, 3530, 3325, 3025, 2940, 2840, 2660, 2540, 2435, 2245, 1885, 1305, 875, 550, 345, 205, 0, 0, 0, 0, 0, 0)
            Case 9
                Mat = Array(3750, 3725, 3580, 3530, 3325, 3025, 2940, 2840, 2640, 2540, 2415, 1850, 1370, 995, 720, 495, 310, 170, 0, 0, 0, 0, 0, 0)
            Case 10
                Mat = Array(3750, 3750, 3640, 3530, 3325, 3025, 2940, 2840, 2660, 2540, 2435, 2245, 1885, 1270, 855, 565, 375, 255, 0, 0, 0, 0, 0, 0)
            Case 11
                Mat = Array(3600, 3000, 2700, 2485, 2330, 2185, 2150, 2125, 2075, 2015, 1980, 1945, 1910, 1605, 1545, 1285, 995, 770, 565, 430, 310, 240, 170, 135)
            Case 12
                Mat = Array(3600, 3095, 2795, 2570, 2390, 2255, 2220, 2170, 2135, 2110, 2090, 2075, 1930, 1750, 1720, 1525, 1185, 925, 735, 585, 480, 380, 290, 205)
            Case 13
                Mat = Array(3000, 2530, 2270, 2065, 1910, 1800, 1750, 1715, 1680, 1645, 1610, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 14
                Mat = Array(3600, 3230, 2975, 2760, 2570, 2435, 2390, 2330, 2290, 2255, 2230, 2210, 1930, 1785, 1565, 1360, 1185, 925, 705, 550, 430, 325, 255, 190)
            Case 15
                Mat = Array(3600, 3300, 3070, 2870, 2700, 2570, 2520, 2470, 2460, 2435, 2425, 2245, 1930, 1820, 1800, 1610, 1370, 855, 615, 465, 345, 275, 205, 170)
            Case 16
                Mat = Array(3360, 3025, 2845, 2665, 2520, 2390, 2330, 2280, 2230, 2170, 2125, 2075, 1930, 1680, 1460, 1115, 860, 650, 495, 395, 290, 225, 155, 120)
            Case 17
                Mat = Array(3360, 3035, 2845, 2675, 2530, 2400, 2340, 2280, 2245, 2185, 2135, 2090, 1930, 1720, 1680, 1305, 945, 685, 515, 375, 290, 225, 165, 130)
            Case 18
                Mat = Array(3750, 3600, 3325, 3070, 2880, 2785, 2750, 2710, 2660, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        End Select
    Case "2500#"
        Select Case (k)
            Case 1
                Mat = Array(6170, 5625, 5470, 5280, 4990, 4560, 4475, 4440, 4200, 3430, 2230, 1430, 860, 430, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 2
                Mat = Array(6250, 6250, 6070, 5880, 5540, 5040, 4905, 4730, 4200, 3430, 2230, 1430, 860, 430, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 3
                Mat = Array(5785, 5470, 5315, 5145, 4850, 4440, 4355, 4320, 3945, 3260, 2230, 1430, 860, 430, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 4
                Mat = Array(5145, 4680, 4560, 4405, 4150, 3805, 3740, 3740, 3685, 3085, 2230, 1430, 860, 430, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 5
                Mat = Array(5785, 5660, 5435, 5330, 5180, 5040, 4905, 4730, 4430, 4230, 4060, 3745, 2345, 1370, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 6
                Mat = Array(6250, 6250, 6015, 5775, 5540, 5040, 4905, 4730, 4430, 4230, 4060, 3745, 2630, 1685, 1315, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 7
                Mat = Array(6250, 6250, 6015, 5775, 5540, 5040, 4905, 4730, 4430, 4230, 4060, 3745, 2655, 1800, 1200, 800, 515, 315, 0, 0, 0, 0, 0, 0)
            Case 8
                Mat = Array(6250, 6250, 6070, 5880, 5540, 5040, 4905, 4730, 4430, 4230, 4060, 3745, 3145, 2170, 1455, 915, 570, 345, 0, 0, 0, 0, 0, 0)
            Case 9
                Mat = Array(6250, 6205, 5965, 5880, 5540, 5040, 4905, 4730, 4400, 4230, 4030, 3085, 2285, 1655, 1200, 830, 515, 285, 0, 0, 0, 0, 0, 0)
            Case 10
                Mat = Array(6250, 6250, 6070, 5880, 5540, 5040, 4905, 4730, 4430, 4230, 4060, 3745, 3145, 2115, 1430, 945, 630, 430, 0, 0, 0, 0, 0, 0)
            Case 11
                Mat = Array(6000, 5000, 4500, 4140, 3880, 3640, 3580, 3540, 3460, 3360, 3300, 3240, 3180, 2675, 2570, 2145, 1655, 1285, 945, 715, 515, 400, 285, 230)
            Case 12
                Mat = Array(6000, 5160, 4660, 4280, 3980, 3760, 3700, 3620, 3560, 3520, 3480, 3460, 3220, 2915, 2865, 2545, 1970, 1545, 1230, 970, 800, 630, 485, 345)
            Case 13
                Mat = Array(5000, 4220, 3780, 3440, 3180, 3000, 2920, 2860, 2800, 2740, 2680, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
            Case 14
                Mat = Array(6000, 5380, 4960, 4600, 4285, 4060, 3980, 3880, 3820, 3760, 3720, 3680, 3220, 2970, 2605, 2265, 1970, 1545, 1170, 915, 715, 545, 430, 315)
            Case 15
                Mat = Array(6000, 5500, 5120, 4780, 4500, 4280, 4200, 4120, 4100, 4060, 4040, 3745, 3220, 3030, 13000, 2685, 2285, 1430, 1030, 770, 570, 455, 345, 285)
            Case 16
                Mat = Array(5600, 5040, 4740, 4440, 4200, 3980, 3880, 3800, 3720, 3620, 3540, 3460, 3220, 2800, 2430, 1860, 1430, 1085, 830, 660, 485, 370, 260, 200)
            Case 17
                Mat = Array(5600, 5060, 4740, 4260, 4220, 4000, 3900, 3800, 3740, 3640, 3560, 3480, 3220, 2865, 2800, 2170, 1570, 1145, 855, 630, 485, 370, 275, 215)
            Case 18
                Mat = Array(6250, 6000, 5540, 5120, 4800, 4640, 4580, 4520, 4430, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)

        End Select
End Select

TempArray = Array(100, 200, 300, 400, 500, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350, 1400, 1450, 1500)
i = 0
Do While temp >= TempArray(i)

    FlangeRating = (temp - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (Mat(i + 1) - Mat(i)) + Mat(i)
    i = i + 1
Loop

End Function

Function FlangeWeight(PipeNPS As Single, FlangeType As String, FlangeRating As String)

Dim PipeNPSArray, FlangeWeightArray As Variant
Dim FlangeTypeArray As Variant
Dim FlangeRatingArray As Variant
Dim WeldNeck, SlipOnThrd, LapJoint, Blind As Single
Dim NA As Single

Dim i, n, k As Integer

PipeNPSArray = Array(1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 24)
FlangeTypeArray = Array(WeldNeck, SlipOnThrd, LapJoint, Blind)
FlangeRatingArray = Array(150#, 300#, 600#, 900#, 1500#, 2500#)

Select Case FlangeType & FlangeRating
        
        Case "WeldNeck150#"
            FlangeWeightArray = Array(4, 6, 10, 11.5, 12, 16.5, 21, 26, 42, 54, 88, 114, 140, 165, 197, 268)
        Case "SlipOnThrd150#"
            FlangeWeightArray = Array(3, 5, 8, 9, 12, 13, 15, 19, 30, 43, 64, 90, 106, 130, 165, 220)
        Case "LapJoint150#"
            FlangeWeightArray = Array(3, 5, 8, 9, 11, 13, 15, 19, 30, 43, 64, 105, 140, 160, 195, 275)
        Case "Blind150#"
            FlangeWeightArray = Array(4, 5, 7, 9, 13, 17, 20, 27, 47, 70, 123, 140, 180, 220, 285, 430)
        
        Case "WeldNeck300#"
            FlangeWeightArray = Array(7, 9, 12, 18, 20, 26.5, 36, 45, 69, 100, 142, 206, 250, 320, 400, 580)
        Case "SlipOnThrd300#"
            FlangeWeightArray = Array(6.5, 7, 10, 14, 17, 24, 31, 39, 58, 81, 115, 165, 220, 280, 325, 490)
        Case "LapJoint300#"
            FlangeWeightArray = Array(6.5, 7, 10, 14, 17, 24, 28, 39, 58, 91, 140, 190, 234, 305, 375, 550)
        Case "Blind300#"
            FlangeWeightArray = Array(7, 8, 12, 16, 21, 28, 37, 50, 81, 124, 185, 250, 315, 415, 515, 800)
         
        Case "WeldNeck600#"
            FlangeWeightArray = Array(8, 12, 18, 23, 26, 42, 68, 81, 120, 190, 226, 347, 481, 555, 690, 977)
        Case "SlipOnThrd600#"
            FlangeWeightArray = Array(7, 9, 13, 16, 21, 37, 63, 80, 115, 177, 215, 259, 366, 476, 612, 876)
        Case "LapJoint600#"
            FlangeWeightArray = Array(7, 9, 12, 15, 20, 36, 63, 78, 112, 195, 240, 290, 400, 469, 604, 866)
        Case "Blind600#"
            FlangeWeightArray = Array(8, 10, 15, 20, 29, 41, 68, 86, 140, 231, 295, 378, 527, 665, 855, 1250)
        
        Case "WeldNeck900#"
            FlangeWeightArray = Array(14, 24, 31, 36, NA, 53, 86, 110, 187, 268, 372, 562, 685, 924, 1164, 2107)
        Case "SlipOnThrd900#"
            FlangeWeightArray = Array(14, 22, 31, 36, NA, 53, 83, 110, 172, 245, 326, 400, 459, 647, 792, 1480)
        Case "LapJoint900#"
            FlangeWeightArray = Array(14, 21, 25, 29, NA, 51, 81, 105, 190, 277, 371, 415, 488, 670, 868, 1659)
        Case "Blind900#"
            FlangeWeightArray = Array(14, 25, 32, 35, NA, 54, 87, 115, 200, 290, 415, 520, 619, 880, 1107, 2099)
        
        Case "WeldNeck1500#"
            FlangeWeightArray = Array(14, 25, 36, 48, NA, 73, 132, 165, 275, 455, 690, 940, 1250, 1625, 2050, 3325)
        Case "SlipOnThrd1500#"
            FlangeWeightArray = Array(14, 25, 36, 48, NA, 73, 132, 165, 260, 436, 667, 940, 1250, 1625, 2050, 2825)
        Case "LapJoint1500#"
            FlangeWeightArray = Array(14, 25, 35, 47, NA, 75, 140, 170, 286, 485, 749, 890, 1250, 1475, 1775, 2825)
        Case "Blind1500#"
            FlangeWeightArray = Array(14, 25, 35, 48, NA, 73, 140, 160, 302, 510, 775, 975, 1300, 1750, 2225, 3625)
        
        Case "WeldNeck2500#"
            FlangeWeightArray = Array(28, 42, 52, 94, NA, 146, 244, 378, 576, 1068, 1608, NA, NA, NA, NA, NA)
        Case "SlipOnThrd2500#"
            FlangeWeightArray = Array(25, 38, 55, 83, "NA", 127, 210, 323, 485, 925, 1300, NA, NA, NA, NA, NA)
        Case "LapJoint2500#"
            FlangeWeightArray = Array(24, 37, 53, 80, NA, 122, 204, 314, 471, 897, 1262, NA, NA, NA, NA, NA)
        Case "Blind2500#"
            FlangeWeightArray = Array(25, 39, 56, 86, NA, 133, 223, 345, 533, 1025, 1464, NA, NA, NA, NA, NA)
        
        End Select
n = 0
k = 0
Do While i < 16
If PipeNPS = PipeNPSArray(i) Then
        FlangeWeight = FlangeWeightArray(i)
        Exit Do
    End If
    i = i + 1
    FlangeType = FlangeTypeArray(n)
    FlangeRating = FlangeRatingArray(k)
Loop
n = n + 1
k = k + 1

End Function

Function CapWeight(CapSize As Single, CapSchedule As String) As Single
Dim CapSizeArray, CapWeightArray As Variant
Dim CapScheduleArray As Variant
Dim i, n As Integer
Dim SCH40, SCH80, SCH160, NA As Single

CapSizeArray = Array(1.5, 2, 2.5, 3, 3.5, 4, 5, 6, 8, 10, 12, 14, 16, 18, 20, 24, 30, 36, 40, 42, 48)
CapScheduleArray = Array(SCH40, SCH80, SCH160)

Select Case CapSchedule

    Case "SCH40"
        CapWeightArray = Array(0.54, 0.8, 1, 1.7, 2.3, 2.8, 4.6, 6.9, 11.8, 20.8, 30.3, 36.5, 43.5, 57, 75.7, 101, 137, 192, 234, 257, 331)
    Case "SCH80"
        CapWeightArray = Array(0.67, 0.92, 1.3, 2.1, 3, 3.5, 5.8, 9.3, 16, 26, 38, 47, 57, 78, 100, 145, 182, 256, 313, 343, 442)
    Case "SCH160"
        CapWeightArray = Array(0.6, 1.3, 1.8, 2.9, NA, 5.9, 10, 15, 31, 57, 95, 200, 297, 360, 420, 535, NA, NA, NA, NA, NA)
End Select
n = 0
Do While i < 21

    If CapSize = CapSizeArray(i) Then
       CapWeight = CapWeightArray(i)
        
    End If
    i = i + 1
    CapSchedule = CapScheduleArray(n)
Loop
n = n + 1

End Function


Public Function DPPipe(ID, Velocity, Re, dens1, dens2, Visc, L_eq, e) As Single
'
'  Function to calulate Pressur drop in pipe of Length EqLength
'
'   ID - Pipe inside diameter, inches
'   Re - Reynolds number
'   dens1 - Desnity at inlet, lb/ft3
'   dens2 - density at outlet, lb/ft3
'   visc - Average fluid viscosity, lb/hr-ft
'   L_eq - Equivalent pipe length, ft


Dim Dens As Single          'average density
Dim f As Single             'Fanning friction factor
Dim Fguess As Single        'Initial guess for friction factor
Dim x As Single             'Intemediate variable
Dim Error As Single
'
'Calculate average viscosity
Dens = (dens1 + dens2) / 2
'
'correct viscosity units for pressure drop calc. to lb/s-ft

Visc = Visc * 2.419 / 3600
'
'Calculate fluid velocity in ft/s

'Laminar Flow
If Re < 2300 Then
   DPPipe = 32 * Visc * L_eq * Velocity / ((ID / 12) ^ 2 * 32.2 * Dens) * Dens / 144
'turbulent flow (2300-4000 actually transition)
ElseIf Re > 2300 Then
'   Calculate friction factor assuming fully rough flow
    Fguess = 0.05
    Error = 1
    Do While Error > 0.01
        x = (e / (ID / 12) / 3.7 + 2.51 / Re / Sqr(Fguess))
        f = (1 / (-2 * Application.WorksheetFunction.Log10(x))) ^ 2
        Error = Abs(f / Fguess - 1)
        Fguess = f
    Loop
    '
   
    'Pressure drop Calculation
    DPPipe = f * L_eq * Velocity ^ 2 / (ID / 12 * 2 * 32.2) * Dens / 144
'
End If
'
End Function
Function FuelTrainPipeDP(PipeNomSize As Single, PipeSCH As String, FuelMW As Single, FuelTemp As Single, FuelPressure As Single, FuelFlowRate As Single, FuelTrainLength As Single)

Dim OD As Single
Dim Wall As Single
Dim ID As Single
Dim a As Single
Dim StdDensity As Single
Dim ActDensity As Single
Dim Mass As Single
Dim ActFlow As Single
Dim vel As Single
Dim Ren As Single

OD = PipeOD(PipeNomSize)
Wall = PipeWall(PipeNomSize, PipeSCH)
ID = PipeID(OD, Wall)
a = Area(ID)
StdDensity = IdealGasDensity(FuelMW, 0, 60)
ActDensity = IdealGasDensity(FuelMW, FuelPressure, FuelTemp)
Mass = StdDensity * FuelFlowRate
ActFlow = Mass / ActDensity
vel = ActFlow / (a * 25)
Ren = Reynolds(vel, ID, 0.02, ActDensity)
FuelTrainPipeDP = DPPipe(ID, vel / 3600, Ren, ActDensity, ActDensity, 0.02, FuelTrainLength, 0.00015)

End Function
Function PipeVelocity(PipeNomSize As Single, PipeSCH As String, FuelMW As Single, FuelTemp As Single, FuelPressure As Single, FuelFlowRate As Single)

Dim OD As Single
Dim Wall As Single
Dim ID As Single
Dim a As Single
Dim StdDensity As Single
Dim ActDensity As Single
Dim Mass As Single
Dim ActFlow As Single

OD = PipeOD(PipeNomSize)
Wall = PipeWall(PipeNomSize, PipeSCH)
ID = PipeID(OD, Wall)
a = Area(ID)
StdDensity = IdealGasDensity(FuelMW, 0, 60)
ActDensity = IdealGasDensity(FuelMW, FuelPressure, FuelTemp)
Mass = StdDensity * FuelFlowRate
ActFlow = Mass / ActDensity
PipeVelocity = ActFlow / (a * 25)

End Function

Function ElbowWeight(ElbowOD As Single, Schedule As String, radius As String) As Single

Dim ElbowODArray, ElbowWeightArray As Variant
Dim ScheduleArray As Variant
Dim i, n As Integer
Dim SCH40SR, SCH40LR, SCH80SR, SCH80LR, SCH160SR, SCH160LR, NA As Single

ElbowODArray = Array(1.9, 2.375, 2.875, 3.5, 4, 4.5, 5.563, 6.625, 8.625, 10.75, 12.75, 14, 16, 18, 20, 24)
ScheduleArray = Array(SCH40SR, SCH40LR, SCH80SR, SCH80LR, SCH160SR, SCH160LR)

Select Case Schedule & radius
    Case "SCH40SR"
        ElbowWeightArray = Array(0.5, 1, 2, 3, 4.3, 6.1, 9.7, 16.7, 32.4, 56.3, 80, 105, 130, 165, 215, 300)
    Case "SCH40LR"
        ElbowWeightArray = Array(0.8, 1.6, 3.2, 4.8, 6.6, 8.9, 15.1, 24, 47.8, 83.4, 123, 155, 206, 262, 324, 466)
    Case "SCH80SR"
        ElbowWeightArray = Array(0.8, 1.5, 2.6, 3.8, 5.4, 7.6, 13.8, 22.8, 47.3, 75, 105, 140, 175, 215, 280, 400)
    Case "SCH80LR"
        ElbowWeightArray = Array(1.2, 2.1, 3.8, 6.3, 8.6, 12.5, 21.2, 34.4, 71.3, 111, 158, 201, 270, 348, 422, 604)
    Case "SCH160SR"
        ElbowWeightArray = Array(1.5, 2.9, 5.5, 9.8, NA, 20, 30, 60, 125, 258, 455, 550, 800, 1025, 1295, 1450)
    Case "SCH160LR"
        ElbowWeightArray = Array(1.8, 3.2, 6, 10, NA, 22, 33, 62, 122, 270, 460, 563, 825, 1240, 1510, 1760)

End Select
n = 0
Do While i < 15
    If ElbowOD = ElbowODArray(i) Then
        ElbowWeight = ElbowWeightArray(i)
        Exit Do
    End If
    i = i + 1
        Schedule = ScheduleArray(n)
  Loop
n = n + 1

End Function






Function PipeWeight(OD As Single, ID As Single, length As Single, Density As Single)

'OD and ID in inches
'Length in feet
'Density in lb/ft3

Density = 490.75

Dim OutsideArea, InsideArea, CrossSectionalArea As Single

OutsideArea = pi * OD ^ 2 / 4
InsideArea = pi * ID ^ 2 / 4

CrossSectionalArea = (OutsideArea - InsideArea) / 144

PipeWeight = CrossSectionalArea * length * Density
End Function


Function LossCoefficientReturn(Reynolds As Single, ReturnRadius As Single, ID As Single, Angle As Single, frictionFactor As Single) As Single

Dim A1, B1, C1 As Single
Dim gamma, gamma_loc As Single


A1 = 0.7 + 0.35 * (Angle / 90)
B1 = 0.21 / (ReturnRadius / ID) ^ 0.5
C1 = 1  'for circular or square cross section, C1=1.0

gamma = A1 * B1 * C1
gamma = gamma_loc + 0.0175 * (ReturnRadius / ID) * Angle * frictionFactor

LossCoefficientReturn = gamma
End Function

Function LossCoefficientMiteredReturn(Reynolds As Single, ReturnRadius As Single, ID As Single, frictionFactor As Single) As Single

Dim gamma, D_ratio, i As Single
Dim D_ratioArray, GammaArray As Variant

D_ratio = (ReturnRadius) / ID

D_ratioArray = Array(0, 0.25, 0.48, 0.7, 0.97, 1.2, 2.4, 3.6, 4.8, 6, 7.25)
GammaArray = Array(1.1, 0.95, 0.72, 0.6, 0.42, 0.38, 0.32, 0.38, 0.41, 0.4, 0.41)

i = 0
Do While D_ratio > D_ratioArray(i)
    
    gamma = (D_ratio - D_ratioArray(i)) / (D_ratioArray(i + 1) - D_ratioArray(i)) * (GammaArray(i + 1) - GammaArray(i)) + GammaArray(i)
    i = i + 1
Loop

LossCoefficientMiteredReturn = gamma * 2

End Function
Function ReturnRadius(PipeNom As Single, PipeOD As Single, ReturnSpacing As String) As Single

Dim Spacing As Single

If ReturnSpacing = "SR" Then
    Spacing = 2 * PipeNom
ElseIf ReturnSpacing = "LR" Then
    Spacing = 3 * PipeNom
End If

ReturnRadius = Spacing / 2

End Function

Function ReturnWeight(OD As Single, ID As Single, ReturnRadius As Single, Number As Integer, Density As Single)

'OD and ID in inches
'Length in feet
'Density in lb/ft3

Dim OutsideArea, InsideArea, CrossSectionalArea, length As Single

OutsideArea = pi * OD ^ 2 / 4
InsideArea = pi * ID ^ 2 / 4
length = pi * ReturnRadius / 12
CrossSectionalArea = (OutsideArea - InsideArea) / 144

ReturnWeight = CrossSectionalArea * length * Density * Number
End Function
Public Function Reynolds(Velocity As Single, diameter As Single, Viscosity As Single, Density As Single)

'Calculates Reynolds number
'Velocity in ft/s
'Diameter in inches
'Viscosity (dynamic) in cP (centipoise)
'Density in lb/ft3

'convert units from ft/s to ft/hr
Velocity = Velocity * 3600

'convert units from cP to lb/ft*hr
Viscosity = Viscosity * 2.419


Reynolds = Velocity * (diameter / 12) / (Viscosity / Density)

End Function


Public Function Prandtl2(Viscosity As Single, SpecificHeat As Single, ThermalConductivity As Single)

'Calculates Prandtl number
'Viscosity (dynamic) in cP (centipoise)
'Specific Heat in BTU/lb*°F
'Thermal Conductivity in BTU/hr*ft*°F

'convert units from cP to lb/ft*hr
Viscosity = Viscosity * 2.419


Prandtl2 = Viscosity * SpecificHeat / ThermalConductivity

End Function


Function API530RuptureWall(Material As String, PipeOD As Single, Pressure As Single, Temperature As Single, CorrosionAllowance As Single)

End Function

Dim TempArray, Stress As Variant
Dim TempHigh, TempLow, StressHigh, StressLow, AllowableStress As Single
Dim i As Integer

TempArray = Array(100, 300, 400, 500, 600, 700, 800, 900, 1000, 1100, 11200, 1300, 1400, 1500)

Select Case Material
    Case "SA-106-B"
        Stress = Array(0, 0, 0, 0, 0, 21, 13, 7.5, 4.2, 0, 0, 0, 0, 0)
    Case "SA-335-P11"
        Stress = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 17.5, 6.8, 2.55)
    Case "SA-335-P22"
        Stress = Array(0, 0, 0, 0, 0, 0, 0, 16, 8.6, 4.5, 1.7, 0, 0, 0)
    Case "SA-312-TP304"
        Stress = Array(0, 0, 0, 0, 0, 0, 0, 0, 17.5, 9.3, 5.2, 2.9, 1.6)
   
End Select

i = 0
Do While TempArray(i) <> 0
    If Temperature <= TempArray(i) Then
        TempHigh = TempArray(i)
        TempLow = TempArray(i - 1)
        Exit Do
    End If
i = i + 1
Loop

On Error GoTo Error

StressHigh = Stress(i) * 10 ^ 3
StressLow = Stress(i - 1) * 10 ^ 3

AllowableStress = (Temperature - TempLow) / (TempHigh - TempLow) * (StressHigh - StressLow) + StressLow

API530RuptureWall = Pressure * PipeOD / (2 * AllowableStress + Pressure) + CorrosionAllowance
Exit Function

Error:
API530RuptureWall = "Error"


Function API530ElasticWall(Material As String, PipeID As Single, Pressure As Single, Temperature As Single, CorrosionAllowance As Single)

Dim TempArray, Stress As Variant
Dim TempHigh, TempLow, StressHigh, StressLow, AllowableStress As Single
Dim i As Integer

TempArray = Array(100, 300, 400, 500, 600, 700, 800, 900, 1000, 1100, 11200, 1300, 1400, 1500)

Select Case Material
    Case "SA-106-B"
        Stress = Array(20, 20, 18.5, 17.5, 16.5, 15.5, 14.5, 13.5, 11.5, 0, 0, 0, 0, 0)
    Case "SA-335-P11"
        Stress = Array(17.8, 17.8, 17, 16.5, 16, 15.5, 15, 14, 13, 11.5, 0, 0, 0, 0)
    Case "SA-335-P22"
        Stress = Array(17.8, 17.8, 17.8, 17.8, 18, 18, 18, 17, 15.5, 13.5, 10.8, 0, 0, 0)
    Case "SA-312-TP304"
        Stress = Array(19, 19, 19, 17.5, 17, 16.25, 15.75, 15.5, 15, 14.9, 14.5, 14, 13, 11.5)
   
End Select

i = 0
Do While TempArray(i) <> 0
    If Temperature <= TempArray(i) Then
        TempHigh = TempArray(i)
        TempLow = TempArray(i - 1)
        Exit Do
    End If
i = i + 1
Loop

On Error GoTo Error

StressHigh = Stress(i) * 10 ^ 3
StressLow = Stress(i - 1) * 10 ^ 3

AllowableStress = (Temperature - TempLow) / (TempHigh - TempLow) * (StressHigh - StressLow) + StressLow

API530ElasticWall = Pressure * (PipeID - 2 * CorrosionAllowance) / (2 * AllowableStress - Pressure) + CorrosionAllowance
Exit Function

Error:
API530ElasticWall = "Error"

End Function

Public Function frictionFactor(Re As Single, diameter As Single)

Dim AverageDensity As Single        'average density
Dim f As Single                     'Fanning friction factor
Dim Fguess As Single                'Initial guess for friction factor
Dim x As Single                     'Intemediate variable
Dim Error As Single                 'convergence error
Dim e As Single
'Diameter in inches
'Reynolds number is unitless

Fguess = 0.05
e = 0.00015
    Error = 1
    Do While Error > 0.01
        x = (e / (diameter / 12) / 3.7 + 2.51 / Re / Sqr(Fguess))
        f = (1 / (-2 * Application.WorksheetFunction.Log10(x))) ^ 2
        Error = Abs(f / Fguess - 1)
        Fguess = f
    Loop
    
frictionFactor = f

End Function
Function ReturnSurfaceArea(PipeOD As Single, ReturnRadius As Single, Number As Integer)

ReturnSurfaceArea = (pi * ReturnRadius * (pi * PipeOD) * Number) / 144

End Function

Function SecVIIIMinWall(Material As String, PipeID As Single, Pressure As Single, Temperature As Single, CorrosionAllowance As Single, Efficiency As Single)

Dim TempArray, Stress As Variant
Dim TempHigh, TempLow, StressHigh, StressLow, AllowableStress As Single
Dim i As Integer

TempArray = Array(-20, 100, 200, 300, 400, 500, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350, 1400, 1450, 1500, 1550, 1600, 1650, 1700, 1750, 1800)

Select Case Material
    Case "SA-106-B"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 15.6, 13, 10.8, 8.7, 5.9, 4, 2.5, 0)
    Case "SA-333-6"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 15.6, 13, 10.8, 8.7, 5.9, 4, 2.5, 0)
    Case "SA-335-P11"
         Stress = Array(17.1, 17.1, 17.1, 17.1, 16.8, 16.2, 15.7, 15.4, 15.1, 14.8, 14.4, 14, 13.6, 9.3, 6.3, 4.2, 2.8, 1.9, 1.2, 0)
    Case "SA-335-P22"
        Stress = Array(17.1, 17.1, 17.1, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 13.6, 10.8, 8, 5.7, 3.8, 2.4, 1.4, 0)
    Case "SA-335-P5"
        Stress = Array(17.1, 17.1, 17.1, 16.6, 16.5, 16.4, 16.2, 15.9, 15.6, 15.1, 14.5, 13.8, 10.9, 8, 5.8, 4.2, 2.9, 1.8, 1, 0)
    Case "SA-335-P9"
        Stress = Array(17.1, 17.1, 17.1, 16.6, 16.5, 16.4, 16.2, 15.9, 15.6, 15.1, 14.5, 13.8, 13, 10.6, 7.4, 5, 3.3, 2.2, 1.5)
    Case "SA-312-TP304"
        Stress = Array(20, 20, 16.7, 15, 13.8, 12.9, 12.3, 12, 11.7, 11.5, 11.2, 11, 10.8, 10.6, 10.4, 10.1, 9.8, 7.7, 6.1, 4.7, 3.7, 2.9, 2.3, 1.8, 1.4)
    Case "SA-312-TP304W"
        Stress = Array(17, 17, 14.2, 12.7, 11.7, 11, 10.4, 10.2, 10, 9.8, 9.6, 9.4, 9.2, 9, 8.8, 8.6, 8.3, 6.6, 5.2, 4, 3.1, 2.5, 2, 1.6, 1.2)
    Case "SA-312-TP316"
        Stress = Array(20, 20, 17.3, 15.6, 14.3, 13.3, 12.6, 12.3, 12.1, 11.9, 11.8, 11.6, 11.5, 11.4, 11.3, 11.2, 11.1, 9.8, 7.4, 5.5, 4.1, 3.1, 2.3, 1.7, 1.3)
    Case "SA-312-TP321"
        Stress = Array(20, 20, 18, 16.5, 15.3, 14.3, 13.5, 13.2, 13, 12.7, 12.6, 12.4, 12.3, 12.1, 12, 9.6, 6.9, 5, 3.6, 2.6, 1.7, 1.1, 0.8, 0.5, 0.3)
    Case "SA-312-TP347"
        Stress = Array(20, 20, 18.4, 17.1, 16, 15, 14.3, 14, 13.8, 13.7, 13.6, 13.5, 13.4, 13.4, 13.4, 12.1, 9.1, 6.1, 4.4, 3.3, 2.2, 1.5, 1.2, 0.9, 0.8)
    Case "SA-312-TP347H"
        Stress = Array(20, 20, 18.4, 17.1, 16, 15, 14.3, 14, 13.8, 13.7, 13.6, 13.5, 13.4, 13.4, 13.4, 13.4, 13.3, 10.5, 7.9, 5.9, 4.4, 3.2, 2.5, 1.8, 1.3)
    Case "HR 160"
        Stress = Array(22.8, 22.8, 20.2, 18.54999924, 16.9, 15.35000038, 13.8, 13.60000038, 13.39999962, 13.19999981, 13, 12.9375, 12.875, 12.8125, 12.75, 12.77499962, 12.80000019, 12.82500076, 12.85, 12.72500038, 12.60000038, 12.47500038, 12.35, 12.02500057, 11.70000076)
    'Case "SB-407" UNS 8800
    '    Stress = Array(20, 20, 18.5, 17.8, 17.2, 16.8, 16.3, 16.1, 15.9, 15.7, 15.5, 15.3, 15.1, 14.9, 14.7, 14.5, 13, 9.8, 6.6, 4.2, 2, 1.6, 1.1, 1, 0.8)
    Case "SB-407" 'UNS N08810
        Stress = Array(16.7, 16.7, 15.4, 14.4, 13.6, 12.9, 12.2, 11.9, 11.6, 11.4, 11.1, 10.9, 10.7, 10.5, 10.4, 10.2, 10#, 9.3, 7.4, 5.9, 4.7, 3.8, 3#, 2.4, 1.9, 1.4, 1.1, 0.86, 0.71, 0.56, 0.44)
    Case "SB-622" 'UNS N06230
        Stress = Array(30, 30, 28.2, 26.4, 24.7, 23.1, 22, 21.5, 21.2, 21, 20.9, 20.9, 20.9, 20.9, 20.9, 20.9, 20.9, 19, 15.6, 12.9, 10.6, 8.5, 6.7, 5.3, 4.1, 2.9, 2.1, 1.5, 1.1, 0.7, 0.45)
    Case "S30815"
        Stress = Array(24.9, 24.9, 24.7, 22, 19.9, 18.5, 17.7, 17.4, 17.2, 17, 16.8, 16.6, 16.4, 16.2, 14.9, 11.6, 9, 6.9, 5.2, 4, 3.1, 2.4, 1.9, 1.6, 1.3, 1, 0.86, 0.71)
    Case "740H"
        TempArray = Array(70, 1292, 1337, 1382, 1424, 1472)
        Stress = Array(107.7, 88.2, 87, 83.5, 81.8, 79.3)
    Case "SB-167" 'UNS N06617
        Stress = Array(23.3, 23.3, 23.3, 23.3, 23.3, 23.3, 22.5, 22.1, 21.9, 21.7, 21.5, 21.3, 21.2, 21, 20.9, 20.9, 20.8, 20.7, 18.1, 14.5, 11.2, 8.7, 6.6, 5.1, 3.9, 3, 2.3, 1.8, 1.4, 1.1, 0.73)
    Case "SB-444" 'UNS N06625
        Stress = Array(26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.6, 26.4, 26.3, 26.2, 26.1, 20, 15, 11.6, 8.5, 6.7, 4.9, 3.8, 2.6, 1.9)
End Select

i = 0
Do While TempArray(i) <> 0
    If Temperature <= TempArray(i) Then
        TempHigh = TempArray(i)
        TempLow = TempArray(i - 1)
        Exit Do
    End If
i = i + 1
Loop

On Error GoTo Error

StressHigh = Stress(i) * 10 ^ 3
StressLow = Stress(i - 1) * 10 ^ 3

AllowableStress = (Temperature - TempLow) / (TempHigh - TempLow) * (StressHigh - StressLow) + StressLow

SecVIIIMinWall = Pressure * (PipeID / 2) / (AllowableStress * Efficiency + 0.4 * Pressure) + CorrosionAllowance
Exit Function

Error:
SecVIIIMinWall = "Temp too high"

End Function
Function SecVIIIMaxPressure(Material As String, PipeID As Single, PipeWall As Single, Temperature As Single, CorrosionAllowance As Single, Efficiency As Single)

Dim TempArray, Stress As Variant
Dim TempHigh, TempLow, StressHigh, StressLow, AllowableStress As Single
Dim i As Integer

TempArray = Array(-20, 100, 200, 300, 400, 500, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350, 1400, 1450, 1500, 1550, 1600, 1650)

Select Case Material
    Case "SA-106-B"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 15.6, 13, 10.8, 8.7, 5.9, 4, 2.5, 0)
    Case "SA-333-6"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 15.6, 13, 10.8, 8.7, 5.9, 4, 2.5, 0)
    Case "SA-335-P11"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 16.8, 16.2, 15.7, 15.4, 15.1, 14.8, 14.4, 14, 13.6, 9.3, 6.3, 4.2, 2.8, 1.9, 1.2, 0)
    Case "SA-335-P22"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 13.6, 10.8, 8, 5.7, 3.8, 2.4, 1.4, 0)
    Case "SA-335-P5"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 16.6, 16.5, 16.4, 16.2, 15.9, 15.6, 15.1, 14.5, 13.8, 10.9, 8, 5.8, 4.2, 2.9, 1.8, 1, 0)
    Case "SA-335-P9"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 16.6, 16.5, 16.4, 16.2, 15.9, 15.6, 15.1, 14.5, 13.8, 13, 10.6, 7.4, 5, 3.3, 2.2, 1.5)
    Case "SA-312-TP304"
        Stress = Array(20, 20, 16.7, 15, 13.8, 12.9, 12.3, 12, 11.7, 11.5, 11.2, 11, 10.8, 10.6, 10.4, 10.1, 9.8, 7.7, 6.1, 4.7, 3.7, 2.9, 2.3, 1.8, 1.4)
    Case "SA-312-TP304W"
        Stress = Array(17, 17, 14.2, 12.7, 11.7, 11, 10.4, 10.2, 10, 9.8, 9.6, 9.4, 9.2, 9, 8.8, 8.6, 8.3, 6.6, 5.2, 4, 3.1, 2.5, 2, 1.6, 1.2)
    Case "SA-312-TP316"
        Stress = Array(20, 20, 17.3, 15.6, 14.3, 13.3, 12.6, 12.3, 12.1, 11.9, 11.8, 11.6, 11.5, 11.4, 11.3, 11.2, 11.1, 9.8, 7.4, 5.5, 4.1, 3.1, 2.3, 1.7, 1.3)
    Case "SA-312-TP321"
        Stress = Array(20, 20, 18, 16.5, 15.3, 14.3, 13.5, 13.2, 13, 12.7, 12.6, 12.4, 12.3, 12.1, 12, 9.6, 6.9, 5, 3.6, 2.6, 1.7, 1.1, 0.8, 0.5, 0.3)
    Case "SA-312-TP347"
        Stress = Array(20, 20, 18.4, 17.1, 16, 15, 14.3, 14, 13.8, 13.7, 13.6, 13.5, 13.4, 13.4, 13.4, 12.1, 9.1, 6.1, 4.4, 3.3, 2.2, 1.5, 1.2, 0.9, 0.8)
    Case "SA-312-TP347H"
        Stress = Array(20, 20, 18.4, 17.1, 16, 15, 14.3, 14, 13.8, 13.7, 13.6, 13.5, 13.4, 13.4, 13.4, 13.4, 13.3, 10.5, 7.9, 5.9, 4.4, 3.2, 2.5, 1.8, 1.3)
    Case "HR 160"
        Stress = Array(22.8, 22.8, 20.2, 18.54999924, 16.9, 15.35000038, 13.8, 13.60000038, 13.39999962, 13.19999981, 13, 12.9375, 12.875, 12.8125, 12.75, 12.77499962, 12.80000019, 12.82500076, 12.85, 12.72500038, 12.60000038, 12.47500038, 12.35, 12.02500057, 11.70000076)
    'Case "SB-407" UNS 8800
    '    Stress = Array(20, 20, 18.5, 17.8, 17.2, 16.8, 16.3, 16.1, 15.9, 15.7, 15.5, 15.3, 15.1, 14.9, 14.7, 14.5, 13, 9.8, 6.6, 4.2, 2, 1.6, 1.1, 1, 0.8)
    Case "SB-407" 'UNS N08810
        Stress = Array(16.7, 16.7, 16.7, 16.7, 16.7, 16.7, 16.5, 16.1, 15.7, 15.3, 15, 14.7, 14.5, 14.2, 14, 13.8, 11.6, 9.3, 7.4, 5.9, 4.7, 3.8, 3, 2.4, 1.9)
    Case "S30815"
        Stress = Array(24.9, 24.9, 24.7, 22, 19.9, 18.5, 17.7, 17.4, 17.2, 17, 16.8, 16.6, 16.4, 16.2, 14.9, 11.6, 9, 6.9, 5.2, 4, 3.1, 2.4, 1.9, 1.6, 1.3, 1, 0.86, 0.71)
    Case "SB-167" 'UNS N06617
        Stress = Array(23.3, 23.3, 23.3, 23.3, 23.3, 23.3, 22.5, 22.1, 21.9, 21.7, 21.5, 21.3, 21.2, 21, 20.9, 20.9, 20.8, 20.7, 18.1, 14.5, 11.2, 8.7, 6.6, 5.1, 3.9, 3, 2.3, 1.8, 1.4, 1.1, 0.73)
    Case "SB-444" 'UNS N06625
        Stress = Array(26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.6, 26.4, 26.3, 26.2, 26.1, 20, 15, 11.6, 8.5, 6.7, 4.9, 3.8, 2.6, 1.9)
    Case "SB-622" 'UNS N06230
        Stress = Array(30, 30, 28.2, 26.4, 24.7, 23.1, 22, 21.5, 21.2, 21, 20.9, 20.9, 20.9, 20.9, 20.9, 20.9, 20.9, 19, 15.6, 12.9, 10.6, 8.5, 6.7, 5.3, 4.1, 2.9, 2.1, 1.5, 1.1, 0.7, 0.45)
    Case "740H"
        TempArray = Array(70, 1292, 1337, 1382, 1424, 1472)
        Stress = Array(107.7, 88.2, 87, 83.5, 81.8, 79.3)
End Select

i = 0
Do While TempArray(i) <> 0
    If Temperature <= TempArray(i) Then
        TempHigh = TempArray(i)
        TempLow = TempArray(i - 1)
        Exit Do
    End If
i = i + 1
Loop

On Error GoTo Error

StressHigh = Stress(i) * 10 ^ 3
StressLow = Stress(i - 1) * 10 ^ 3

AllowableStress = (Temperature - TempLow) / (TempHigh - TempLow) * (StressHigh - StressLow) + StressLow

SecVIIIMaxPressure = AllowableStress * Efficiency * (0.875 * PipeWall - CorrosionAllowance) / (PipeID / 2 + 0.6 * (0.875 * PipeWall - CorrosionAllowance))
Exit Function

Error:
SecVIIIMaxPressure = "Temp too high or value not found"

End Function
Function SecVIIIAllowableStress(Material As String, Temperature As Single) As Single

Dim TempArray, Stress As Variant
Dim TempHigh, TempLow, StressHigh, StressLow, AllowableStress As Single
Dim i As Integer

TempArray = Array(-20, 100, 200, 300, 400, 500, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100, 1150, 1200, 1250, 1300, 1350, 1400, 1450, 1500, 1550, 1600, 1650, 1700, 1750, 1800)

Select Case Material
    Case "SA-106-B"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 15.6, 13, 10.8, 8.7, 5.9, 4, 2.5, 0)
    Case "SA-333-6"
        Stress = Array(17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 17.1, 15.6, 13, 10.8, 8.7, 5.9, 4, 2.5, 0)
    Case "SA-335-P11"
       Stress = Array(17.1, 17.1, 17.1, 17.1, 16.8, 16.2, 15.7, 15.4, 15.1, 14.8, 14.4, 14, 13.6, 9.3, 6.3, 4.2, 2.8, 1.9, 1.2, 0)
    Case "SA-335-P22"
        Stress = Array(17.1, 17.1, 17.1, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 16.6, 13.6, 10.8, 8, 5.7, 3.8, 2.4, 1.4, 0)
    Case "SA-335-P5"
        Stress = Array(17.1, 17.1, 17.1, 16.6, 16.5, 16.4, 16.2, 15.9, 15.6, 15.1, 14.5, 13.8, 10.9, 8, 5.8, 4.2, 2.9, 1.8, 1, 0)
    Case "SA-335-P9"
        Stress = Array(17.1, 17.1, 17.1, 16.6, 16.5, 16.4, 16.2, 15.9, 15.6, 15.1, 14.5, 13.8, 13, 10.6, 7.4, 5, 3.3, 2.2, 1.5)
    Case "SA-312-TP304"
        Stress = Array(20, 20, 16.7, 15, 13.8, 12.9, 12.3, 12, 11.7, 11.5, 11.2, 11, 10.8, 10.6, 10.4, 10.1, 9.8, 7.7, 6.1, 4.7, 3.7, 2.9, 2.3, 1.8, 1.4)
    Case "SA-312-TP304W"
        Stress = Array(17, 17, 14.2, 12.7, 11.7, 11, 10.4, 10.2, 10, 9.8, 9.6, 9.4, 9.2, 9, 8.8, 8.6, 8.3, 6.6, 5.2, 4, 3.1, 2.5, 2, 1.6, 1.2)
    Case "SA-312-TP316"
        Stress = Array(20, 20, 17.3, 15.6, 14.3, 13.3, 12.6, 12.3, 12.1, 11.9, 11.8, 11.6, 11.5, 11.4, 11.3, 11.2, 11.1, 9.8, 7.4, 5.5, 4.1, 3.1, 2.3, 1.7, 1.3)
    Case "SA-312-TP321"
        Stress = Array(20, 20, 18, 16.5, 15.3, 14.3, 13.5, 13.2, 13, 12.7, 12.6, 12.4, 12.3, 12.1, 12, 9.6, 6.9, 5, 3.6, 2.6, 1.7, 1.1, 0.8, 0.5, 0.3)
    Case "SA-312-TP347"
        Stress = Array(20, 20, 18.4, 17.1, 16, 15, 14.3, 14, 13.8, 13.7, 13.6, 13.5, 13.4, 13.4, 13.4, 12.1, 9.1, 6.1, 4.4, 3.3, 2.2, 1.5, 1.2, 0.9, 0.8)
    Case "SA-312-TP347H"
        Stress = Array(20, 20, 18.4, 17.1, 16, 15, 14.3, 14, 13.8, 13.7, 13.6, 13.5, 13.4, 13.4, 13.4, 13.4, 13.3, 10.5, 7.9, 5.9, 4.4, 3.2, 2.5, 1.8, 1.3)
    Case "HR 160"
        Stress = Array(22.8, 22.8, 20.2, 18.54999924, 16.9, 15.35000038, 13.8, 13.60000038, 13.39999962, 13.19999981, 13, 12.9375, 12.875, 12.8125, 12.75, 12.77499962, 12.80000019, 12.82500076, 12.85, 12.72500038, 12.60000038, 12.47500038, 12.35, 12.02500057, 11.70000076)
    'Case "SB-407" UNS 8800
    '    Stress = Array(20, 20, 18.5, 17.8, 17.2, 16.8, 16.3, 16.1, 15.9, 15.7, 15.5, 15.3, 15.1, 14.9, 14.7, 14.5, 13, 9.8, 6.6, 4.2, 2, 1.6, 1.1, 1, 0.8)
    Case "SB-407" 'UNS N08810
        Stress = Array(16.7, 16.7, 15.4, 14.4, 13.6, 12.9, 12.2, 11.9, 11.6, 11.4, 11.1, 10.9, 10.7, 10.5, 10.4, 10.2, 10#, 9.3, 7.4, 5.9, 4.7, 3.8, 3#, 2.4, 1.9, 1.4, 1.1, 0.86, 0.71, 0.56, 0.44)
    Case "S30815"
        Stress = Array(24.9, 24.9, 24.7, 22, 19.9, 18.5, 17.7, 17.4, 17.2, 17, 16.8, 16.6, 16.4, 16.2, 14.9, 11.6, 9, 6.9, 5.2, 4, 3.1, 2.4, 1.9, 1.6, 1.3, 1, 0.86, 0.71)
    Case "SB-167" 'UNS N06617
        Stress = Array(23.3, 23.3, 23.3, 23.3, 23.3, 23.3, 22.5, 22.1, 21.9, 21.7, 21.5, 21.3, 21.2, 21, 20.9, 20.9, 20.8, 20.7, 18.1, 14.5, 11.2, 8.7, 6.6, 5.1, 3.9, 3, 2.3, 1.8, 1.4, 1.1, 0.73)
    Case "SB-444" 'UNS N06625
        Stress = Array(26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.7, 26.6, 26.4, 26.3, 26.2, 26.1, 20, 15, 11.6, 8.5, 6.7, 4.9, 3.8, 2.6, 1.9)
    Case "SB-622" 'UNS N06230
        Stress = Array(30, 30, 28.2, 26.4, 24.7, 23.1, 22, 21.5, 21.2, 21, 20.9, 20.9, 20.9, 20.9, 20.9, 20.9, 20.9, 19, 15.6, 12.9, 10.6, 8.5, 6.7, 5.3, 4.1, 2.9, 2.1, 1.5, 1.1, 0.7, 0.45)
    Case "740H"
        TempArray = Array(70, 1292, 1337, 1382, 1424, 1472)
        Stress = Array(107.7, 88.2, 87, 83.5, 81.8, 79.3)
End Select

i = 0
Do While TempArray(i) <> 0
    If Temperature <= TempArray(i) Then
        TempHigh = TempArray(i)
        TempLow = TempArray(i - 1)
        Exit Do
    End If
i = i + 1
Loop

On Error GoTo Error

StressHigh = Stress(i) * 10 ^ 3
StressLow = Stress(i - 1) * 10 ^ 3

SecVIIIAllowableStress = (Temperature - TempLow) / (TempHigh - TempLow) * (StressHigh - StressLow) + StressLow
Exit Function

Error:
SecVIIIAllowableStress = "Temp too high or value not found"

End Function

Function TubeResistance(OD As Single, ID As Single, TubeThermCond As Single)

Dim OutsideRadius, InsideRadius As Single

'Radii must be in feet
OutsideRadius = (OD / 2) / 12
InsideRadius = (ID / 2) / 12

TubeResistance = OutsideRadius * Log(OutsideRadius / InsideRadius) / TubeThermCond

End Function

Function LossCoefficientReducer(alpha As Single, initialDiameter As Single, finalDiameter As Single, frictionFactor) As Single

'Dim crossSection1, crossSection2, n0, A As Single
'alpha = Atn((initialDiameter - finalDiameter) / length) * 180 / pi
'crossSection1 = pi * initialDiameter ^ 2 / 4
'crossSection0 = pi * finalDiameter ^ 2 / 4
'n0 = crossSection0 / crossSection1
'A = 19 / ((n0 ^ 0.5) * (Tan(alpha * pi / 180) ^ 0.75))
'LossCoefficientReducer = A / Reynolds

Dim crossSection1, crossSection0, n0, alphaR As Single

crossSection1 = pi * initialDiameter ^ 2 / 4
crossSection0 = pi * finalDiameter ^ 2 / 4
n0 = crossSection0 / crossSection1
alphaR = 0.01745 * alpha

LossCoefficientReducer = ((-0.0125 * n0 ^ 4) + (0.0224 * n0 ^ 3) - (0.00723 * n0 ^ 2) + (0.00444 * n0) - 0.00745) * ((alphaR ^ 3) - (2 * pi * alphaR ^ 2) - (10 * alphaR)) + frictionFactor

End Function

Function oxygenConcentration(measuredConcentration As Single, O2refValue As Single, measuredO2percent As Single) As Single

oxygenConcentration = measuredConcentration * ((20.9 - O2refValue) / (20.9 - measuredO2percent))

End Function

Function nitrogenConcentration(measuredO2ppm As Single, measuredO2percent As Single) As Single

nitrogenConcentration = measusredO2ppm * ((20.99 - 3) / (20.99 - measuredO2percent))

End Function

Function valveSize(Cv As Single, range As String) As String

Dim i As Integer
If range = "60-90%" Then
    For i = 7 To 58
        If Cv > Worksheets("Cv Tables").Cells(i, "J").Value And Cv < Worksheets("Cv Tables").Cells(i, "M").Value Then
            valveSize = Worksheets("Cv Tables").Cells(i, 3)
            Exit For
        Else
            valveSize = "Cv outside bounds"
        End If
    Next i

ElseIf range = "20-40%" Then
    For i = 7 To 58
            If Cv > Worksheets("Cv Tables").Cells(i, "F").Value And Cv < Worksheets("Cv Tables").Cells(i, "H").Value Then
                valveSize = Worksheets("Cv Tables").Cells(i, 3)
                Exit For
            Else
                valveSize = "Cv outside bounds"
            End If
        Next i
End If

End Function


Function BallRotation(Cv, valveSize As String) As Single

Dim CvTable As range
Set CvTable = Worksheets("Cv Tables").range("C7:N58")

Dim y1, y2, x1, x2 As Single
Dim i As Integer

For i = 8 To 10
    If Cv > Application.WorksheetFunction.VLookup(valveSize, CvTable, i, False) And Cv < Application.WorksheetFunction.VLookup(valveSize, CvTable, i + 1, False) Then
        y1 = Application.WorksheetFunction.VLookup(valveSize, CvTable, i, False)
        y2 = Application.WorksheetFunction.VLookup(valveSize, CvTable, i + 1, False)
        x1 = (i - 2) / 10
        x2 = (i - 1) / 10
        Exit For
    End If
Next i


BallRotation = x2 - ((y2 - Cv) * (x2 - x1) / (y2 - y1))

End Function

Function Reinforcement(P As Double, HeaderNom As Single, BranchNom As Single, BranchSCH As String, CA As Double, Material As String, temp As Single)

'P: Internal design pressure
't: Specified vessel wall thickness (not including forming allowances. For pipe it's nominal thickness less manufacturing undertolerance (12.5%))
'tr: Minimum required thickness of shell
'trn: Required thickness of a seamless nozzle wall.
'tn: Branch pipe nominal wall thickness or nozzle wall thickness not including forming allowances.
Dim HeaderTHK, BranchTHK As Single ' pipe nominal thickness
Dim HeaderOD, BranchOD As Single
Dim Ea, Eb As Single 'Joint efficiency
Ea = 0.85
Eb = 1
Dim Sv As Double 'allowable stress in vessel
Dim Sn As Double 'allowable stress in nozzle
Dim k As Integer
Dim tmin, tcmin, tcActual, st, leg As Single
Dim WeldSizeArray As Variant
WeldSizeArray = Array(0.1875, 0.25, 0.375, 0.5, 0.625, 0.75)
Dim T, tn, tr, trn As Double
Dim a, A1, A1a, A1B, result1, result2, result3, result4, result5 As Double
Dim A2, A2a, A2b, A42 As Double
Dim f, E1, fr1, fr2, fr3 As Double
f = 1
E1 = 1
fr1 = 1
Sv = SecVIIIAllowableStress(Material, temp)
Sn = Sv
fr2 = Sn / Sv
fr3 = Sn / Sv

'HeaderTHK = PipeWall(HeaderNom, HeaderSCH)
BranchTHK = PipeWall(BranchNom, BranchSCH)
HeaderOD = PipeOD(HeaderNom)
BranchOD = PipeOD(BranchNom)

Dim d As Double 'branch ID, finished diameter of circular opening.
d = BranchOD - (2 * (BranchTHK - CA))

Dim THK, THKNom As Single
Dim HeaderSCH As String

THK = P * HeaderOD / 2 / (Sv * Ea + 0.4 * P)

Dim THKArray As Variant
Dim SCH As Variant
SCH = Array("SCH5S", "SCH10S", "SCH10", "SCH20", "SCH30", "SCHSTD", "SCH40", "SCH60", "SCHXS", "SCH80", "SCH100", "SCH120", "SCH140", "CH160", "SCHXXS")
Select Case HeaderNom
    Case 1
        THKArray = Array(0.065, 0.109, 0, 0, 0, 0.133, 0.133, 0, 0.179, 0.179, 0, 0, 0, 0.25, 0.358, 10)
    Case 1.5
        THKArray = Array(0.065, 0.109, 0, 0, 0, 0.145, 0.145, 0, 0.2, 0.2, 0, 0, 0, 0.281, 0.4, 10)
    Case 2
        THKArray = Array(0.065, 0.109, 0, 0, 0, 0.154, 0.154, 0, 0.218, 0.218, 0, 0, 0, 0.344, 0.436, 10)
    Case 2.5
        THKArray = Array(0.083, 0.12, 0, 0, 0, 0.203, 0.203, 0, 0.276, 0.276, 0, 0, 0, 0.375, 0.552, 10)
    Case 3
        THKArray = Array(0.083, 0.12, 0, 0, 0, 0.216, 0.216, 0, 0.3, 0.3, 0, 0, 0, 0.438, 0.6, 10)
    Case 3.5
        THKArray = Array(0.083, 0.12, 0, 0, 0, 0.226, 0.226, 0, 0.318, 0.318, 0, 0, 0, 0, 0, 10)
    Case 4
        THKArray = Array(0.083, 0.12, 0, 0, 0, 0.237, 0.237, 0, 0.337, 0.337, 0, 0.438, 0, 0.531, 0.674, 10)
    Case 5
        THKArray = Array(0.109, 0.134, 0, 0, 0, 0.258, 0.258, 0, 0.375, 0.375, 0, 0.5, 0, 0.625, 0.75, 10)
    Case 6
        THKArray = Array(0.109, 0.134, 0, 0, 0, 0.28, 0.28, 0, 0.432, 0.432, 0, 0.562, 0, 0.719, 0.864, 10)
    Case 8
        THKArray = Array(0.109, 0.148, 0, 0.25, 0.277, 0.322, 0.322, 0.406, 0.5, 0.5, 0.594, 0.719, 0.812, 0.906, 0.875, 10)
    Case 10
        THKArray = Array(0.134, 0.165, 0, 0.25, 0.307, 0.365, 0.365, 0.5, 0.5, 0.594, 0.719, 0.844, 1, 1.125, 1, 10)
    Case 12
        THKArray = Array(0.156, 0.18, 0, 0.25, 0.33, 0.375, 0.406, 0.562, 0.5, 0.688, 0.844, 1, 1.125, 1.312, 1, 10)
    Case 14
        THKArray = Array(0.156, 0.25, 0.25, 0.312, 0.378, 0.375, 0.438, 0.594, 0.5, 0.75, 0.938, 1.094, 1.25, 1.406, 0, 10)
    Case 16
        THKArray = Array(0.165, 0.25, 0.25, 0.312, 0.375, 0.375, 0.5, 0.656, 0.5, 0.844, 1.031, 1.219, 1.438, 1.594, 0, 10)
    Case 18
        THKArray = Array(0.165, 0.25, 0.25, 0.312, 0.438, 0.375, 0.562, 0.75, 0.5, 0.98, 1.156, 1.375, 1.562, 1.781, 0, 10)
    Case 20
        THKArray = Array(0.188, 0.25, 0.25, 0.375, 0.5, 0.375, 0.594, 0.812, 0.5, 1.031, 1.281, 1.5, 1.75, 1.969, 0, 10)
    Case 22
        THKArray = Array(0.188, 0.25, 0.25, 0.375, 0.5, 0.375, 0, 0.875, 0.5, 1.125, 1.375, 1.625, 1.875, 2.125, 0, 10)
    Case 24
        THKArray = Array(0.218, 0.25, 0.25, 0.375, 0.562, 0.375, 0.688, 0.969, 0.5, 1.218, 1.531, 1.812, 2.062, 2.344, 0, 10)
    Case 26
        THKArray = Array(0, 0, 0.312, 0.5, 0, 0.375, 0, 0, 0.5, 0, 0, 0, 0, 0, 0, 10)
    Case 28
        THKArray = Array(0, 0, 0.312, 0.5, 0.625, 0.375, 0, 0, 0.5, 0, 0, 0, 0, 0, 0, 10)
    Case 30
        THKArray = Array(0.25, 0.312, 0.312, 0.5, 0.625, 0.375, 0, 0, 0.5, 0, 0, 0, 0, 0, 0, 10)
    Case 32
        THKArray = Array(0, 0, 0.312, 0.5, 0.625, 0.375, 0.688, 0, 0.5, 0, 0, 0, 0, 0, 0, 10)
    Case 34
        THKArray = Array(0, 0, 0.312, 0.5, 0.625, 0.375, 0.688, 0, 0.5, 0, 0, 0, 0, 0, 0, 10)
    Case 36
        THKArray = Array(0, 0, 0.312, 0.5, 0.625, 0.375, 0.75, 0, 0.5, 0, 0, 0, 0, 0, 0, 10)
    Case 42
        THKArray = Array(0, 0, 0, 0.375, 0, 0, 0, 0, 0.5, 0, 0, 0, 0, 0, 0, 10)
             
End Select

Dim i, j, q As Integer
For q = 1 To 16
    If THK < THKArray(q) Then
        If THKArray(q) = 10 Then
        Reinforcement = "Reinforcement pad Needed"
        Exit Function
        Else
        THKNom = THKArray(q)
        HeaderSCH = SCH(q)
        End If
    Exit For
    End If
Next
For j = 1 To 10

    T = (THKNom * 0.875) - CA

    tn = BranchTHK - CA
    tr = P * HeaderOD / 2 / (Sv * Eb + 0.4 * P) 'Appendix 1-1, cylindrical shells (circumferential stress)
    trn = P * d / 2 / (Sn * Eb - 0.6 * P)

    a = (d * tr * f) + 2 * (tn * tr * f * (1 - fr1))
    A1a = d * (E1 * T - f * tr) - (2 * tn * (E1 * T - f * tr) * (1 - fr1))
    A1B = 2 * (T + tn) * (E1 * T - f * tr) - 2 * tn * (E1 * T - f * tr) * (1 - fr1)

    If A1a > A1B Then
        result1 = A1a
        Else
        result1 = A1B
    End If
    A1 = result1
    A2a = 5 * (tn - trn) * fr2 * T
    A2b = 5 * (tn - trn) * fr2 * tn
    If A2a < A2b Then
        result2 = A2a
        Else
        result2 = A2b
    End If
    A2 = result2
   'Fillet weld size:
    If T < 0.75 Then
        result3 = T
        Else
        result3 = 0.75
    End If
    If tn < result3 Then
        result4 = tn
        Else
        result4 = result3
    End If
    tmin = result4
    st = 0.7 * tmin
    If st < 0.25 Then
        result5 = st
        Else
        result5 = 0.25
    End If
    tcmin = result5
    For i = 1 To 6
        For k = 1 To 6
            If tcmin < WeldSizeArray(k) Then
            leg = WeldSizeArray(k)
            Exit For
        End If
    Next
    tcActual = 0.7 * leg
        If tcmin < tcActual Then
            Exit For
        End If
    Next

    A42 = (leg ^ 2) * fr2
Dim Aavailable As Single
Aavailable = A1 + A2 + A42
    If Aavailable < a Then
        THKNom = THKArray(q + 1)
        q = q + 1
        HeaderSCH = SCH(q)
        Else
        Reinforcement = HeaderSCH
        Exit For
    End If
Next
End Function

Function HC2DesignDuty(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[DesignDuty]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    HC2DesignDuty = Recordset(0)
    Else: HC2DesignDuty = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function InnerCoilDia(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[InnerCoilDia]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerCoilDia = Recordset(0)
    Else: InnerCoilDia = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function InnerCoilSplits(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[InnerCoilSplits]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerCoilSplits = Recordset(0)
    Else: InnerCoilSplits = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function InnerCoilPitch(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[InnerCoilPitch]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerCoilPitch = Recordset(0)
    Else: InnerCoilPitch = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function InnerCoilTurns(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[InnerCoilTurns]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerCoilTurns = Recordset(0)
    Else: InnerCoilTurns = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function OuterCoilDia(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[OuterCoilDia]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterCoilDia = Recordset(0)
    Else: OuterCoilDia = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function OuterCoilSplits(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[OuterCoilSplits]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterCoilSplits = Recordset(0)
    Else: OuterCoilSplits = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function OuterCoilPitch(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[OuterCoilPitch]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterCoilPitch = Recordset(0)
    Else: OuterCoilPitch = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function OuterCoilTurns(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[OuterCoilTurns]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterCoilTurns = Recordset(0)
    Else: OuterCoilTurns = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function HC2DesignFlow(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[DesignFlow]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    HC2DesignFlow = Recordset(0)
    Else: HC2DesignFlow = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function ShellID(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[ShellID]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellID = Recordset(0)
    Else: ShellID = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function ShellThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[ShellThickness]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellThickness = Recordset(0)
    Else: ShellThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function ShellLength(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Standard Coils].[ShellLength]"
'define table
SQL = SQL & "FROM [HC2 Standard Coils]"
'constraint #1
SQL = SQL & "WHERE [HC2 Standard Coils].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellLength = Recordset(0)
    Else: ShellLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function ShellBottomPlateThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[BottomPlateThickness]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellBottomPlateThickness = Recordset(0)
    Else: ShellBottomPlateThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function ShellTopPlateThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[TopPlateThickness]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellTopPlateThickness = Recordset(0)
    Else: ShellTopPlateThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellSupportType(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[ShellSupportType]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellSupportType = Recordset(0)
    Else: ShellSupportType = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellSupportSize(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[ShellSupportSize]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellSupportSize = Recordset(0)
    Else: ShellSupportSize = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellSupportFlatbarThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[ShellSupportFlatbarThickness]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellSupportFlatbarThickness = Recordset(0)
    Else: ShellSupportFlatbarThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellSupportNumber(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[ShellSupportNumber]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellSupportNumber = Recordset(0)
    Else: ShellSupportNumber = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellSupportLength(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[ShellSupportLength]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellSupportLength = Recordset(0)
    Else: ShellSupportLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellSupportWeldFactor(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[ShellSupportWeldFactor]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellSupportWeldFactor = Recordset(0)
    Else: ShellSupportWeldFactor = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellBottomPlateFlatbarThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[BottomPlateFlatbarThickness]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellBottomPlateFlatbarThickness = Recordset(0)
    Else: ShellBottomPlateFlatbarThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellBottomPlateFlatbarHeight(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[BottomPlateFlatbarHeight]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellBottomPlateFlatbarHeight = Recordset(0)
    Else: ShellBottomPlateFlatbarHeight = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function InnerCoilInletSleeve(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[InnerCoilInletSleeve]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerCoilInletSleeve = Recordset(0)
    Else: InnerCoilInletSleeve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function InnerCoilOutletSleeve(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[InnerCoilOutletSleeve]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerCoilOutletSleeve = Recordset(0)
    Else: InnerCoilOutletSleeve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function OuterCoilInletSleeve(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[OuterCoilInletSleeve]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterCoilInletSleeve = Recordset(0)
    Else: OuterCoilInletSleeve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function OuterCoilOutletSleeve(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Shell Subassembly].[OuterCoilOutletSleeve]"
'define table
SQL = SQL & "FROM [Shell Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Shell Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterCoilOutletSleeve = Recordset(0)
    Else: OuterCoilOutletSleeve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilType(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[CoilType]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilType = Recordset(0)
    Else: CoilType = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function InnerWrapperRingBolts(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[InnerWrapperRingBolts]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerWrapperRingBolts = Recordset(0)
    Else: InnerWrapperRingBolts = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function InnerLiftingStrapWidth(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[InnerLiftingStrapWidth]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerLiftingStrapWidth = Recordset(0)
    Else: InnerLiftingStrapWidth = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function InnerLiftingStrapThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[InnerLiftingStrapThickness]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerLiftingStrapThickness = Recordset(0)
    Else: InnerLiftingStrapThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilSpacerType(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[SpacerType]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilSpacerType = Recordset(0)
    Else: CoilSpacerType = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilSpacerSize(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[SpacerSize]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilSpacerSize = Recordset(0)
    Else: CoilSpacerSize = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilSpacerNumber(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[SpacerNumber]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilSpacerNumber = Recordset(0)
    Else: CoilSpacerNumber = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilSpacerLength(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[SpacerLength]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilSpacerLength = Recordset(0)
    Else: CoilSpacerLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilSpacerFlatbarThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[SpacerFlatbarThickness]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilSpacerFlatbarThickness = Recordset(0)
    Else: CoilSpacerFlatbarThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilSpacerFlatbarWidth(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[SpacerFlatbarWidth]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilSpacerFlatbarWidth = Recordset(0)
    Else: CoilSpacerFlatbarWidth = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function OuterLiftingStrapWidth(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[OuterLiftingStrapWidth]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterLiftingStrapWidth = Recordset(0)
    Else: OuterLiftingStrapWidth = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function InnerWrapperLength(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[InnerWrapperLength]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerWrapperLength = Recordset(0)
    Else: InnerWrapperLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function InnerWrapperThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[InnerWrapperThickness]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    InnerWrapperThickness = Recordset(0)
    Else: InnerWrapperThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function OuterWrapperLength(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[OuterWrapperLength]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterWrapperLength = Recordset(0)
    Else: OuterWrapperLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function OuterWrapperThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[OuterWrapperThickness]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    OuterWrapperThickness = Recordset(0)
    Else: OuterWrapperThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilLiftingLugs(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[LiftingLugs]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilLiftingLugs = Recordset(0)
    Else: CoilLiftingLugs = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilLiftingLugLength(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[LiftingLugLength]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilLiftingLugLength = Recordset(0)
    Else: CoilLiftingLugLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilLiftingLugThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Coil Subassembly].[LiftingLugThickness]"
'define table
SQL = SQL & "FROM [Coil Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Coil Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilLiftingLugThickness = Recordset(0)
    Else: CoilLiftingLugThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function CrossPlateEndHeight(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Target Plate Subassembly].[CrossPlateEndHeight]"
'define table
SQL = SQL & "FROM [Target Plate Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Target Plate Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CrossPlateEndHeight = Recordset(0)
    Else: CrossPlateEndHeight = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CrossPlateThickness(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Target Plate Subassembly].[CrossPlateThickness]"
'define table
SQL = SQL & "FROM [Target Plate Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Target Plate Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CrossPlateThickness = Recordset(0)
    Else: CrossPlateThickness = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CrossPlateEndLength(HC2Model As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Target Plate Subassembly].[CrossPlateEndLength]"
'define table
SQL = SQL & "FROM [Target Plate Subassembly]"
'constraint #1
SQL = SQL & "WHERE [Target Plate Subassembly].[HeaterModel]='" & HC2Model & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CrossPlateEndLength = Recordset(0)
    Else: CrossPlateEndLength = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Public Function By(vaporFrac, W, liqDensity, vapDensity, tubeID, splits)

Dim Wv, a As Single
Wv = vaporFrac * W
a = (3.14159265 * (tubeID / 12) ^ 2 / 4) * splits
By = 2.16 * Wv / (a * (liqDensity * vapDensity) ^ (1 / 2))


End Function

Public Function Bx(vaporFrac, W, liqDensity, vapDensity, liqViscosity, vapViscosity, tension)

Dim Wv, Wl As Single
Wv = W * vaporFrac
Wl = W - Wv
Bx = 531 * (Wl / Wv) * ((liqDensity * vapDensity) ^ (1 / 2) / liqDensity ^ (2 / 3)) * ((liqViscosity ^ (1 / 3)) / tension)

End Function

Public Function Regime(By, Bx)

Dim Curve1, Curve2, Curve3, Curve4, Curve5, Curve6a, Curve6b As Double

Curve1 = (7448.8364 + (518.6933 * Bx)) / (1 + (0.2382651 * Bx) + (0.00382501 * Bx ^ 2))
Curve2 = 17572.88 * (0.9984239 ^ Bx) * (Bx ^ (-0.6013042))
Curve3 = 17506.68 * (1.022221 ^ Bx) * (Bx ^ (-0.5993874))
Curve4 = 1 / (1.207535E-05 + (8.668342E-06 * Application.WorksheetFunction.ln(Bx)) + (-1.595291E-07 * (Application.WorksheetFunction.ln(Bx)) ^ 3))
Curve5 = (1101.62806079346 + (2.13476359666886 * Bx)) / (1 + (0.00456241507210919 * Bx) + (2.19974405025479E-07 * Bx ^ 2))
Curve6a = (-121578032969.123 + (1132441858.63 * Bx)) / (1 + (2078.38669941481 * Bx) + (14.1169249629428 * Bx ^ 2))
Curve6b = Bx / (-0.0155083596492835 + (0.000110589835680723 * Bx) + (3.66484927002637E-07 * Bx ^ 2))


If Bx < 1300 And By < Curve1 Then
    Regime = "Stratified"

ElseIf Bx < 20 And By > Curve1 And By < Curve2 Then
    Regime = "Wave"
    
ElseIf Bx > 100 And Bx < 5000 And By < Curve5 Then
    Regime = "Plug"
ElseIf Bx > 5000 And By < Curve6b Then
    Regime = "Plug"

ElseIf Bx > 0.8 And Bx < 125 And By > Curve4 Then
    Regime = "Dispersed"
ElseIf Bx > 125 And By > Curve6a Then
    Regime = "Dispersed"
    
ElseIf Bx > 125 And Bx < 180 And By > Curve6b And By < Curve6a Then
    Regime = "Bubble"
ElseIf Bx > 180 And By > Curve6b Then
    Regime = "Bubble"

ElseIf Bx < 125 And By > Curve3 And By < Curve4 Then
    Regime = "Annular"
ElseIf Bx > 125 And Bx < 140 And By > Curve3 And By < Curve4 Then
    Regime = "Annular"

Else: Regime = "Slug"

End If


End Function
Public Function TwoPhasedP(vaporFrac, W, liqDensity As Single, vapDensity As Single, liqViscosity As Single, vapViscosity As Single, tubeID As Single, splits, flowType, Leq)


Dim Wv, Wl, vapor, liquid, zero, a As Single
Wv = W * vaporFrac
Wl = W - Wv
a = (3.14159265 * (tubeID / 12) ^ 2 / 4) * splits

Dim vapFlow As Single
vapor = Wv / a
vapFlow = vapor / vapDensity / 3600

Dim liqFlow As Single
liquid = Wl / a
liqFlow = liquid / liqDensity / 3600

Dim REv As Single
REv = Reynolds(vapFlow, tubeID, vapViscosity, vapDensity)

Dim REl As Single
REl = Reynolds(liqFlow, tubeID, liqViscosity, liqDensity)

Dim f As Double
f = frictionFactorCheng(REv, tubeID)

' redifining:
vapor = Wv / a
vapFlow = vapor / vapDensity / 3600
liquid = Wl / a
liqFlow = liquid / liqDensity / 3600

Dim dPv, dPl, Fl, xSquared, x, Hx, fH As Single
dPv = DPPipe(tubeID, vapFlow, REv, vapDensity, vapDensity, vapViscosity, Leq, 0.00015)
dPl = DPPipe(tubeID, liqFlow, REl, liqDensity, liqDensity, liqViscosity, Leq, 0.00015)
Fl = frictionFactorCheng(REl, tubeID)
x = (Wl / Wv) ^ 0.5 * ((vapDensity / liqDensity) * (Fl / f)) ^ 0.5
xSquared = (x) ^ 2
Hx = (Wl / Wv) * (liqViscosity / vapViscosity)
fH = Exp((0.211 * Log(Hx) - 3.993))

If flowType = "Bubble" Then
    zero = 14.2 * x ^ 0.75 / (Wl / a) ^ 0.1
    TwoPhasedP = zero ^ 2 * dPv

ElseIf flowType = "Plug" Then
    zero = 27.315 * x ^ 0.855 / (Wl / a) ^ 0.17
    TwoPhasedP = zero ^ 2 * dPv
    
ElseIf flowType = "Stratified" Then
    zero = 15400 * x / (Wl / a) ^ 0.8
    TwoPhasedP = zero ^ 2 * dPv

ElseIf flowType = "Slug" Then
    zero = 1190 * x ^ 0.815 / (Wl / a) ^ 0.5
    TwoPhasedP = zero ^ 2 * dPv

ElseIf flowType = "Dispersed" Then
    zero = 10 ^ 0.5
    TwoPhasedP = zero ^ 2 * dPv
    
ElseIf flowType = "Annular" Then
    zero = (4.8 - 0.3125 * tubeID) * x ^ (0.343 - 0.021 * tubeID)
    TwoPhasedP = zero ^ 2 * dPv

ElseIf flowType = "Wave" Then
    zero = 15400 * x / (Wl / a) ^ 0.8
    TwoPhasedP = zero ^ 2 * dPv
    
Else: TwoPhasedP = "define Regime"

End If

End Function
Public Function frictionFactorCheng(Re As Single, diameter As Single)

Dim a As Double
Dim b As Double
Dim ftdia As Single
'Diameter in inches
'Reynolds number is unitless

ftdia = diameter / 12
a = 1 / (1 + (Re / 2720) ^ 9)
b = 1 / (1 + (Re / (160 * (ftdia / 0.00015))) ^ 2)

frictionFactorCheng = 1 / ((Re / 64) ^ a * (1.8 * Application.WorksheetFunction.Log10(Re / 6.8)) ^ (2 * (1 - a) * b) * (2 * Application.WorksheetFunction.Log10(3.7 * ftdia / 0.00015)) ^ (2 * (1 - a) * (1 - b)))
End Function
-------------------------------------------------------------------------------
VBA MACRO FluidFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/FluidFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Function DowAVaporPressure(Temperature As Single)
Dim TempArray, VaporPressureArray As Variant

TempArray = Array(60, 120, 300, 360, 420, 480, 540, 600, 660, 720, 780)
VaporPressureArray = Array(0, 0.003, 0.64, 2.03, 5.38, 12.25, 24.72, 45.31, 76.89, 122.7, 186.4)
'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
    If VaporPressureArray(i + 1) = -1 Then
        DowAVaporPressure = "Beyond Fluid Limits"
        Exit Function
    End If
    DowAVaporPressure = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporPressureArray(i + 1) - VaporPressureArray(i)) + VaporPressureArray(i)
    i = i + 1
Loop

End Function
Function DowAVaporViscosity(Temperature As Single)
Dim TempArray, VaporViscosityArray As Variant

TempArray = Array(60, 120, 300, 360, 420, 480, 540, 600, 660, 720, 780)
VaporViscosityArray = Array(0.0054, 0.006, 0.0079, 0.0086, 0.0092, 0.0098, 0.0105, 0.0113, 0.0121, 0.013, 0.0142)

'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
    If VaporViscosityArray(i + 1) = -1 Then
        DowAVaporViscosity = "Beyond Fluid Limits"
        Exit Function
    End If
    DowAVaporViscosity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporViscosityArray(i + 1) - VaporViscosityArray(i)) + VaporViscosityArray(i)
    i = i + 1
Loop
End Function

Function DowALiquidEnthalpy(Temperature As Single)
Dim TempArray, LiquidEnthalpyArray As Variant

TempArray = Array(60, 120, 300, 360, 420, 480, 540, 600, 660, 720, 780)
LiquidEnthalpyArray = Array(2.5, 26.2, 103, 131.1, 160.6, 191.4, 223.5, 256.9, 291.7, 327.9, 365.9)

'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
    If LiquidEnthalpyArray(i + 1) = -1 Then
        DowALiquidEnthalpy = "Beyond Fluid Limits"
        Exit Function
    End If
    DowALiquidEnthalpy = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (LiquidEnthalpyArray(i + 1) - LiquidEnthalpyArray(i)) + LiquidEnthalpyArray(i)
    i = i + 1
Loop
End Function
Function DowAVaporDensity(Temperature As Single)
Dim TempArray, VaporDensityArray As Variant

TempArray = Array(60, 120, 300, 360, 420, 480, 540, 600, 660, 720, 780)
VaporDensityArray = Array(0, 0, 0.013, 0.0388, 0.0967, 0.21, 0.4102, 0.7389, 1.254, 2.045, 3.27)

'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
    If VaporDensityArray(i + 1) = -1 Then
        DowAVaporDensity = "Beyond Fluid Limits"
        Exit Function
    End If
    DowAVaporDensity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporDensityArray(i + 1) - VaporDensityArray(i)) + VaporDensityArray(i)
    i = i + 1
Loop
End Function
Function DowAVaporEnthalpy(Temperature As Single)
Dim TempArray, VaporEnthalpyArray As Variant

TempArray = Array(60, 120, 300, 360, 420, 480, 540, 600, 660, 720, 780)
VaporEnthalpyArray = Array(177.6, 193.5, 251.1, 273.1, 296.3, 320.5, 345.5, 371.1, 397, 422.9, 448.4)

'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
    If VaporEnthalpyArray(i + 1) = -1 Then
        DowAVaporEnthalpy = "Beyond Fluid Limits"
        Exit Function
    End If
    DowAVaporEnthalpy = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporEnthalpyArray(i + 1) - VaporEnthalpyArray(i)) + VaporEnthalpyArray(i)
    i = i + 1
Loop
End Function
Function DowJVaporPressure(Temperature As Single)

Dim TempArray, VaporPressureArray As Variant

TempArray = Array(100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 650)

VaporPressureArray = Array(0.05, 0.1, 0.18, 0.31, 0.52, 0.85, 1.33, 2.02, 2.99, 4.31, 6.06, 8.36, 11.3, 15.02, 19.63, 25.27, 32.1, 40.25, 49.9, 61.2, 74.34, 89.5, 106.87, 126.67, 149.14, 174.52, 203.12, 235.24, 252.74)

'interpolation between tabulated temperature values

i = 0

Do While Temperature > TempArray(i)

If VaporPressureArray(i + 1) = -1 Then

DowJVaporPressure = "Beyond Fluid Limits"

Exit Function

End If

DowJVaporPressure = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporPressureArray(i + 1) - VaporPressureArray(i)) + VaporPressureArray(i)

i = i + 1

Loop

End Function
Function DowJLiquidEnthalpy(Temperature As Single)

Dim TempArray, LiquidEnthalpyArray As Variant

TempArray = Array(100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 650)

LiquidEnthalpyArray = Array(9.7, 18.4, 27.3, 36.6, 46, 55.7, 65.7, 75.8, 86.2, 96.8, 107.6, 118.6, 129.9, 141.3, 152.9, 164.7, 176.6, 188.8, 201.1, 213.7, 226.4, 239.4, 252.5, 265.9, 279.5, 293.4, 307.6, 322.2, 329.7)

'interpolation between tabulated temperature values

i = 0

Do While Temperature > TempArray(i)

If LiquidEnthalpyArray(i + 1) = -1 Then

DowJLiquidEnthalpy = "Beyond Fluid Limits"

Exit Function

End If

DowJLiquidEnthalpy = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (LiquidEnthalpyArray(i + 1) - LiquidEnthalpyArray(i)) + LiquidEnthalpyArray(i)

i = i + 1

Loop

End Function
Function DowJVaporDensity(Temperature As Single)

Dim TempArray, VaporDensityArray As Variant

TempArray = Array(100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 650)

VaporDensityArray = Array(0.0011, 0.002, 0.0037, 0.0063, 0.0103, 0.0163, 0.0249, 0.0369, 0.0532, 0.0749, 0.1031, 0.1392, 0.1847, 0.2413, 0.3109, 0.3955, 0.4976, 0.6199, 0.7656, 0.9383, 1.143, 1.384, 1.669, 2.007, 2.409, 2.892, 3.48, 4.21, 4.648)

'interpolation between tabulated temperature values

i = 0

Do While Temperature > TempArray(i)

If VaporDensityArray(i + 1) = -1 Then

DowJVaporDensity = "Beyond Fluid Limits"

Exit Function

End If

DowJVaporDensity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporDensityArray(i + 1) - VaporDensityArray(i)) + VaporDensityArray(i)

i = i + 1

Loop

End Function
Function DowJVaporEnthalpy(Temperature As Single)

Dim TempArray, VaporEnthalpyArray As Variant

TempArray = Array(100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 650)

VaporEnthalpyArray = Array(171.4, 178.1, 185, 192.1, 199.4, 206.9, 214.5, 222.4, 230.4, 238.5, 246.8, 255.2, 263.8, 272.4, 281.2, 290, 298.9, 307.9, 316.9, 325.9, 335, 344, 353, 362, 370.8, 379.5, 388, 396.1, 400)

'interpolation between tabulated temperature values

i = 0

Do While Temperature > TempArray(i)

If VaporEnthalpyArray(i + 1) = -1 Then

DowJVaporEnthalpy = "Beyond Fluid Limits"

Exit Function

End If

DowJVaporEnthalpy = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporEnthalpyArray(i + 1) - VaporEnthalpyArray(i)) + VaporEnthalpyArray(i)

i = i + 1

Loop

End Function
Function DowJVaporViscosity(Temperature As Single)

Dim TempArray, VaporViscosityArray As Variant

TempArray = Array(100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 650)

VaporViscosityArray = Array(0.006, 0.006, 0.006, 0.007, 0.007, 0.007, 0.007, 0.007, 0.008, 0.008, 0.008, 0.008, 0.009, 0.009, 0.009, 0.009, 0.009, 0.01, 0.01, 0.01, 0.011, 0.011, 0.011, 0.012, 0.012, 0.012, 0.013, 0.013, 0.014)

'interpolation between tabulated temperature values

i = 0

Do While Temperature > TempArray(i)

If VaporViscosityArray(i + 1) = -1 Then

DowJVaporViscosity = "Beyond Fluid Limits"

Exit Function

End If

DowJVaporViscosity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (VaporViscosityArray(i + 1) - VaporViscosityArray(i)) + VaporViscosityArray(i)

i = i + 1

Loop

End Function
Function FluidDensity(Fluid As String, Temperature As Single)

Dim TempArray, FluidDensityArray As Variant

TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)


'all of the fluids go here
    Select Case Fluid
        Case "Shell Thermia C"
            FluidDensityArray = Array(55.06016, 55.06016, 54.01282, 52.89067, 51.76852, 50.64637, 49.52422, 48.40207, 47.27992, 46.08296, 44.96081, 43.83866, 42.6417)
        Case "Shell Thermia B"
            FluidDensityArray = Array(55.06016, 55.06016, 54.01282, 52.89067, 51.76852, 50.64637, 49.52422, 48.40207, 47.27992, 46.08296, 44.96081, 43.83866, 42.6417)
        Case "Calflo AF"
            FluidDensityArray = Array(53.8596, 53.8596, 52.3635, 51.61545, 50.11935, 48.62325, 47.8752, 46.3791, 45.63105, 44.13495, 43.3869, 41.8908, 40.3947)
        Case "Calflo FG"
            FluidDensityArray = Array(54, 53.11155, 52.3635, 50.8674, 50.11935, 48.62325, 47.8752, 46.3791, 45.63105, 44.13495, 42.63885, 41.8908, 40.3947, 38.896)
        Case "Calflo HTF"
            FluidDensityArray = Array(55.6, 54.5, 53.4, 52.3, 51.2, 50.1, 49, 47.9, 46.8, 45.7, 44.6, 43.5, 42.4)
        Case "Calflo LT"
            FluidDensityArray = Array(51, 50.8674, 50.11935, 48.62325, 47.8752, 46.3791, 44.883, 44.13495, 42.63885, 41.8908, 40.3947, 39.64665)
        Case "Dowtherm A"
            FluidDensityArray = Array(67.3245, 67.3245, 65.27, 63.88, 62.46, 61#, 59.5, 57.96, 56.37, 54.72, 53#, 51.2, 49.29, 47.25, 45.03, 42.57, 39.74)
        Case "Dowtherm G"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 725)
            FluidDensityArray = Array(65.91, 64.57, 63.22, 61.88, 60.53, 59.19, 57.84, 56.5, 55.15, 53.81, 52.46, 51.12, 49.77, 48.43, 47.75)
        Case "Dowtherm Q"
            FluidDensityArray = Array(62.05, 61.12, 59.79, 58.47, 57.14, 55.82, 54.49, 53.17, 51.84, 50.52, 49.2, 47.87, 46.55, 45.23)
        Case "Dowtherm RP"
            FluidDensityArray = Array(65.25, 64.76, 63.545, 62.33, 61.105, 59.88, 58.63, 57.38, 56.1, 54.82, 53.5, 52.18, 50.795, 49.41, 47.94)
        Case "Dowtherm HT"
            FluidDensityArray = Array(62.93, 62.93, 61.72, 60.5, 59.29, 58.07, 56.86, 55.65, 54.43, 53.22, 52, 50.79, 49.58, 48.36, 47.15)
        Case "Dowtherm J"
            FluidDensityArray = Array(55, 54.385, 53.09, 51.725, 50.36, 48.88, 47.4, 45.755, 44.11, 42.205, 40.3, 37.88, 35.46, 32.16)
        Case "Dowtherm MX"
            FluidDensityArray = Array(63, 61.9, 59.25, 58.1, 56.8, 55.52, 54.23, 52.9, 51.57, 50.17, 48.7, 47.2, 45.67, 44, 42.333)
        Case "Dowtherm T"
            TempArray = Array(20, 100, 180, 260, 340, 420, 500, 580, 600)
            FluidDensityArray = Array(55.66, 53.75, 51.84, 49.93, 48.02, 46.11, 44.19, 42.28, 41.8)
        Case "Mobiltherm 603"
            FluidDensityArray = Array(55.3928, 54.00175, 52.96125, 51.87567, 50.88026, 49.8134, 48.73823, 47.69634, 46.65585, 45.61535, 44.57485, 43.53436, 42.4453, -1, -1, -1, -1)
        Case "Therminol 55"
            FluidDensityArray = Array(56, 54.85, 53.7, 52.5, 51.3, 50.15, 49, 47.75, 46.5, 45.25, 43.9, 42.55, 41.1, 40.018)
        Case "Therminol 59"
            FluidDensityArray = Array(62.5, 61.25, 60, 58.75, 57.5, 56.2, 54.9, 53.55, 52.2, 50.8, 49.3, 47.8, 46.2, 44.88)
        Case "Therminol 66"
            FluidDensityArray = Array(64.52, 63.35, 62.2, 61.05, 59.9, 58.75, 57.5, 56.3, 55.1, 53.75, 52.5, 51.1, 49.7, 48.2, 46.6)
        Case "Therminol SP"
            FluidDensityArray = Array(56, 54.6, 53.66033554, 52.50888443, 51.3, 50.15, 49, 47.75, 46.50192261, 45.25, 43.949312, 42.55, 41.11230469, -1, -1, -1, -1)
        Case "Therminol D-12"
            FluidDensityArray = Array(49.3, 48.05, 46.8, 45.55, 44.2, 42.8, 41.4, 39.8, 38.2, 36.3, 34.3, 32.2)
        Case "Therminol VP-1"
            FluidDensityArray = Array(67.7, 66.87, 65.5, 64.1, 62.7, 61.25, 59.8, 58.3, 56.8, 55.2, 53.5, 51.8, 50, 48, 45.9, 43.4, 40.6)
        Case "Therminol 62"
            FluidDensityArray = Array(61.3, 60.04999924, 58.8, 57.54999924, 56.3, 55.04999924, 53.7, 52.34999847, 50.9, 49.5, 48, 46.34999847, 44.6, 43.9, 43.1, 42.82267076)
        Case "Therminol 72"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 715)
            FluidDensityArray = Array(69.515, 67.95, 66.4, 64.8, 63.2, 61.65, 60.1, 58.5, 56.9, 55.35, 53.8, 52.2, 50.6, 49.1, 47.5, 47)
        Case "Therminol 75"
            TempArray = Array(0, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700, 725)
            FluidDensityArray = Array(69.4, 68.15, 67.525, 66.9, 66.275, 65.65, 65.025, 64.4, 63.775, 63.15, 62.525, 61.9, 61.275, 60.6, 59.925, 59.3, 58.65, 57.95, 57.25, 56.6, 55.85, 55.1, 54.35, 53.6, 52.85, 52.05, 51.25, 50.4, 49.6)
        Case "Therminol XP"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidDensityArray = Array(56.3, 55.2, 54.1, 53, 51.9, 50.8, 49.7, 48.5, 47.3, 46.15, 44.9, 43.65, 42.3)
        Case "Paratherm NF"
            FluidDensityArray = Array(55.31, 54.6077, 53.8596, 53.1116, 52.3635, 51.6903, 51.08, 50.36, 49.63, 48.9, 48.17, 47.45, 46.72, 46.002)
        Case "Paratherm MG"
            FluidDensityArray = Array(51.6, 50.1, 48.5, 47, 45.4, 43.9, 42.4, 40.8, 39.3, 37.7, 36.2, -1, -1, -1, -1, -1, -1)
        Case "Paratherm MR"
            FluidDensityArray = Array(52.21, 50.81, 49.42, 48.03, 46.64, 45.25, 43.85, 42.46, 41.07, 39.68, 38.29, 36.89, 35.59)
        Case "Paratherm HE"
            FluidDensityArray = Array(55.34, 54.28, 53.17, 52.06, 50.95, 49.83, 48.72, 47.61, 46.5, 45.39, 44.28, 43.16, 42.05, 41.79, 41.56, 41.34, 41.12, 40.91848)
        Case "Paratherm HR"
            TempArray = Array(0, 25, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700)
            FluidDensityArray = Array(62, 61, 61, 60, 59, 59, 58, 57, 57, 56, 56, 55, 54, 54, 53, 52, 52, 51, 51, 50, 49, 49, 48, 47, 47, 46, 46, 45, 44)
        Case "Petrotherm"
            FluidDensityArray = Array(54.9, 54.9, 53.4, 52.5, 51.1, 49.6, 48.8, 47.3, 46.5, 45, 44.09, 42.93, 41.77)
        Case "PetroCanada FG"
            FluidDensityArray = Array(53.3, 53.3, 52, 51, 50, 48, 46, 45.54999924, 45.1, 44.29999924, 43.5, 42, 40.5)
        Case "Syltherm 800"
            FluidDensityArray = Array(60, 59.005, 57.44, 55.905, 54.37, 52.825, 51.28, 49.675, 48.07, 46.345, 44.62, 42.73, 40.84, 38.72, 36.6, 34.205, 31.81)
        Case "TEG 0.6"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidDensityArray = Array(69.84473, 68.21669, 67.22503, 65.96065, 64.5513, 63.05228, 61.48027, 61.10832, 59.86032)
        Case "TEG 0.98"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidDensityArray = Array(72.17655984, 70.75329097, 69.32001136, 67.97056072, 66.57249608, 65.13477008, 63.66580776, 62.29666908, 58.8421)
        Case "EGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(66.68, 66.55, 66.41, 66.27, 66.11, 65.96, 65.79, 65.62, 65.45, 65.27, 65.08, 64.88, 64.68, 64.48, 64.27, 64.05, 63.82, 63.59, 63.36, 63.11, 62.87, 62.61, 62.35, 62.08, 61.81, 61.53, 61.25, 60.96, 60.66, 60.36, 60.05, 59.73, 59.41, 59.08, 58.75)
        Case "EGlycol 0.4"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(67.93, 67.79, 67.64, 67.49, 67.33, 67.17, 66.99, 66.82, 66.63, 66.44, 66.25, 66.05, 65.84, 65.63, 65.41, 65.18, 64.95, 64.72, 64.47, 64.22, 63.97, 63.71, 63.44, 63.17, 62.89, 62.6, 62.31, 62.02, 61.71, 61.41, 61.09, 60.77, 60.44, 60.11, 59.77, 59.43)
        Case "EGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(69.26, 69.12, 68.97, 68.82, 68.66, 68.49, 68.32, 68.14, 67.96, 67.77, 67.58, 67.38, 67.17, 66.96, 66.74, 66.51, 66.28, 66.05, 65.8, 65.56, 65.3, 65.04, 64.78, 64.51, 64.23, 63.95, 63.66, 63.36, 63.06, 62.76, 62.44, 62.13, 61.8, 61.47, 61.14, 60.8, 60.45, 60.1)
        Case "Dynalene EG-XT 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 120, 140, 160, 180, 200, 220)
            FluidDensityArray = Array(69.26, 69.12, 68.97, 68.82, 68.66, 68.49, 68.32, 68.14, 67.96, 67.77, 67.58, 67.38, 67.17, 66.74, 66.28, 65.8, 65.3, 64.78, 64.23)
        Case "EGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(70.4, 70.26, 70.1, 69.94, 69.78, 69.6, 69.43, 69.24, 69.06, 68.86, 68.66, 68.46, 68.25, 68.03, 67.81, 67.58, 67.34, 67.1, 66.86, 66.61, 66.35, 66.09, 65.82, 65.54, 65.27, 64.98, 64.69, 64.39, 64.09, 63.78, 63.47, 63.15, 62.82, 62.49, 62.15, 61.81, 61.46, 61.11, 60.75)
        Case "EGlycol 0.7"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(71.33, 71.17, 71.01, 70.84, 70.66, 70.48, 70.29, 70.1, 69.9, 69.7, 69.49, 69.28, 69.06, 68.83, 68.6, 68.36, 68.12, 67.87, 67.62, 67.36, 67.1, 66.83, 66.55, 66.27, 65.98, 65.69, 65.39, 65.09, 64.78, 64.47, 64.15, 63.82, 63.49, 63.16, 62.81, 62.47, 62.11, 61.76, 61.39)
        Case "EGlycol 0.8"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(71.67, 71.49, 71.3, 71.1, 70.9, 70.7, 70.48, 70.27, 70.05, 69.82, 69.59, 69.35, 69.1, 68.85, 68.6, 68.34, 68.08, 67.81, 67.53, 67.25, 66.96, 66.67, 66.37, 66.07, 65.76, 65.45, 65.13, 64.81, 64.48, 64.15, 63.81, 63.46, 63.11, 62.75, 62.39, 62.03)
        Case "EGlycol 0.9"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(72.45, 72.26, 72.06, 71.86, 71.65, 71.44, 71.22, 70.99, 70.77, 70.53, 70.29, 70.05, 69.8, 69.55, 69.29, 69.02, 68.75, 68.48, 68.2, 67.91, 67.62, 67.33, 67.03, 66.72, 66.41, 66.1, 65.78, 65.45, 65.12, 64.78, 64.44, 64.09, 63.74, 63.38, 63.02, 62.66)
        Case "PGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidDensityArray = Array(65, 64.9, 64.79, 64.67, 64.53, 64.39, 64.24, 64.08, 63.91, 63.73, 63.54, 63.33, 63.12, 62.9, 62.67, 62.43, 62.18, 61.92, 61.65, 61.37, 61.08, 61.5, 61.2, 60.89, 60.57, 60.24, 59.91, 59.56, 59.2, 58.84, 58.46, 58.08, 57.89)
        Case "PGlycol 0.4"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidDensityArray = Array(65.71, 65.6, 65.48, 65.35, 65.21, 65.06, 64.9, 64.73, 64.55, 64.36, 64.16, 63.95, 63.74, 63.51, 63.27, 63.02, 62.76, 62.49, 62.22, 61.93, 61.63, 61.32, 61, 61.61, 61.29, 60.96, 60.61, 60.26, 59.91, 59.54, 59.16, 58.77, 58.38, 58.18)
        Case "PGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(66.46, 66.35, 66.23, 66.11, 65.97, 65.82, 65.67, 65.5, 65.33, 65.14, 64.95, 64.74, 64.53, 64.3, 64.06, 63.82, 63.57, 63.3, 63.03, 62.74, 62.45, 62.14, 61.83, 61.5, 61.17, 61.95, 61.62, 61.27, 60.92, 60.56, 60.19, 59.81, 59.42, 59.02, 58.62, 58.41)
         Case "PGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidDensityArray = Array(67.05, 66.93, 66.81, 66.68, 66.54, 66.38, 66.22, 66.05, 65.87, 65.68, 65.47, 65.26, 65.04, 64.81, 64.57, 64.32, 64.06, 63.79, 63.51, 63.22, 62.92, 62.61, 62.29, 61.97, 61.63, 61.28, 60.92, 61.87, 61.52, 61.16, 60.78, 60.4, 60.02, 59.62, 59.22, 58.8, 58.59)
        Case "PGlycol 0.9"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidDensityArray = Array(68, 67.75, 67.49, 67.23, 66.97, 66.71, 66.44, 66.18, 65.91, 65.64, 65.37, 65.09, 64.82, 64.54, 64.26, 63.98, 63.7, 63.42, 63.13, 62.85, 62.56, 62.27, 61.97, 61.68, 61.38, 61.08, 60.78, 60.48)
        Case "Salt"
            TempArray = Array(275, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100)
            FluidDensityArray = Array(125.5321, 124.8495, 123.4844, 122.1192, 120.754, 119.3888, 118.0236, 116.6584, 115.2932, 113.928, 112.5628, 111.1976, 109.8324, 108.4673, 107.1021, 105.7369, 104.3717, 103.0065)
        Case "Xceltherm HT"
            FluidDensityArray = Array(64.282, 65.19, 63.97, 62.75, 61.53, 60.31, 59.08, 57.86, 56.64, 55.42, 54.19, 52.97, 51.75, 50.53, 49.3)
        Case "Xceltherm 600"
            FluidDensityArray = Array(53.55, 53.55, 52.5, 51.44, 50.39, 49.34, 48.29, 47.23, 46.18, 45.13, 44.07, 43.02, 41.97, 41.14)
        Case "Syltherm XLT"
            FluidDensityArray = Array(55.88, 53.295, 51.75, 50.09, 48.43, 46.555, 44.68, 42.5, 40.32, 37.76, 35.2, 32.3884)
        Case "Texatherm 46"
            FluidDensityArray = Array(54.611, 54.611, 53.411, 52.211, 51.011, 49.811, 48.611, 47.411, 46.211, 45.011, 43.811, 42.611, 41.411, 40.211)
        Case "Multitherm PG-1"
            FluidDensityArray = Array(55.44, 54.72, 53.99, 53.26, 52.53, 51.81, 51.08, 50.36, 49.63, 48.9, 48.17, 47.45, 46.72, 46.002)
        Case "Multitherm OG-1"
            TempArray = Array(0, 20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidDensityArray = Array(55.13, 54.98, 54.16, 53.11, 52.21, 51.24, 50.26, 49.37, 48.47, 47.58, 46.6, 45.73, 44.96, 44.21)
        Case "Mobiltherm 43"
            FluidDensityArray = Array(55.9728, 54.9744, 53.7264, 52.4784, 51.2304, 49.9824, 48.7344, 47.4864, 46.2384, 44.9904, 43.7424, 42.4944, 41.2464)
        Case "Sun 21"
            TempArray = Array(150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidDensityArray = Array(52.416, 51.4176, 50.2944, 49.296, 47.9856, 46.8, 45.552, 44.4288, 43.368, 41.3712, 40.70976)
        Case "Marlotherm SH"
            FluidDensityArray = Array(66, 65.6, 64.4, 63.1, 62.2, 60.6, 59.3, 58.2, 56.5, 55.6, 54.5, 53, 52, 50.5, 48.8)
        Case "Marlotherm LH"
            TempArray = Array(-4, 32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572, 608, 644, 680)
            FluidDensityArray = Array(64.1, 63.1, 62.2, 61.2, 60.3, 59.3, 58.4, 57.4, 65.6, 55.6, 54.5, 53.4, 52.4, 51.3, 50.2, 49.1, 47.8, 46.6, 45.3, 43.9)
        Case "Duratherm FG"
            FluidDensityArray = Array(54, 53.516595, 52.395095, 51.273595, 50.152095, 49.030595, 47.909095, 46.787595, 45.666095, 44.544595, 43.423095, 42.301595, 41.180095, 40.058595)
        Case "Noco 21"
            TempArray = Array(50, 150, 300, 450, 600)
            FluidDensityArray = Array(54.182, 52.43952, 49.31812, 45.57244, 41.20248)
        Case "Chemtherm 550"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidDensityArray = Array(57.5, 56.2, 55.2, 54, 53.1, 52, 51, 49.8, 48.8, 47.8, 46.5, 45.4, 44.2)
        Case "Thermalane 550"
            TempArray = Array(0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600)
            FluidDensityArray = Array(54.9, 54.5, 54, 53.6, 53.1, 52.7, 52.2, 51.8, 51.3, 50.9, 50.4, 50, 49.5, 49.1, 48.6, 48.2, 47.7, 47.3, 46.8, 46.4, 45.9, 45.5, 45, 44.6, 44.1, 43.7, 43.2, 42.8, 42.3, 41.9, 41.4)
        Case "Shell S2"
            TempArray = Array(32, 68, 104, 212, 302, 392, 482, 572, 644)
            FluidDensityArray = Array(54.686928, 53.875364, 53.0638, 50.629108, 48.568984, 46.571288, 44.511164, 42.513468, 40.89034)
        Case "Thermoil 100"
            TempArray = Array(0, 100, 212, 392, 500, 600)
            FluidDensityArray = Array(56.6, 54.09, 51.32, 46.78, 44.04, 41.49)
        Case "Uconhtf 500"
            TempArray = Array(0, 100, 200, 300, 400, 500)
            FluidDensityArray = Array(66.6, 64, 61, 58.25, 55.5, 52.5)
        Case "Multitherm IG-1"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)
            FluidDensityArray = Array(54.98, 54.23, 53.19, 52.18, 51.24, 50.26, 49.37, 48.47, 47.58, 46.65, 45.74, 44.95, 44.08)
        Case "Multitherm IG-4"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidDensityArray = Array(54.89, 54.2, 53.14, 52.02, 50.83, 49.65, 48.46, 47.21, 45.97, 44.59, 43.22, 41.79, 40.23)
        Case "Hydroclear O/S 32"
            FluidDensityArray = Array(54.795, 53.905, 53.015, 52.125, 51.235, 50.345, 49.455, 48.565, 47.402, 46.785, 45.816, 45.005, 44.33, 0, 0, 0, 0)
        Case "Hydroclear O/S 46"
            FluidDensityArray = Array(54.751, 53.861, 52.971, 52.081, 51.191, 50.301, 49.411, 48.521, 47.389, 46.741, 45.797, 44.961, 44.318, 0, 0, 0, 0)
        Case "Hydroclear C/S 32"
            FluidDensityArray = Array(55.001, 54.131, 53.261, 52.391, 51.521, 50.651, 49.781, 48.911, 47.832, 47.171, 46.253, 45.431, 44.736, 0, 0, 0, 0)
        Case "Hydroclear C/S 46"
            FluidDensityArray = Array(55.001, 54.131, 53.261, 52.391, 51.521, 50.651, 49.781, 48.911, 47.832, 47.171, 46.253, 45.431, 44.736, 0, 0, 0, 0)
        Case "Xceltherm SST"
            FluidDensityArray = Array(62, 61, 59, 58, 57, 56, 54, 53, 52, 51, 49, 48, 47, 46, 44, 0, 0)
        Case "Paratherm HT"
            TempArray = Array(50, 100, 200, 300, 350, 400, 450, 500, 550)
            FluidDensityArray = Array(64, 62.4, 59.1, 55.8, 54.2, 52.6, 51, 49.4, 47.8)
        Case "Chevron 22"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidDensityArray = Array(54.43896, 53.508753, 52.391256, 51.273759, 50.106318, 48.938877, 47.715249, 46.491621, 45.1961985, 43.900776, 42.2527316, 41.153856, 39.7)
        Case "Mobiltherm 605"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidDensityArray = Array(54.628, 53.503, 52.378, 51.253, 50.128, 49.003, 47.878, 46.753, 45.628, 44.503, 43.378, 42.253)
        Case "Duratherm 600"
            TempArray = Array(15, 55, 95, 145, 195, 245, 295, 345, 395, 445, 495, 545, 600)
            FluidDensityArray = Array(53.63, 52.72, 51.8, 50.66, 49.52, 48.37, 47.23, 46.08, 44.94, 43.8, 42.65, 41.51, 40.25)
        Case "Chevron 46"
            TempArray = Array(32, 104, 122, 212, 302, 392, 482, 572, 662)
            FluidDensityArray = Array(54.5688, 52.8715, 52.572, 50.575, 48.441, 46.276, 43.88, 41.683, 38.782)
        Case "Phillips 66 - OS 32"
            TempArray = Array(60, 100, 320, 550)
            FluidDensityArray = Array(53.8597, 46.3792, 44.1351, 41.8909)
        Case "Phillips 66 - OS 46"
            TempArray = Array(60, 100, 320, 550)
            FluidDensityArray = Array(53.93455, 46.45403, 44.20987, 41.96571)
        Case "Phillips 66 - CS 32"
            TempArray = Array(60, 100, 320, 550)
            FluidDensityArray = Array(53.71013, 46.22961, 43.98545, 41.7413)
        Case "Duratherm HTO"
            TempArray = Array(15, 55, 105, 155, 205, 255, 305, 355, 405, 455, 505, 555, 600)
            FluidDensityArray = Array(52.53, 51.63, 50.51, 49.39, 48.27, 47.15, 46.03, 44.91, 43.79, 42.67, 41.55, 40.43, 39.44)
        Case "Duratherm HF"
            TempArray = Array(40, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 640)
            FluidDensityArray = Array(54.39, 54.26, 53.63, 52.99, 52.36, 51.73, 51.09, 50.46, 49.82, 49.19, 48.55, 47.92, 47.29, 46.78)
        Case "Seriola 1510"
            TempArray = Array(32, 50, 59, 68, 86, 104, 122, 140, 158, 176, 194, 212, 230, 248, 266, 284, 302, 320, 338, 356, 374, 392, 410, 428, 446, 464, 482, 500, 518, 536, 554, 572, 590)
            FluidDensityArray = Array(52.98, 52.56, 52.38, 52.2, 51.78, 51.42, 51, 50.64, 50.22, 49.86, 49.44, 49.08, 48.66, 48.3, 47.88, 47.52, 47.1, 46.74, 46.32, 45.96, 45.54, 45.18, 44.76, 44.4, 43.98, 43.62, 43.2, 42.84, 42.42, 42.06, 41.64, 41.28, 40.86)
        Case "Therminol 68"
            TempArray = Array(-14, 0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 660, 680)
            FluidDensityArray = Array(66.2, 65.8, 65.3, 64.8, 64.3, 63.8, 63.3, 62.8, 62.3, 61.8, 61.3, 60.8, 60.3, 59.8, 59.3, 58.8, 58.3, 57.8, 57.3, 56.8, 56.3, 55.8, 55.3, 54.8, 54.3, 53.8, 53.3, 52.8, 52.3, 51.8, 51.3, 50.8, 50.3, 49.8, 49.3, 48.8)
        Case "Seriola K 3000"
            TempArray = Array(32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572)
            FluidDensityArray = Array(52.795, 51.934, 51.079, 50.223, 49.368, 48.519, 47.67, 46.821, 45.972, 45.129, 44.286, 43.444, 42.607, 41.764, 40.934, 40.098)
    End Select



'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
    If FluidDensityArray(i + 1) = -1 Then
        FluidDensity = "Beyond Fluid Limits"
        Exit Function
    End If
    FluidDensity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (FluidDensityArray(i + 1) - FluidDensityArray(i)) + FluidDensityArray(i)
    i = i + 1
Loop
        
End Function

Function FluidSpecificHeat(Fluid As String, Temperature As Single)

Dim TempArray, FluidSpecificHeatArray As Variant

TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)


'all of the fluids go here
    Select Case Fluid
        Case "Shell Thermia C"
            FluidSpecificHeatArray = Array(0, 0.437633262, 0.461620469, 0.485607676, 0.509594883, 0.53358209, 0.557569296, 0.581556503, 0.60554371, 0.629530917, 0.653518124, 0.67750533, 0.701492537)
        Case "Shell Thermia B"
            FluidSpecificHeatArray = Array(0, 0.437633262, 0.461620469, 0.485607676, 0.509594883, 0.53358209, 0.557569296, 0.581556503, 0.60554371, 0.629530917, 0.653518124, 0.67750533, 0.701492537)
        Case "Calflo AF"
            FluidSpecificHeatArray = Array(0, 0.45, 0.47, 0.49, 0.51, 0.53, 0.56, 0.58, 0.6, 0.62, 0.64, 0.66, 0.69)
        Case "Calflo FG"
            FluidSpecificHeatArray = Array(0.42, 0.44, 0.46, 0.49, 0.51, 0.53, 0.55, 0.58, 0.6, 0.62, 0.64, 0.66, 0.69, 0.71)
        Case "Calflo HTF"
            FluidSpecificHeatArray = Array(0.4249, 0.4449, 0.4649, 0.4849, 0.5049, 0.5249, 0.5449, 0.5649, 0.5849, 0.6049, 0.6249, 0.6449, 0.6649)
        Case "Calflo LT"
            FluidSpecificHeatArray = Array(0.47, 0.49, 0.51, 0.54, 0.56, 0.58, 0.6, 0.63, 0.65, 0.67, 0.69, 0.72)
        Case "Dowtherm A"
            FluidSpecificHeatArray = Array(0, 0.362, 0.388, 0.407, 0.426, 0.444, 0.463, 0.481, 0.5, 0.518, 0.537, 0.558, 0.579, 0.596, 0.611, 0.633, 0.675)
        Case "Dowtherm G"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 725)
            FluidSpecificHeatArray = Array(0.361, 0.384, 0.407, 0.431, 0.454, 0.477, 0.5, 0.524, 0.547, 0.57, 0.593, 0.616, 0.64, 0.663, 0.674)
        Case "Dowtherm Q"
            FluidSpecificHeatArray = Array(0, 0.387, 0.409, 0.429, 0.45, 0.471, 0.491, 0.511, 0.531, 0.551, 0.57, 0.589, 0.609, 0.629)
        Case "Dowtherm RP"
            FluidSpecificHeatArray = Array(0, 0.38, 0.4, 0.42, 0.4395, 0.459, 0.479, 0.499, 0.5185, 0.538, 0.558, 0.578, 0.5975, 0.617, 0.6365)
        Case "Dowtherm HT"
            FluidSpecificHeatArray = Array(0, 0.349, 0.375, 0.4, 0.426, 0.451, 0.477, 0.503, 0.528, 0.554, 0.579, 0.605, 0.63, 0.656, 0.681)
        Case "Dowtherm J"
            FluidSpecificHeatArray = Array(0.4, 0.43, 0.45, 0.472, 0.494, 0.5185, 0.543, 0.569, 0.595, 0.621, 0.647, 0.6835, 0.72, 0.774)
        Case "Dowtherm MX"
            FluidSpecificHeatArray = Array(0.36, 0.377, 0.399, 0.42, 0.442, 0.464, 0.4852, 0.506, 0.5277, 0.5493, 0.571, 0.5927, 0.6143, 0.636, 0.6577)
        Case "Dowtherm T"
            TempArray = Array(20, 100, 180, 260, 340, 420, 500, 580, 600)
            FluidSpecificHeatArray = Array(0.45, 0.482, 0.513, 0.545, 0.577, 0.608, 0.64, 0.672, 0.68)
        Case "Mobiltherm 603"
            FluidSpecificHeatArray = Array(0.401181818, 0.441865, 0.46628, 0.490165, 0.515111, 0.538836, 0.562721, 0.587879, 0.612879, 0.636763, 0.6660648, 0.684532, 0.708045, -1, -1, -1, -1)
        Case "Therminol 55"
            FluidSpecificHeatArray = Array(0.423, 0.447, 0.471, 0.4945, 0.518, 0.5415, 0.565, 0.5885, 0.612, 0.635, 0.658, 0.6815, 0.705, 0.728)
        Case "Therminol 59"
            FluidSpecificHeatArray = Array(0.373, 0.394, 0.416, 0.437, 0.459, 0.4805, 0.503, 0.5245, 0.547, 0.5695, 0.593, 0.616, 0.64, 0.663)
        Case "Therminol SP"
            FluidSpecificHeatArray = Array(0.423, 0.447, 0.471, 0.4945, 0.518, 0.5415, 0.565, 0.5885, 0.612, 0.635, 0.658, 0.6815, 0.704999, 0.728449, 0.751899, 0.775349, 0.798799)
        Case "Therminol 66"
            FluidSpecificHeatArray = Array(0.3425, 0.3655, 0.388, 0.4105, 0.434, 0.457, 0.48, 0.504, 0.528, 0.553, 0.578, 0.603, 0.628, 0.6545, 0.682)
        Case "Therminol D-12"
            FluidSpecificHeatArray = Array(0.465, 0.491, 0.517, 0.5435, 0.57, 0.5975, 0.626, 0.6545, 0.684, 0.7155, 0.741, 0.77)
        Case "Therminol VP-1"
            FluidSpecificHeatArray = Array(0.346, 0.3627, 0.382, 0.401, 0.42, 0.4385, 0.457, 0.4745, 0.492, 0.51, 0.528, 0.5455, 0.563, 0.5825, 0.602, 0.6275, 0.662)
        Case "Therminol 62"
            FluidSpecificHeatArray = Array(0.436, 0.458000004, 0.476, 0.492500007, 0.509, 0.524000049, 0.538, 0.552000046, 0.565, 0.577499986, 0.589, 0.600499988, 0.612, 0.617, 0.622, 0.624065217)
        Case "Therminol 72"
             TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 715)
            FluidSpecificHeatArray = Array(0.346, 0.3645, 0.382, 0.4005, 0.418, 0.4365, 0.454, 0.4725, 0.49, 0.5085, 0.526, 0.5445, 0.562, 0.5805, 0.598, 0.604)
        Case "Therminol 75"
            TempArray = Array(0, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700, 725)
            FluidSpecificHeatArray = Array(0.338, 0.358, 0.368, 0.378, 0.388, 0.398, 0.408, 0.418, 0.428, 0.438, 0.448, 0.457, 0.466, 0.475, 0.483, 0.493, 0.5005, 0.508, 0.516, 0.525, 0.531, 0.538, 0.546, 0.552, 0.5585, 0.566, 0.572, 0.578, 0.584)
        Case "Therminol XP"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidSpecificHeatArray = Array(0.389, 0.4225, 0.454, 0.485, 0.515, 0.544, 0.572, 0.599, 0.625, 0.65, 0.673, 0.6965, 0.718)
        Case "Paratherm NF"
            FluidSpecificHeatArray = Array(0, 0.44, 0.465, 0.489, 0.511, 0.536, 0.56, 0.585, 0.609, 0.633, 0.659, 0.681, 0.705, 0.73)
        Case "Paratherm MG"
            FluidSpecificHeatArray = Array(0.5148, 0.5298, 0.5448, 0.5598, 0.5748, 0.5898, 0.6048, 0.6198, 0.6348, 0.6498, 0.6648, -1, -1, -1, -1, -1, -1)
        Case "Paratherm MR"
            FluidSpecificHeatArray = Array(0.4977, 0.5157, 0.5337, 0.5517, 0.5697, 0.5877, 0.6057, 0.6237, 0.6417, 0.6597, 0.6777, 0.6957, 0.7137)
        Case "Paratherm HE"
            FluidSpecificHeatArray = Array(0.4172, 0.4414, 0.4656, 0.4898, 0.5139, 0.5381, 0.5623, 0.5865, 0.6107, 0.6349, 0.6591, 0.6833, 0.7075, 0.6831, 0.688, 0.6928, 0.6977, 0.71461)
        Case "Paratherm HR"
            TempArray = Array(0, 25, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700)
            FluidSpecificHeatArray = Array(0.44, 0.45, 0.46, 0.46, 0.47, 0.48, 0.49, 0.49, 0.5, 0.51, 0.52, 0.52, 0.53, 0.54, 0.55, 0.55, 0.56, 0.57, 0.58, 0.58, 0.59, 0.6, 0.61, 0.61, 0.62, 0.63, 0.64, 0.64, 0.65)
        Case "Petrotherm"
            FluidSpecificHeatArray = Array(0, 0.43, 0.453, 0.479, 0.505, 0.55, 0.5565, 0.582, 0.60825, 0.64, 0.66, 0.68, 0.712)
        Case "PetroCanada FG"
            FluidSpecificHeatArray = Array(0, 0.446, 0.46, 0.485000014, 0.51, 0.529999971, 0.55, 0.575000048, 0.6, 0.624500036, 0.649, 0.669499993, 0.69)
        Case "Syltherm 800"
            FluidSpecificHeatArray = Array(0, 0.3805, 0.392, 0.403, 0.414, 0.4255, 0.437, 0.4485, 0.46, 0.471, 0.482, 0.4935, 0.505, 0.5165, 0.528, 0.539, 0.55)
        Case "TEG 0.6"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidSpecificHeatArray = Array(0.70554174, 0.73148, 0.75865284, 0.78706025, 0.81670222, 0.84757876, 0.87968988, 0.91303555, 0.9476158)
        Case "TEG 0.98"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidSpecificHeatArray = Array(0.472324704, 0.540834204, 0.609305623, 0.677738963, 0.746134222, 0.814491402, 0.882810501, 0.951320001, 0.6812)
        Case "EGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.845, 0.848, 0.852, 0.856, 0.86, 0.864, 0.868, 0.871, 0.875, 0.879, 0.883, 0.887, 0.891, 0.895, 0.898, 0.902, 0.906, 0.91, 0.914, 0.918, 0.921, 0.925, 0.929, 0.933, 0.937, 0.941, 0.944, 0.948, 0.952, 0.956, 0.96, 0.964, 0.967, 0.971, 0.975)
        Case "EGlycol 0.4"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.792, 0.796, 0.801, 0.805, 0.81, 0.814, 0.819, 0.824, 0.828, 0.833, 0.837, 0.842, 0.846, 0.851, 0.855, 0.86, 0.865, 0.869, 0.874, 0.878, 0.883, 0.887, 0.892, 0.896, 0.901, 0.905, 0.91, 0.915, 0.919, 0.924, 0.928, 0.933, 0.937, 0.942, 0.946, 0.951)
        Case "EGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.73, 0.735, 0.74, 0.745, 0.751, 0.756, 0.761, 0.766, 0.772, 0.777, 0.782, 0.787, 0.793, 0.798, 0.803, 0.808, 0.814, 0.819, 0.824, 0.829, 0.835, 0.84, 0.845, 0.85, 0.856, 0.896, 0.901, 0.905, 0.91, 0.915, 0.919, 0.924, 0.928, 0.933, 0.937, 0.942, 0.946, 0.951)
        Case "Dynalene EG-XT 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 120, 140, 160, 180, 200, 220)
            FluidSpecificHeatArray = Array(0.73, 0.735, 0.74, 0.745, 0.751, 0.756, 0.761, 0.766, 0.772, 0.777, 0.782, 0.781, 0.793, 0.803, 0.814, 0.824, 0.835, 0.845, 0.856)
        Case "EGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.668, 0.674, 0.68, 0.686, 0.692, 0.698, 0.704, 0.71, 0.716, 0.722, 0.728, 0.734, 0.74, 0.746, 0.751, 0.757, 0.763, 0.769, 0.775, 0.781, 0.787, 0.793, 0.799, 0.805, 0.811, 0.817, 0.823, 0.829, 0.834, 0.84, 0.846, 0.852, 0.858, 0.864, 0.87, 0.876, 0.882, 0.888, 0.894)
        Case "EGlycol 0.7"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.61, 0.617, 0.624, 0.63, 0.637, 6.43, 0.65, 0.657, 0.663, 0.67, 0.676, 0.683, 0.69, 0.696, 0.703, 0.709, 0.716, 0.723, 0.729, 0.736, 0.742, 0.749, 0.756, 0.762, 0.769, 0.775, 0.782, 0.789, 0.795, 0.802, 0.808, 0.815, 0.822, 0.828, 0.835, 0.841, 0.848, 0.855, 0.861)
        Case "EGlycol 0.8"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.572, 0.579, 0.586, 0.593, 0.601, 0.608, 0.615, 0.622, 0.63, 0.637, 0.644, 0.652, 0.659, 0.667, 0.673, 0.681, 0.688, 0.695, 0.702, 0.71, 0.717, 0.724, 0.731, 0.739, 0.746, 0.753, 0.76, 0.768, 0.775, 0.782, 0.789, 0.797, 0.804, 0.811, 0.819, 0.826)
        Case "EGlycol 0.9"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.511, 0.519, 0.527, 0.534, 0.542, 0.55, 0.558, 0.566, 0.574, 0.582, 0.59, 0.598, 0.606, 0.614, 0.621, 0.629, 0.637, 0.645, 0.653, 0.661, 0.669, 0.677, 0.685, 0.693, 0.701, 0.709, 0.716, 0.724, 0.732, 0.74, 0.748, 0.756, 0.764, 0.772, 0.78, 0.788)
        Case "PGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidSpecificHeatArray = Array(0.898, 0.902, 0.906, 0.909, 0.913, 0.917, 0.92, 0.924, 0.928, 0.931, 0.935, 0.939, 0.942, 0.946, 0.95, 0.953, 0.957, 0.961, 0.964, 0.968, 0.971, 0.962, 0.966, 0.97, 0.973, 0.977, 0.981, 0.985, 0.989, 0.992, 0.996, 1, 1.002)
        Case "PGlycol 0.4"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidSpecificHeatArray = Array(0.855, 0.859, 0.864, 0.868, 0.872, 0.877, 0.881, 0.886, 0.898, 0.894, 0.899, 0.903, 0.908, 0.912, 0.916, 0.921, 0.925, 0.929, 0.934, 0.938, 0.943, 0.947, 0.951, 0.94, 0.945, 0.95, 0.955, 0.96, 0.965, 0.969, 0.974, 0.979, 0.984, 0.987)
        Case "PGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.799, 0.804, 0.809, 0.814, 0.82, 0.825, 0.83, 0.835, 0.84, 0.845, 0.85, 0.855, 0.861, 0.866, 0.871, 0.876, 0.881, 0.886, 0.891, 0.896, 0.902, 0.907, 0.912, 0.917, 0.922, 0.908, 0.914, 0.92, 0.926, 0.932, 0.938, 0.944, 0.95, 0.956, 0.962, 0.965)
        Case "PGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidSpecificHeatArray = Array(0.684, 0.691, 0.698, 0.705, 0.712, 0.719, 0.727, 0.734, 0.741, 0.748, 0.755, 0.762, 0.769, 0.776, 0.783, 0.791, 0.798, 0.805, 0.812, 0.819, 0.826, 0.833, 0.84, 0.847, 0.855, 0.862, 0.869, 0.876, 0.883, 0.89, 0.897, 0.904, 0.912, 0.919, 0.926, 0.933, 0.936)
        Case "PGlycol 0.9"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidSpecificHeatArray = Array(0.55, 0.558, 0.566, 0.574, 0.583, 0.591, 0.599, 0.607, 0.615, 0.623, 0.631, 0.639, 0.647, 0.656, 0.664, 0.672, 0.68, 0.688, 0.696, 0.704, 0.712, 0.72, 0.729, 0.737, 0.745, 0.753, 0.761, 0.769)
        Case "Salt"
            TempArray = Array(275, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100)
            FluidSpecificHeatArray = Array(0.356, 0.356, 0.356, 0.356, 0.356, 0.356, 0.356, 0.358, 0.359, 0.36, 0.361, 0.362, 0.363, 0.364, 0.366, 0.367, 0.368, 0.369, 0.37)
        Case "Xceltherm HT"
            FluidSpecificHeatArray = Array(0, 0.372, 0.397, 0.422, 0.447, 0.472, 0.497, 0.522, 0.547, 0.572, 0.597, 0.622, 0.647, 0.672, 0.695)
        Case "Xceltherm 600"
            FluidSpecificHeatArray = Array(0, 0.467, 0.49, 0.515, 0.537, 0.56, 0.582, 0.605, 0.627, 0.65, 0.672, 0.695, 0.717, 0.74)
        Case "Syltherm XLT"
            FluidSpecificHeatArray = Array(0, 0.388, 0.405, 0.422, 0.439, 0.456, 0.473, 0.49, 0.507, 0.524, 0.541, 0.558)
        Case "Texatherm 46"
            FluidSpecificHeatArray = Array(0, 0.4263, 0.4513, 0.4763, 0.5013, 0.5263, 0.5513, 0.5763, 0.6013, 0.6263, 0.6513, 0.6763, 0.7013, 0.7263)
        Case "Multitherm PG-1"
            FluidSpecificHeatArray = Array(0.416, 0.44, 0.464, 0.488, 0.512, 0.536, 0.56, 0.585, 0.609, 0.633, 0.659, 0.681, 0.705, 0.72)
        Case "Mobiltherm 43"
            FluidSpecificHeatArray = Array(0.415, 0.438, 0.463, 0.486, 0.511, 0.535, 0.559, 0.583, 0.607, 0.631, 0.655, 0.679, 0.703)
        Case "Multitherm OG-1"
            TempArray = Array(0, 20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidSpecificHeatArray = Array(0.418, 0.42, 0.456, 0.476, 0.49, 0.514, 0.536, 0.562, 0.587, 0.605, 0.631, 0.659, 0.684, 0.7)
        Case "Marlotherm SH"
            FluidSpecificHeatArray = Array(0.35, 0.361, 0.385, 0.4134, 0.436, 0.4594, 0.4828, 0.51, 0.535, 0.5522, 0.583, 0.61, 0.63, 0.66, 0.68)
        Case "Marlotherm LH"
            TempArray = Array(-4, 32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572, 608, 644, 680)
            FluidSpecificHeatArray = Array(0.353, 0.37, 0.387, 0.401, 0.418, 0.435, 0.449, 0.466, 0.482, 0.497, 0.514, 0.53, 0.547, 0.561, 0.578, 0.595, 0.609, 0.626, 0.642, 0.657)
        Case "Sun 21"
            TempArray = Array(150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidSpecificHeatArray = Array(0.515, 0.539, 0.565, 0.593, 0.611, 0.649, 0.676, 0.703, 0.732, 0.761, 0.7858)
        Case "Duratherm FG"
            FluidSpecificHeatArray = Array(0.87, 0.857610501, 0.839610501, 0.821610501, 0.803610501, 0.785610501, 0.767610501, 0.749610501, 0.731610501, 0.713610501, 0.695610501, 0.677610501, 0.659610501, 0.641610501)
        Case "Noco 21"
            TempArray = Array(50, 150, 300, 450, 600)
            FluidSpecificHeatArray = Array(0.45, 0.52, 0.59, 0.68, 0.76)
        Case "Chemtherm 550"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidSpecificHeatArray = Array(0.422, 0.445, 0.461, 0.52, 0.541, 0.568, 0.59, 0.614, 0.622, 0.639, 0.66, 0.684, 0.71)
        Case "Thermalane 550"
            TempArray = Array(0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600)
            FluidSpecificHeatArray = Array(0.468, 0.476, 0.485, 0.494, 0.502, 0.511, 0.52, 0.528, 0.537, 0.546, 0.554, 0.563, 0.571, 0.58, 0.589, 0.597, 0.606, 0.615, 0.623, 0.632, 0.641, 0.649, 0.658, 0.667, 0.675, 0.684, 0.692, 0.701, 0.71, 0.718, 0.727)
        Case "Shell S2"
            TempArray = Array(32, 68, 104, 212, 302, 392, 482, 572, 644)
            FluidSpecificHeatArray = Array(0.432072414, 0.449508172, 0.466705084, 0.519012358, 0.56248233, 0.606191148, 0.64966112, 0.693131092, 0.728002608)
        Case "Thermoil 100"
            TempArray = Array(0, 100, 212, 392, 500, 600)
            FluidSpecificHeatArray = Array(0.4, 0.461, 0.514, 0.602, 0.652, 0.7)
        Case "Uconhtf 500"
            TempArray = Array(0, 100, 200, 300, 400, 500)
            FluidSpecificHeatArray = Array(0.435, 0.48, 0.518, 0.545, 0.56, 0.572)
        Case "Multitherm IG-1"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)
            FluidSpecificHeatArray = Array(0.419, 0.442, 0.466, 0.49, 0.514, 0.538, 0.563, 0.587, 0.611, 0.635, 0.659, 0.684, 0.705)
        Case "Multitherm IG-4"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidSpecificHeatArray = Array(0.438, 0.456, 0.486, 0.516, 0.547, 0.577, 0.607, 0.637, 0.667, 0.697, 0.728, 0.758, 0.788)
        Case "Hydroclear O/S 32"
            FluidSpecificHeatArray = Array(0.4357, 0.4557, 0.48, 0.4957, 0.5157, 0.5357, 0.5557, 0.5757, 0.61, 0.6157, 0.65, 0.6557, 0.7, 0, 0, 0, 0)
        Case "Hydroclear O/S 46"
            FluidSpecificHeatArray = Array(0.4282, 0.4482, 0.47, 0.4882, 0.5082, 0.5282, 0.5482, 0.5682, 0.5882, 0.6082, 0.6282, 0.6482, 0.6682, 0, 0, 0, 0)
        Case "Hydroclear C/S 32"
            FluidSpecificHeatArray = Array(0.4357, 0.4557, 0.48, 0.4957, 0.5157, 0.5357, 0.5557, 0.5757, 0.61, 0.6157, 0.65, 0.6557, 0.7, 0, 0, 0, 0)
        Case "Hydroclear C/S 46"
            FluidSpecificHeatArray = Array(0.4282, 0.4497, 0.47, 0.4927, 0.5142, 0.5357, 0.5572, 0.5787, 0.61, 0.6217, 0.64, 0.6647, 0.69, 0, 0, 0, 0)
        Case "Xceltherm SST"
            FluidSpecificHeatArray = Array(0.44, 0.46, 0.47, 0.49, 0.5, 0.52, 0.53, 0.55, 0.56, 0.58, 0.59, 0.61, 0.62, 0.64, 0.65, 0, 0)
        Case "Paratherm HT"
            TempArray = Array(50, 100, 200, 300, 350, 400, 450, 500, 550)
            FluidSpecificHeatArray = Array(0.405, 0.423, 0.459, 0.496, 0.514, 0.532, 0.55, 0.569, 0.587)
        Case "Chevron 22"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidSpecificHeatArray = Array(0.45, 0.484, 0.51375, 0.5435, 0.57325, 0.603, 0.62825, 0.6535, 0.67875, 0.704, 0.7255, 0.747, 0.77)
        Case "Mobiltherm 605"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidSpecificHeatArray = Array(0.4439, 0.4689, 0.4939, 0.5189, 0.5439, 0.5689, 0.5939, 0.6189, 0.6439, 0.6689, 0.6939, 0.7189)
        Case "Duratherm 600"
            TempArray = Array(15, 55, 95, 145, 195, 245, 295, 345, 395, 445, 495, 545, 600)
            FluidSpecificHeatArray = Array(0.425, 0.442, 0.459, 0.48, 0.501, 0.523, 0.544, 0.565, 0.586, 0.608, 0.629, 0.65, 0.673)
        Case "Chevron 46"
            TempArray = Array(32, 104, 122, 212, 302, 392, 482, 572, 662)
            FluidSpecificHeatArray = Array(0.432, 0.467, 0.476, 0.52, 0.563, 0.607, 0.65, 0.694, 0.737)
        Case "Phillips 66 - OS 32"
            TempArray = Array(60, 100, 320, 550)
            FluidSpecificHeatArray = Array(0.45, 0.621, 0.665, 0.7)
        Case "Phillips 66 - OS 46"
            TempArray = Array(60, 100, 320, 550)
            FluidSpecificHeatArray = Array(0.45, 0.619, 0.663, 0.7)
        Case "Phillips 66 - CS 32"
            TempArray = Array(60, 100, 320, 550)
            FluidSpecificHeatArray = Array(0.45, 0.621, 0.665, 0.7)
        Case "Duratherm HTO"
            TempArray = Array(15, 55, 105, 155, 205, 255, 305, 355, 405, 455, 505, 555, 600)
            FluidSpecificHeatArray = Array(0.416, 0.433, 0.454, 0.474, 0.495, 0.516, 0.537, 0.562, 0.578, 0.599, 0.62, 0.641, 0.659)
        Case "Multitherm IG-4"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidSpecificHeatArray = Array(0.438, 0.456, 0.516, 0.547, 0.577, , 0.607, 0.637, 0.667, 0.697, 0.728, 0.758, 0.788)
        Case "Duratherm HF"
            TempArray = Array(40, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 640)
            FluidSpecificHeatArray = Array(0.434, 0.436, 0.448, 0.46, 0.471, 0.483, 0.495, 0.507, 0.518, 0.53, 0.542, 0.554, 0.565, 0.575)
        Case "Seriola 1510"
            TempArray = Array(32, 50, 59, 68, 86, 104, 122, 140, 158, 176, 194, 212, 230, 248, 266, 284, 302, 320, 338, 356, 374, 392, 410, 428, 446, 464, 482, 500, 518, 536, 554, 572, 590)
            FluidSpecificHeatArray = Array(0.4296, 0.438, 0.44232, 0.4464, 0.4548, 0.4632, 0.4716, 0.48, 0.4884, 0.4968, 0.5052, 0.5136, 0.522, 0.5304, 0.5388, 0.5472, 0.5556, 0.564, 0.5724, 0.5808, 0.5892, 0.5976, 0.606, 0.6144, 0.6228, 0.6312, 0.6396, 0.648, 0.6564, 0.6648, 0.6732, 0.6816, 0.69)
        Case "Therminol 68"
            TempArray = Array(-14, 0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 660, 680)
            FluidSpecificHeatArray = Array(0.353, 0.359, 0.368, 0.376, 0.385, 0.393, 0.402, 0.41, 0.419, 0.427, 0.436, 0.445, 0.453, 0.462, 0.47, 0.479, 0.487, 0.496, 0.505, 0.513, 0.522, 0.53, 0.539, 0.547, 0.556, 0.564, 0.573, 0.582, 0.59, 0.599, 0.607, 0.616, 0.624, 0.633, 0.642, 0.65)
        Case "Seriola K 3000"
            TempArray = Array(32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572)
            FluidSpecificHeatArray = Array(0.441, 0.459, 0.476, 0.494, 0.512, 0.53, 0.547, 0.565, 0.583, 0.601, 0.618, 0.636, 0.654, 0.671, 0.689, 0.707)
    End Select

'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
    If FluidSpecificHeatArray(i + 1) = -1 Then
        FluidSpecificHeat = "Beyond Fluid Limits"
        Exit Function
    End If
        
    FluidSpecificHeat = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (FluidSpecificHeatArray(i + 1) - FluidSpecificHeatArray(i)) + FluidSpecificHeatArray(i)
    i = i + 1
Loop


End Function
Function FluidViscosity(Fluid As String, Temperature As Single)

Dim TempArray, FluidViscosityArray As Variant

TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)


'all of the fluids go here
    Select Case Fluid
        Case "Shell Thermia C"
            FluidViscosityArray = Array(0, 331.97, 54.19, 16.29, 7.01, 3.76, 2.33, 1.59, 1.16, 0.89, 0.7, 0.57, 0.47)
        Case "Shell Thermia B"
            FluidViscosityArray = Array(0, 331.97, 54.19, 16.29, 7.01, 3.76, 2.33, 1.59, 1.16, 0.89, 0.7, 0.57, 0.47)
        Case "Calflo AF"
            FluidViscosityArray = Array(0, 86.4, 42, 8.28, 7.236, 4.68, 3.072, 2.232, 1.464, 0.708, 0.6264, 0.5376, 0.4536)
        Case "Calflo FG"
            FluidViscosityArray = Array(100, 85.2, 42, 8.16, 6.432, 4.68, 3.072, 2.232, 1.464, 0.708, 0.6156, 0.5376, 0.4536, 0.39)
        Case "Calflo HTF"
            FluidViscosityArray = Array(0, 100, 35.9, 13.83, 7.274, 4.419, 2.942, 2.085, 1.547, 1.19, 0.94, 0.76, 0.7)
        Case "Calflo LT"
            FluidViscosityArray = Array(50, 24.48, 7.236, 4.68, 3.072, 2.232, 1.08, 0.708, 0.6156, 0.5712, 0.5346, 0.5088)
        Case "Dowtherm A"
            FluidViscosityArray = Array(0, 5, 2.6, 1.45, 1.01, 0.75, 0.565, 0.43, 0.365, 0.32, 0.26, 0.225, 0.19, 0.17, 0.15, 0.125, 0.105)
        Case "Dowtherm G"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 725)
            FluidViscosityArray = Array(20.4, 6.51, 3.24, 1.97, 1.34, 0.97, 0.73, 0.57, 0.46, 0.37, 0.31, 0.26, 0.23, 0.2, 0.18)
        Case "Dowtherm Q"
            FluidViscosityArray = Array(0, 5.76, 2.77, 1.62, 1.06, 0.76, 0.57, 0.45, 0.36, 0.3, 0.26, 0.23, 0.2, 0.18)
        Case "Dowtherm RP"
            FluidViscosityArray = Array(0, 88.17, 16.46, 6.1, 3.15, 1.94, 1.34, 0.99, 0.77, 0.62, 0.51, 0.43, 0.37, 0.32, 0.28)
        Case "Dowtherm HT"
            FluidViscosityArray = Array(0, 256, 28.9, 8.73, 4.02, 2.3, 1.5, 1.07, 0.81, 0.64, 0.53, 0.44, 0.38, 0.33, 0.29)
        Case "Dowtherm J"
            FluidViscosityArray = Array(3, 1.055, 0.72, 0.535, 0.42, 0.35, 0.3, 0.26, 0.23, 0.205, 0.19, 0.175, 0.16, 0.15)
        Case "Dowtherm MX"
            FluidViscosityArray = Array(65, 32.3, 12.425, 5.083, 2.575, 1.5167, 0.992, 0.7, 0.533, 0.4133, 0.33, 0.27, 0.22333, 0.19, 0.1567)
        Case "Dowtherm T"
            TempArray = Array(20, 100, 180, 260, 340, 420, 500, 580, 600)
            FluidViscosityArray = Array(184.8, 13.9, 3.87, 1.74, 1#, 0.65, 0.47, 0.36, 0.33)
        Case "Mobiltherm 603"
            FluidViscosityArray = Array(157.3531438, 94.7, 20.2, 7.738889, 3.916667, 2.56, 1.56, 1.184, 0.860667, 0.677333, 0.526, 0.431556, 0.351111, -1, -1, -1, -1)
        Case "Therminol 55"
            FluidViscosityArray = Array(611.801, 79.162, 17.858, 6.821, 3.311, 1.978, 1.3104, 0.9487, 0.719, 0.5643, 0.451, 0.366, 0.2985, 0.24)
        Case "Therminol 59"
            FluidViscosityArray = Array(45.06, 10.252, 4.097, 2.278, 1.447, 1.013, 0.748, 0.579, 0.459, 0.378, 0.315, 0.268, 0.231, 0.2, -1, -1, -1, -1)
        Case "Therminol 66"
            FluidViscosityArray = Array(3000, 441.489, 33.566, 9.652, 4.175, 2.396, 1.546, 1.0975, 0.827, 0.6494, 0.529, 0.444, 0.3795, 0.331, 0.294)
        Case "Therminol SP"
            FluidViscosityArray = Array(611.801, 79.162, 17.858, 6.821, 3.311, 1.978, 1.3104, 0.949, 0.7193, 0.5643, 0.4506, 0.366, 0.2985, -1, -1, -1, -1)
        Case "Therminol D-12"
            FluidViscosityArray = Array(3.0177, 1.5357, 0.9466, 0.6635, 0.4878, 0.3743, 0.2931, 0.234, 0.1881, 0.1534, 0.14, 0.13)
        Case "Therminol VP-1"
            FluidViscosityArray = Array(10.66, 5.84, 2.7283, 1.6142, 1.0707, 0.7772, 0.5911, 0.4692, 0.3828, 0.3197, 0.272, 0.2352, 0.2059, 0.1823, 0.1629, 0.1472, 0.1335)
        Case "Therminol 62"
            FluidViscosityArray = Array(826.758, 49.8742, 10.9132, 4.7208, 2.5505, 1.5998, 1.0748, 0.7648, 0.5622, 0.4276, 0.3315, 0.263500005, 0.2125, 0.196, 0.181, 0.163546584)
        Case "Therminol 72"
             TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 715)
            FluidViscosityArray = Array(3505, 24.448, 6.531, 3.113, 1.798, 1.186, 0.835, 0.622, 0.4795, 0.3803, 0.3059, 0.252, 0.211, 0.176, 0.1488, 0.145)
        Case "Therminol 75"
             TempArray = Array(0, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700, 725)
            FluidViscosityArray = Array(12.418, 10.113, 8.961, 7.809, 6.656, 5.504, 4.352, 3.398, 2.724, 2.212, 1.817, 1.509, 1.275, 1.089, 0.934, 0.81, 0.71, 0.626, 0.555, 0.496, 0.445, 0.402, 0.364, 0.332, 0.304, 0.279, 0.258, 0.2385, 0.2216)
        Case "Therminol XP"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidViscosityArray = Array(1405.5, 120.91, 22.653, 8.164, 3.8775, 1.873, 1.501, 1.075, 0.806, 0.628, 0.5002, 0.409, 0.337)
        Case "Paratherm NF"
            FluidViscosityArray = Array(0, 88, 16, 6.4, 3.45, 2.25, 1.6, 1.2, 0.92, 0.72, 0.57, 0.46, 0.37, 0.3)
        Case "Paratherm MG"
            FluidViscosityArray = Array(25.3, 8.03, 3.67, 2.08, 1.34, 0.95, 0.71, 0.56, 0.46, 0.39, 0.33, -1, -1, -1, -1, -1, -1)
        Case "Paratherm MR"
            FluidViscosityArray = Array(23.8, 7.54, 3.46, 1.97, 1.28, 0.91, 0.69, 0.55, 0.45, 0.38, 0.33, 0.29, 0.27)
        Case "Paratherm HE"
            FluidViscosityArray = Array(2161, 182, 37.5, 12.9, 6.01, 3.39, 2.17, 1.52, 1.13, 0.906, 0.694, 0.581, 0.472, 0.457, 0.442, 0.427, 0.411, 0.39567)
         Case "Paratherm HR"
            TempArray = Array(0, 25, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700)
            FluidViscosityArray = Array(544, 134, 50, 21, 12, 7.1, 4.7, 3.3, 2.5, 1.9, 1.5, 1.3, 1.1, 0.95, 0.82, 0.72, 0.64, 0.57, 0.51, 0.46, 0.42, 0.38, 0.35, 0.32, 0.3, 0.28, 0.26, 0.24, 0.22)
        Case "Petrotherm"
            FluidViscosityArray = Array(0, 95, 46.6, 9.9, 7.236, 5.8, 4.6, 3.2, 2, 1.19, 0.99, 0.84, 0.74)
         Case "PetroCanada FG"
            FluidViscosityArray = Array(0, 75, 50, 29.5, 9, 6.5, 4, 3, 2, 1.5, 1, 0.949999988, 0.9)
         Case "Syltherm 800"
            FluidViscosityArray = Array(0, 12.3, 7.3, 4.72, 3.25, 2.333, 1.72, 1.3, 1.01, 0.792, 0.63, 0.509, 0.42, 0.35, 0.29, 0.25, 0.21)
         Case "TEG 0.6"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidViscosityArray = Array(96.10363512, 15.81621785, 5.200118657, 2.385555908, 1.384046447, 0.717123052, 0.43136791, 0.119746, 0.051181)
         Case "TEG 0.98"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidViscosityArray = Array(1322.774503, 85.96598395, 20.65414652, 7.760105027, 3.863035193, 1.676628651, 0.878739269, 0.34, 0.5442)
        Case "EGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(6.83, 5.38, 4.33, 3.54, 2.95, 2.49, 2.13, 1.84, 1.6, 1.41, 1.25, 1.11, 1, 0.9, 0.82, 0.75, 0.68, 0.63, 0.58, 0.54, 0.5, 0.46, 0.43, 0.4, 0.38, 0.36, 0.34, 0.32, 0.3, 0.29, 0.27, 0.26, 0.25, 0.24, 0.23)
        Case "EGlycol 0.4"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(13.76, 10.13, 7.74, 6.09, 4.91, 4.04, 3.38, 2.87, 2.46, 2.13, 1.87, 1.64, 1.46, 1.3, 1.17, 1.05, 0.95, 0.87, 0.79, 0.73, 0.67, 0.61, 0.57, 0.53, 0.49, 0.45, 0.42, 0.4, 0.37, 0.35, 0.33, 0.31, 0.29, 0.28, 0.26, 0.25)
        Case "EGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(40.38, 27.27, 19.34, 14.26, 10.85, 8.48, 6.77, 5.5, 4.55, 3.81, 3.23, 2.76, 2.39, 2.08, 1.82, 1.61, 1.43, 1.28, 1.15, 1.04, 0.94, 0.85, 0.78, 0.71, 0.66, 0.6, 0.56, 0.52, 0.48, 0.45, 0.42, 0.39, 0.37, 0.34, 0.32, 0.3, 0.29, 0.27)
        Case "Dynalene EG-XT 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 120, 140, 160, 180, 200, 220)
            FluidViscosityArray = Array(40.38, 27.27, 19.34, 14.26, 10.85, 8.48, 6.77, 5.5, 4.55, 3.81, 3.23, 2.76, 2.39, 1.82, 1.43, 1.15, 0.94, 0.78, 0.66)
        Case "EGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(89.67, 60.46, 42.05, 30.08, 22.06, 16.56, 12.68, 9.9, 7.85, 6.33, 5.17, 4.28, 3.58, 3.03, 2.58, 2.23, 1.93, 1.69, 1.49, 1.32, 1.18, 1.06, 0.95, 0.86, 0.78, 0.72, 0.66, 0.61, 0.56, 0.52, 0.48, 0.45, 0.42, 0.39, 0.37, 0.35, 0.33, 0.31, 0.29)
        Case "EGlycol 0.7"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(128.79, 89.93, 63.5, 45.58, 33.31, 24.79, 18.77, 14.45, 11.31, 8.97, 7.22, 5.88, 4.85, 4.04, 3.4, 2.88, 2.47, 2.13, 1.86, 1.63, 1.43, 1.27, 1.14, 1.02, 0.92, 0.83, 0.76, 0.69, 0.63, 0.58, 0.54, 0.5, 0.46, 0.43, 0.4, 0.38, 0.35, 0.33, 0.31)
        Case "EGlycol 0.8"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(65.04, 46.89, 34.48, 25.84, 19.71, 15.29, 12.05, 9.62, 7.79, 6.38, 5.28, 4.41, 3.73, 3.17, 2.72, 2.35, 2.05, 1.8, 1.58, 1.4, 1.25, 1.12, 1.01, 0.91, 0.83, 0.75, 0.69, 0.63, 0.58, 0.53, 0.5, 0.46, 0.43, 0.4, 0.37, 0.35)
        Case "EGlycol 0.9"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(107.77, 71.87, 49.94, 35.91, 26.59, 20.18, 15.65, 12.37, 9.93, 8.1, 6.68, 5.58, 4.71, 4.01, 3.45, 2.98, 2.6, 2.28, 2.01, 1.79, 1.6, 1.43, 1.29, 1.16, 1.06, 0.96, 0.88, 0.81, 0.74, 0.69, 0.63, 0.59, 0.55, 0.51, 0.48, 0.45)
        Case "PGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidViscosityArray = Array(13.42, 9.89, 7.46, 5.75, 4.52, 3.62, 2.94, 2.43, 2.04, 1.73, 1.49, 1.3, 1.14, 1.01, 0.91, 0.82, 0.74, 0.68, 0.62, 0.58, 0.54, 0.5, 0.47, 0.44, 0.42, 0.4, 0.38, 0.36, 0.35, 0.33, 0.32, 0.31, 0.31)
        Case "PGlycol 0.4"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidViscosityArray = Array(40.92, 26.99, 18.5, 13.12, 9.6, 7.21, 5.56, 4.38, 3.52, 2.88, 2.4, 2.03, 1.73, 1.5, 1.31, 1.16, 1.04, 0.93, 0.85, 0.77, 0.71, 0.66, 0.61, 0.57, 0.53, 0.5, 0.48, 0.45, 0.43, 0.43, 0.39, 0.38, 0.37, 0.36)
        Case "PGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(130, 95.97, 61.32, 40.62, 27.83, 19.66, 14.28, 10.65, 8.13, 6.34, 5.04, 4.08, 3.35, 2.79, 2.36, 2.02, 1.75, 1.53, 0.135, 1.2, 1.08, 0.97, 0.88, 0.81, 0.74, 0.69, 0.64, 0.59, 0.56, 0.52, 0.49, 0.47, 0.44, 0.43, 0.42, 0.4)
        Case "PGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidViscosityArray = Array(497.57, 298.75, 182.96, 114.9, 74.19, 49.29, 33.68, 23.65, 17.05, 12.59, 9.51, 7.34, 5.77, 4.62, 3.76, 3.11, 2.61, 2.22, 1.91, 1.66, 1.45, 1.29, 1.15, 1.04, 0.94, 0.86, 0.79, 0.73, 0.68, 0.63, 0.59, 0.53, 0.5, 0.48, 0.47, 0.46, 0.45)
        Case "PGlycol 0.9"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidViscosityArray = Array(1819.72, 983.05, 558.32, 332.02, 205.91, 132.67, 88.51, 60.93, 43.16, 31.37, 23.35, 17.75, 13.76, 10.86, 8.71, 7.09, 5.85, 4.89, 4.13, 3.52, 3.04, 2.64, 2.31, 2.04, 1.82, 1.63, 1.47, 1.33)
        Case "Salt"
            TempArray = Array(275, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800, 850, 900, 950, 1000, 1050, 1100)
            FluidViscosityArray = Array(25.3961, 20.028, 15.8472, 3.8, 4.649, 3.2558, 2.0826, 1.834, 1.72, 1.5535, 1.4626, 1.427, 1.4095, 1.392, 1.3745, 1.357, 1.3227, 1.2889)
        Case "Xceltherm HT"
            FluidViscosityArray = Array(0, 46.732, 15.337, 6.913, 3.707, 2.462, 1.702, 1.21, 0.951, 0.79, 0.642, 0.534, 0.455, 0.389, 0.325)
        Case "Xceltherm 600"
            FluidViscosityArray = Array(0, 75.697, 15.489, 5.574, 2.703, 1.642, 1.085, 0.783, 0.59, 0.46, 0.368, 0.301, 0.252, 0.212)
        Case "Syltherm XLT"
            FluidViscosityArray = Array(0, 1.75, 1.1, 0.805, 0.6, 0.465, 0.37, 0.305, 0.26, 0.225, 0.19, 0.16)
        Case "Texatherm 46"
            FluidViscosityArray = Array(160.3884589, 31.02644555, 11.86849745, 6.001930126, 3.536830151, 2.295908898, 1.593275066, 1.16104712, 0.878250422, 0.684184316, 0.545848914, 0.44413353, 0.367390033)
        Case "Multitherm PG-1"
            FluidViscosityArray = Array(100, 79, 17.7, 6.8, 3.41, 2.08, 1.38, 1.01, 0.766, 0.608, 0.494, 0.409, 0.342, 0.29)
        Case "Multitherm OG-1"
            TempArray = Array(0, 20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidViscosityArray = Array(2421, 563, 199, 35.88, 13.6, 5.9, 3.5, 2.2, 1.5, 1.2, 0.9, 0.7, 0.6, 0.5)
        Case "Mobiltherm 43"
            FluidViscosityArray = Array(1671, 249, 29.6, 10.2, 4.83, 2.76, 1.79, 1.26, 0.95, 0.74, 0.6, 0.5, 0.44)
        Case "Marlotherm SH"
            FluidViscosityArray = Array(180, 157.8504111, 21.6948065, 6.073390816, 3.492279928, 2.226172047, 1.522037297, 1.167034289, 0.770403048, 0.624343293, 0.498335673, 0.399598957, 0.342009224, 0.283537197, 0.234850612)
        Case "Marlotherm LH"
            TempArray = Array(-4, 32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572, 608, 644, 680)
            FluidViscosityArray = Array(16.5, 7.6, 4, 2.5, 1.8, 1.4, 1, 0.79, 0.64, 0.54, 0.47, 0.4, 0.36, 0.32, 0.29, 0.25, 0.23, 0.21, 0.2, 0.18)
        Case "Sun 21"
            TempArray = Array(150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidViscosityArray = Array(32.41, 4.9, 2.5, 1.58, 0.99, 0.78, 0.58, 0.42, 0.32, 0.21, 0.107)
        Case "Noco 21"
            FluidViscosityArray = Array(300, 155.3067737, 30.38191525, 11.69836582, 5.943467581, 3.51503471, 2.288494896, 1.592092082, 1.162691904, 0.881170599, 0.687629291, 0.549443164, 0.447687221, 0.370809114)
        Case "Duratherm FG"
            FluidViscosityArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 600, 650)
            FluidViscosityArray = Array(432.1737293, 64.42137583, 21.15842802, 9.60288278, 5.20339424, 3.153951642, 2.06547121, 1.431440364, 1.035877578, 0.775636725, 0.470139414, 0.377372213)
        Case "Chemtherm 550"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidViscosityArray = Array(56.49513, 40.097763, 29.763288, 10.334475, 4.547169, 2.480274, 1.659134615, 1.275336504, 0.891538462, 0.740120173, 0.588701923, 0.510392606, 0.432083333)
        Case "Thermalane 550"
            TempArray = Array(0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600)
            FluidViscosityArray = Array(0.034475809, 0.034227781, 0.033979754, 0.033773064, 0.033525037, 0.03327701, 0.033028982, 0.032780955, 0.032574265, 0.032326238, 0.03207821, 0.031830183, 0.031582156, 0.031375466, 0.031127439, 0.030879411, 0.030631384, 0.030383357, 0.030176667, 0.299286396, 0.029680612, 0.029432585, 0.029184557, 0.028977868, 0.028729841, 0.028481813, 0.028233786, 0.027985758, 0.027779069, 0.027531041, 0.027283014)
        Case "Shell S2"
            TempArray = Array(32, 68, 104, 212, 302, 392, 482, 572, 644)
            FluidViscosityArray = Array(132.28, 76.764, 21.248, 3.8118, 2.31619, 0.82058, 0.8, 0.7, 0.6)
        Case "Thermoil 100"
            TempArray = Array(100, 212, 392, 500, 600)
            FluidViscosityArray = Array(86.65, 8.85, 1.57, 0.92, 0.6)
        Case "Uconhtf 500"
            TempArray = Array(0, 100, 200, 300, 400, 500)
            FluidViscosityArray = Array(2133.7, 61.511, 10.748, 5.1319, 3.1116, 1.9342)
        Case "Multitherm IG-1"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)
            FluidViscosityArray = Array(550, 199, 46, 13.63, 6.32, 3.45, 2.25, 1.56, 1.16, 0.907, 0.725, 0.598, 0.499)
        Case "Multitherm IG-4"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidViscosityArray = Array(806, 199, 38.8, 12.9, 5.93, 3.31, 2.11, 1.46, 1.08, 0.822, 0.653, 0.528, 0.433)
        Case "Hydroclear O/S 32"
            FluidViscosityArray = Array(0, 156.64, 33.64, 13.68, 7.22, 4.4, 2.939, 2.087, 1.46, 1.195, 0.93, 0.7657, 0.7, 0, 0, 0, 0)
        Case "Hydroclear O/S 46"
            FluidViscosityArray = Array(0, 220.06, 43.465, 16.83, 8.585, 5.093, 3.324, 2.317, 1.56, 1.287, 1.02, 0.805, 0.77, 0, 0, 0, 0)
        Case "Hydroclear C/S 32"
            FluidViscosityArray = Array(0, 152.399, 32.939, 13.444, 7.119, 4.348, 2.906, 2.067, 1.45, 1.186, 0.93, 0.7612, 0.71, 0, 0, 0, 0)
        Case "Hydroclear C/S 46"
            FluidViscosityArray = Array(0, 225.41, 44.52, 17.24, 8.793, 5.217, 3.405, 2.374, 1.56, 1.318, 1.02, 0.824, 0.76, 0, 0, 0, 0)
        Case "Xceltherm SST"
            FluidViscosityArray = Array(544, 50, 12, 4.7, 2.5, 1.5, 1.1, 0.82, 0.64, 0.51, 0.42, 0.35, 0.3, 0.26, 0.22, 0, 0)
        Case "Paratherm HT"
            TempArray = Array(50, 100, 200, 300, 350, 400, 450, 500, 550)
            FluidViscosityArray = Array(331, 35.5, 4.9, 1.8, 3.4, 1, 0.9, 0.7, 0.6)
        Case "Chevron 22"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidViscosityArray = Array(120, 80, 42, 10, 7.2, 4.9, 2.5, 1.5, 0.72, 0.65, 0.55, 0.5, 0.47)
        Case "Mobiltherm 605"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidViscosityArray = Array(114.484, 23.231, 9.139, 4.714, 2.821, 1.854, 1.301, 0.957, 0.73, 0.572, 0.46, 0.376)
        Case "Duratherm 600"
            TempArray = Array(15, 55, 95, 145, 195, 245, 295, 345, 395, 445, 495, 545, 600)
            FluidViscosityArray = Array(857.83, 148.89, 43.07, 14.35, 6.54, 3.63, 2.29, 1.59, 1.17, 0.9, 0.73, 0.6, 0.5)
        Case "Chevron 46"
            TempArray = Array(32, 104, 122, 212, 302, 392, 482, 572, 662)
            FluidViscosityArray = Array(419.36, 34.79, 23.147, 5.117, 2.094, 1.26, 0.9245, 0.59195, 0.3933)
         Case "Phillips 66 - OS 32"
            TempArray = Array(104, 212, 400, 500, 600)
            FluidViscosityArray = Array(23.753, 3.913, 0.9375, 0.6177, 0.4576)
        Case "Phillips 66 - OS 46"
            TempArray = Array(104, 212, 400, 500, 600)
            FluidViscosityArray = Array(34.1, 4.9356, 1.085, 0.6936, 0.4983)
        Case "Phillips 66 - CS 32"
            TempArray = Array(104, 212, 400, 500, 600)
            FluidViscosityArray = Array(23.676, 3.9, 0.934, 0.6156, 0.456)
        Case "Duratherm HTO"
            TempArray = Array(15, 55, 105, 155, 205, 255, 305, 355, 405, 455, 505, 555, 600)
            FluidViscosityArray = Array(839.77, 145.75, 32.74, 11.76, 5.62, 3.21, 2.07, 1.45, 1.08, 0.84, 0.68, 0.57, 0.49)
        Case "Multitherm IG-4"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidViscosityArray = Array(806, 199, 38.8, 12.9, 5.93, 3.31, 2.11, 1.46, 1.08, 0.822, 0.653, 0.528, 0.433)
        Case "Duratherm HF"
            TempArray = Array(40, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 640)
            FluidViscosityArray = Array(938.07, 595.89, 100.14, 29.58, 12.37, 6.46, 3.91, 2.62, 1.89, 1.44, 1.14, 0.94, 0.79, 0.7)
        Case "Seriola 1510"
            TempArray = Array(32, 50, 59, 68, 86, 104, 122, 140, 158, 176, 194, 212, 230, 248, 266, 284, 302, 320, 338, 356, 374, 392, 410, 428, 446, 464, 482, 500, 518, 536, 554, 572, 590)
            FluidViscosityArray = Array(570.6920159, 182.4700665, 119.5454639, 83.17632336, 45.64744231, 28.0892284, 18.68145181, 13.14325277, 9.649196134, 7.324451795, 5.711189511, 4.552729701, 3.696965266, 3.049551323, 2.549690171, 2.156908056, 1.843496223, 1.590017106, 1.382540372, 1.210891544, 1.067512693, 0.946703734, 0.844106486, 0.756347127, 0.680784136, 0.61532782, 0.558309231, 0.508383715, 0.464459077, 0.425641489, 0.391194322, 0.360506532, 0.333068148)
        Case "Therminol 68"
            TempArray = Array(-14, 0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 660, 680)
            FluidViscosityArray = Array(2046, 731.7, 228.6, 92.6, 45.1, 24.8, 15.1, 9.88, 6.86, 5, 3.8, 2.98, 2.39, 1.96, 1.65, 1.4, 1.21, 1.05, 0.93, 0.827, 0.74, 0.67, 0.612, 0.558, 0.513, 0.475, 0.442, 0.409, 0.384, 0.36, 0.339, 0.322, 0.302, 0.285, 0.273, 0.26)
        Case "Seriola K 3000"
            TempArray = Array(32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572)
            FluidViscosityArray = Array(119.641, 38.617, 16.609, 8.681, 5.188, 3.42, 2.421, 1.808, 1.399, 1.128, 0.929, 0.786, 0.676, 0.589, 0.518, 0.462)
    End Select

'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
   If FluidViscosityArray(i + 1) = -1 Then
        FluidViscosity = "Beyond Fluid Limits"
        Exit Function
    End If
    FluidViscosity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (FluidViscosityArray(i + 1) - FluidViscosityArray(i)) + FluidViscosityArray(i)
    i = i + 1
Loop
        
End Function
Function FluidThermalConductivity(Fluid As String, Temperature As Single)

Dim TempArray, FluidThermalConductivityArray As Variant

TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)


'all of the fluids go here
    Select Case Fluid
        Case "Shell Thermia C"
            FluidThermalConductivityArray = Array(0, 0.076429535, 0.075276867, 0.0741242, 0.072971532, 0.071818865, 0.070666198, 0.06951353, 0.068360863, 0.067208195, 0.066055528, 0.06490286, 0.063750193)
        Case "Shell Thermia B"
            FluidThermalConductivityArray = Array(0, 0.076429535, 0.075276867, 0.0741242, 0.072971532, 0.071818865, 0.070666198, 0.06951353, 0.068360863, 0.067208195, 0.066055528, 0.06490286, 0.063750193)
        Case "Calflo AF"
            FluidThermalConductivityArray = Array(0.087, 0.10617485, 0.1206188, 0.13088595, 0.1375304, 0.14110625, 0.1421676, 0.14126855, 0.1389632, 0.13580565, 0.13235, 0.12915035, 0.1267608)
        Case "Calflo FG"
            FluidThermalConductivityArray = Array(0.196, 0.18014745, 0.1671296, 0.15661615, 0.1482768, 0.14178125, 0.1367992, 0.13300035, 0.1300544, 0.12763105, 0.1254, 0.12303095, 0.1201936, 0.11655765)
        Case "Calflo HTF"
            FluidThermalConductivityArray = Array(0.084, 0.083, 0.0818, 0.0808, 0.0798, 0.0788, 0.0778, 0.0768, 0.0758, 0.0748, 0.0738, 0.0728, 0.0718)
        Case "Calflo LT"
            FluidThermalConductivityArray = Array(0.18, 0.1609185, 0.156837, 0.1527555, 0.148674, 0.1445925, 0.140511, 0.1364295, 0.132348, 0.1282665, 0.124185, 0.1201035)
        Case "Dowtherm A"
            FluidThermalConductivityArray = Array(0.083, 0.0805, 0.078, 0.0755, 0.073, 0.0705, 0.068, 0.0655, 0.063, 0.0605, 0.058, 0.0555, 0.053, 0.0505, 0.048, 0.0455, 0.043)
        Case "Dowtherm G"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 725)
            FluidThermalConductivityArray = Array(0.0737, 0.0718, 0.07, 0.0681, 0.0663, 0.0644, 0.0625, 0.0607, 0.0588, 0.057, 0.0551, 0.0532, 0.0514, 0.0495, 0.0486)
        Case "Dowtherm Q"
            FluidThermalConductivityArray = Array(0.073, 0.0712, 0.0693, 0.0672, 0.065, 0.0627, 0.0604, 0.058, 0.0555, 0.053, 0.0505, 0.048, 0.0455, 0.044, -1, -1, -1)
        Case "Dowtherm RP"
            TempArray = Array(50, 150, 250, 350, 450, 550, 650)
            FluidThermalConductivityArray = Array(0.0766, 0.0725, 0.0683, 0.0642, 0.06, 0.0558, 0.0517)
        Case "Dowtherm HT"
            TempArray = Array(50, 150, 250, 350, 450, 550, 650)
            FluidThermalConductivityArray = Array(0.0739, 0.0711, 0.0682, 0.0652, 0.0619, 0.0584, 0.0545)
        Case "Dowtherm J"
            FluidThermalConductivityArray = Array(0.0788, 0.0754, 0.072, 0.0686, 0.0652, 0.0618, 0.0584, 0.055, 0.0516, 0.0482, 0.0448, 0.0414, 0.0379, 0.0345, -1, -1, -1)
        Case "Dowtherm MX"
            TempArray = Array(-10, 50, 110, 170, 230, 290, 350, 410, 470, 530, 590, 650)
            FluidThermalConductivityArray = Array(0.0735, 0.0715, 0.0695, 0.0675, 0.0655, 0.0635, 0.0615, 0.0595, 0.0575, 0.0555, 0.0535, 0.0514)
        Case "Dowtherm T"
            TempArray = Array(20, 100, 180, 260, 340, 420, 500, 580, 600)
            FluidThermalConductivityArray = Array(0.0813, 0.0756, 0.0699, 0.0642, 0.0585, 0.0528, 0.0471, 0.0414, 0.04)
        Case "Mobiltherm 603"
            FluidThermalConductivityArray = Array(0.078591227, 0.078324, 0.077007, 0.075979, 0.074759, 0.07359, 0.072306, 0.071021, 0.069788, 0.068825, 0.067746, 0.066461, 0.065446, -1, -1, -1, -1)
        Case "Therminol 55"
            FluidThermalConductivityArray = Array(0.0768, 0.0749, 0.0731, 0.0712, 0.0693, 0.0675, 0.0656, 0.0637, 0.0618, 0.0599, 0.058, 0.05615, 0.0542, 0.048759706, -1, -1, -1, -1)
        Case "Therminol 59"
            FluidThermalConductivityArray = Array(0.0716, 0.07055, 0.0694, 0.06815, 0.0688, 0.06525, 0.063573269, 0.06185, 0.059948736, 0.057954095, 0.0559, 0.05365, 0.0513, 0.048759706, -1, -1, -1, -1)
        Case "Therminol 66"
            TempArray = Array(20, 30, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 650, 660, 680, 700)
            FluidThermalConductivityArray = Array(0.06845, 0.068377, 0.0683, 0.0681, 0.0678, 0.0675, 0.0672, 0.0669, 0.0666, 0.0662, 0.0658, 0.0654, 0.065, 0.0646, 0.0641, 0.0636, 0.0631, 0.0625, 0.062, 0.0614, 0.0608, 0.0602, 0.0595, 0.0588, 0.0581, 0.0574, 0.0567, 0.0559, 0.0551, 0.0543, 0.0535, 0.0527, 0.0518, 0.05135, 0.0509, 0.05, 0.0491)
        Case "Therminol SP"
            FluidThermalConductivityArray = Array(0.076837, 0.074945915, 0.07305483, 0.071163745, 0.06927266, 0.06745, 0.0656, 0.0637, 0.0618, 0.0599, 0.058, 0.05615, 0.0542, 0.052252895, 0.05036181, 0.048470725, 0.04657964)
        Case "Therminol D-12"
            FluidThermalConductivityArray = Array(0.0668, 0.06445, 0.062, 0.05935, 0.0565, 0.0536, 0.0505, 0.04735, 0.044, 0.04045, 0.0369, 0.0334)
        Case "Therminol VP-1"
            FluidThermalConductivityArray = Array(0.081, 0.0793, 0.0778, 0.07615, 0.0743, 0.0723, 0.0701, 0.06785, 0.0654, 0.06275, 0.06, 0.0571, 0.054, 0.05075, 0.0474, 0.0424, 0.0402)
        Case "Therminol 62"
            FluidThermalConductivityArray = Array(0.0729, 0.071549997, 0.0702, 0.068800002, 0.0673, 0.065799996, 0.0643, 0.062650003, 0.061, 0.059150003, 0.0572, 0.055149999, 0.0528, 0.0518, 0.0507, 0.05036118)
        Case "Therminol 72"
            FluidThermalConductivityArray = Array(0.0834, 0.0814, 0.0795, 0.0775, 0.0756, 0.0736, 0.0717, 0.0697, 0.0678, 0.0658, 0.0639, 0.0619, 0.06, 0.058, 0.0561, 0.0555)
        Case "Therminol 75"
             TempArray = Array(0, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700, 725)
             FluidThermalConductivityArray = Array(0.0823, 0.0788, 0.0781, 0.0775, 0.0769, 0.0763, 0.0756, 0.075, 0.0744, 0.0738, 0.0731, 0.0725, 0.0719, 0.0713, 0.0706, 0.0699, 0.0693, 0.0686, 0.0679, 0.0671, 0.0664, 0.0656, 0.0649, 0.064, 0.0632, 0.0624, 0.0615, 0.0606, 0.0596)
        Case "Therminol XP"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidThermalConductivityArray = Array(0.0681, 0.0671, 0.066, 0.06485, 0.0635, 0.0621, 0.0605, 0.05885, 0.0571, 0.05525, 0.0533, 0.0512, 0.049)
        Case "Paratherm NF"
            FluidThermalConductivityArray = Array(0, 0.150068739, 0.14667441, 0.143225196, 0.13972928, 0.136194844, 0.13263007, 0.129043141, 0.12544224, 0.121835549, 0.11823125, 0.114637526, 0.11106256, 0.107514534)
        Case "Paratherm MG"
            FluidThermalConductivityArray = Array(0.083, 0.082, 0.081, 0.08, 0.079, 0.078, 0.077, 0.076, 0.075, 0.074, 0.073, -1, -1, -1, -1, -1, -1)
        Case "Paratherm MR"
            FluidThermalConductivityArray = Array(0.0853, 0.0839, 0.0825, 0.0811, 0.0797, 0.0783, 0.0769, 0.0755, 0.0741, 0.0727, 0.0713, 0.0699, 0.0685)
        Case "Paratherm HE"
            FluidThermalConductivityArray = Array(0.0783, 0.0771, 0.0759, 0.0747, 0.0735, 0.0723, 0.0711, 0.0699, 0.0687, 0.0675, 0.0662, 0.065, 0.0638, 0.0636, 0.0635, 0.0633, 0.0631, 0.06271)
        Case "Paratherm HR"
            TempArray = Array(0, 25, 50, 75, 100, 125, 150, 175, 200, 225, 250, 275, 300, 325, 350, 375, 400, 425, 450, 475, 500, 525, 550, 575, 600, 625, 650, 675, 700)
            FluidThermalConductivityArray = Array(0.068, 0.068, 0.068, 0.068, 0.068, 0.068, 0.067, 0.067, 0.066, 0.066, 0.065, 0.065, 0.064, 0.064, 0.063, 0.062, 0.061, 0.06, 0.059, 0.059, 0.057, 0.056, 0.055, 0.054, 0.053, 0.052, 0.05, 0.049, 0.047)
        Case "Petrotherm"
            FluidThermalConductivityArray = Array(0.17, 0.164674925, 0.1596797, 0.155014325, 0.1506788, 0.146673125, 0.1429973, 0.139651325, 0.1366352, 0.133948925, 0.1315925, 0.129565925, 0.1278692)
        Case "PetroCanada FG"
            FluidThermalConductivityArray = Array(0, 0.08, 0.0783, 0.077649996, 0.077, 0.076250002, 0.0755, 0.074649997, 0.0738, 0.072850004, 0.0719, 0.070950001, 0.07)
        Case "Syltherm 800"
            FluidThermalConductivityArray = Array(0.19, 0.180511336, 0.171043529, 0.161593855, 0.152159592, 0.142738016, 0.133326403, 0.123922031, 0.114522176, 0.105124115, 0.095725125, 0.086322482, 0.076913464, 0.067495347, 0.058065407, 0.048620922, 0.039159168)
        Case "TEG 0.6"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidThermalConductivityArray = Array(0.2123, 0.2077, 0.203, 0.1984, 0.1938, 0.1891, 0.1845, 0.1799, 0.1753)
        Case "TEG 0.98"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400)
            FluidThermalConductivityArray = Array(0.15, 0.140882975, 0.1321708, 0.123696825, 0.1152944, 0.106796875, 0.0980376, 0.088849925, 0.0766)
        Case "EGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.238, 0.243, 0.247, 0.251, 0.255, 0.259, 0.263, 0.266, 0.269, 0.272, 0.275, 0.277, 0.28, 0.282, 0.284, 0.285, 0.287, 0.288, 0.289, 0.29, 0.291, 0.291, 0.291, 0.291, 0.291, 0.291, 0.29, 0.289, 0.288, 0.287, 0.286, 0.284, 0.282, 0.28, 0.278)
        Case "EGlycol 0.4"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.216, 0.22, 0.224, 0.227, 0.231, 0.234, 0.237, 0.24, 0.243, 0.246, 0.248, 0.251, 0.253, 0.255, 0.256, 0.258, 0.259, 0.261, 0.262, 0.263, 0.263, 0.264, 0.265, 0.265, 0.265, 0.265, 0.264, 0.264, 0.263, 0.263, 0.262, 0.261, 0.269, 0.258, 0.256, 0.254)
        Case "EGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.193, 0.197, 0.2, 0.204, 0.207, 0.21, 0.212, 0.215, 0.218, 0.22, 0.223, 0.225, 0.227, 0.229, 0.23, 0.232, 0.233, 0.235, 0.236, 0.237, 0.238, 0.239, 0.24, 0.24, 0.24, 0.241, 0.241, 0.241, 0.241, 0.24, 0.24, 0.239, 0.239, 0.238, 0.237, 0.236, 0.235, 0.233)
         Case "Dynalene EG-XT 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 120, 140, 160, 180, 200, 220)
            FluidThermalConductivityArray = Array(0.193, 0.197, 0.2, 0.204, 0.207, 0.21, 0.212, 0.215, 0.218, 0.22, 0.223, 0.225, 0.227, 0.23, 0.233, 0.236, 0.238, 0.24, 0.24)
        Case "EGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.178, 0.181, 0.184, 0.186, 0.189, 0.191, 0.194, 0.196, 0.198, 0.2, 0.202, 0.204, 0.206, 0.208, 0.209, 0.21, 0.212, 0.213, 0.214, 0.215, 0.216, 0.217, 0.218, 0.218, 0.219, 0.219, 0.219, 0.219, 0.22, 0.219, 0.219, 0.219, 0.219, 0.218, 0.218, 0.217, 0.216, 0.215, 0.214)
        Case "EGlycol 0.7"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.167, 0.17, 0.172, 0.174, 0.176, 0.178, 0.18, 0.182, 0.183, 0.185, 0.186, 0.188, 0.189, 0.19, 0.192, 0.193, 0.194, 0.195, 0.196, 0.197, 0.197, 0.198, 0.199, 0.199, 0.2, 0.2, 0.2, 0.2, 0.201, 0.201, 0.201, 0.2, 0.2, 0.2, 0.2, 0.199, 0.199, 0.198, 0.197)
        Case "EGlycol 0.8"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.163, 0.164, 0.166, 0.167, 0.169, 0.17, 0.171, 0.172, 0.173, 0.174, 0.175, 0.176, 0.177, 0.178, 0.179, 0.18, 0.18, 0.181, 0.181, 0.182, 0.182, 0.183, 0.183, 0.183, 0.184, 0.184, 0.69, 0.63, 0.58, 0.53, 0.5, 0.46, 0.43, 0.4, 0.37, 0.35)
        Case "EGlycol 0.9"
            TempArray = Array(0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.153, 0.154, 0.155, 0.156, 0.157, 0.158, 0.159, 0.16, 0.161, 0.161, 0.162, 0.163, 0.163, 0.164, 0.165, 0.165, 0.166, 0.166, 0.167, 0.167, 0.168, 0.168, 0.168, 0.169, 0.169, 0.169, 0.169, 0.169, 0.17, 0.69, 0.63, 0.59, 0.55, 0.51, 0.48, 0.45)
        Case "PGlycol 0.3"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidThermalConductivityArray = Array(0.235, 0.239, 0.243, 0.247, 0.251, 0.254, 0.258, 0.261, 0.263, 0.266, 0.268, 0.27, 0.272, 0.274, 0.276, 0.277, 0.278, 0.279, 0.28, 0.28, 0.28, 0.28, 0.28, 0.28, 0.279, 0.278, 0.277, 0.276, 0.275, 0.273, 0.271, 0.269, 0.268)
        Case "PGlycol 0.4"
            TempArray = Array(10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidThermalConductivityArray = Array(0.211, 0.215, 0.218, 0.222, 0.225, 0.227, 0.23, 0.233, 0.235, 0.237, 0.239, 0.241, 0.243, 0.244, 0.245, 0.246, 0.247, 0.248, 0.249, 0.249, 0.249, 0.249, 0.249, 0.249, 0.249, 0.24, 0.247, 0.246, 0.245, 0.244, 0.242, 0.241, 0.239, 0.238)
        Case "PGlycol 0.5"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.188, 0.191, 0.194, 0.196, 0.199, 0.201, 0.204, 0.206, 0.208, 0.21, 0.211, 0.213, 0.214, 0.215, 0.217, 0.218, 0.218, 0.219, 0.22, 0.22, 0.221, 0.221, 0.221, 0.221, 0.22, 0.22, 0.22, 0.219, 0.218, 0.217, 0.216, 0.215, 0.214, 0.212, 0.211, 0.21)
        Case "PGlycol 0.6"
            TempArray = Array(-30, -20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 325)
            FluidThermalConductivityArray = Array(0.171, 0.174, 0.176, 0.178, 0.179, 0.181, 0.183, 0.184, 0.186, 0.187, 0.188, 0.189, 0.19, 0.191, 0.192, 0.193, 0.193, 0.194, 0.194, 0.194, 0.195, 0.195, 0.195, 0.194, 0.194, 0.194, 0.193, 0.193, 0.192, 0.191, 0.191, 0.19, 0.188, 0.187, 0.186, 0.185, 0.184)
        Case "PGlycol 0.9"
            TempArray = Array(-20, -10, 0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160, 170, 180, 190, 200, 210, 220, 230, 240, 250, 260, 270, 280, 290, 300, 310, 320, 330, 340, 350)
            FluidThermalConductivityArray = Array(0.137, 0.136, 0.136, 0.136, 0.136, 0.135, 0.135, 0.135, 0.134, 0.134, 0.134, 0.133, 0.133, 0.132, 0.132, 0.131, 0.131, 0.13, 0.13, 0.129, 0.129, 0.128, 0.127, 0.127, 0.126, 0.125, 0.125, 0.124)
        Case "Salt"
            TempArray = Array(300, 400, 500, 600, 700, 800, 900, 1000, 1100, 1200)
            FluidThermalConductivityArray = Array(0.255, 0.251, 0.243, 0.223, 0.201, 0.181, 0.171666667, 0.156238095, 0.140809524)
        Case "Xceltherm HT"
            FluidThermalConductivityArray = Array(0, 0.0765, 0.075, 0.0735, 0.072, 0.0705, 0.069, 0.0675, 0.066, 0.0645, 0.063, 0.0615, 0.06, 0.0585, 0.057)
        Case "Xceltherm 600"
            FluidThermalConductivityArray = Array(0.158, 0.154453634, 0.15079737, 0.147044461, 0.14320816, 0.139301719, 0.13533839, 0.131331426, 0.12729408, 0.123239604, 0.11918125, 0.115132271, 0.11110592, 0.107115449)
        Case "Syltherm XLT"
            FluidThermalConductivityArray = Array(0.158, 0.152649, 0.146117, 0.1384805, 0.129816, 0.1202, 0.109709, 0.0984195, 0.086408, 0.073751, 0.060525, 0.0468065)
        Case "Texatherm 46"
            FluidThermalConductivityArray = Array(0, 0.07644152, 0.07531804, 0.07419456, 0.07307108, 0.0719476, 0.07082412, 0.06970064, 0.06857716, 0.06745368, 0.0663302, 0.06520672, 0.06408324, 0.06295976)
        Case "Multitherm PG-1"
            FluidThermalConductivityArray = Array(0.0784, 0.0773, 0.0761, 0.0749, 0.0738, 0.0726, 0.0715, 0.0703, 0.0691, 0.068, 0.0668, 0.0657, 0.0645, 0.063346154)
        Case "Multitherm OG-1"
            TempArray = Array(0, 20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidThermalConductivityArray = Array(0.0818, 0.0815, 0.0794, 0.079, 0.0787, 0.0783, 0.0779, 0.0769, 0.0763, 0.0757, 0.0753, 0.075, 0.0747, 0.074)
        Case "Mobiltherm 43"
            FluidThermalConductivityArray = Array(0.077916667, 0.076666667, 0.075166667, 0.073583333, 0.072083333, 0.0705, 0.068916667, 0.067416667, 0.065833333, 0.064333333, 0.06275, 0.06125, 0.059666667)
        Case "Marlotherm SH"
            FluidThermalConductivityArray = Array(0.169, 0.162291883, 0.155623135, 0.148987159, 0.14237736, 0.135787141, 0.129209905, 0.122639057, 0.116068, 0.109490138, 0.102898875, 0.096287614, 0.08964976, 0.082978716, 0.076267885)
        Case "Marlotherm LH"
            TempArray = Array(-4, 32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572, 608, 644, 680)
            FluidThermalConductivityArray = Array(0.079, 0.077, 0.076, 0.075, 0.073, 0.072, 0.07, 0.069, 0.056, 0.066, 0.065, 0.064, 0.063, 0.061, 0.06, 0.059, 0.057, 0.056, 0.055, 0.053)
        Case "Sun 21"
            TempArray = Array(150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidThermalConductivityArray = Array(0.0746, 0.0735, 0.0723, 0.0709, 0.0694, 0.068, 0.0671, 0.0662, 0.0652, 0.064, 0.06256)
        Case "Duratherm FG"
            FluidThermalConductivityArray = Array(0.084, 0.0832125, 0.082, 0.0808625, 0.0798, 0.0788125, 0.0779, 0.0770625, 0.0763, 0.0756125, 0.075, 0.0744625, 0.074, 0.0736125)
        Case "Noco 21"
            TempArray = Array(50, 150, 300, 450, 600)
            FluidThermalConductivityArray = Array(0.079, 0.075, 0.071, 0.067, 0.064)
        Case "Chemtherm 550"
            TempArray = Array(0, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidThermalConductivityArray = Array(0.077, 0.076, 0.075, 0.0738, 0.0725, 0.0712, 0.07, 0.069, 0.0678, 0.0665, 0.0652, 0.064, 0.063)
        Case "Thermalane 550"
            TempArray = Array(0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600)
            FluidThermalConductivityArray = Array(0.0834, 0.0828, 0.0822, 0.0817, 0.0811, 0.0805, 0.0799, 0.0793, 0.0788, 0.0782, 0.0776, 0.077, 0.0764, 0.0759, 0.0753, 0.0747, 0.0741, 0.0735, 0.073, 0.724, 0.0718, 0.0712, 0.0706, 0.0701, 0.0695, 0.0689, 0.0683, 0.0677, 0.0672, 0.0666, 0.066)
        Case "Shell S2"
            TempArray = Array(32, 68, 104, 212, 302, 392, 482, 572, 644)
            FluidThermalConductivityArray = Array(0.078579304, 0.077423726, 0.076845937, 0.073956992, 0.072223625, 0.069912469, 0.068179102, 0.065867946, 0.064134579)
        Case "Thermoil 100"
            TempArray = Array(0, 100, 212, 392, 500, 600)
            FluidThermalConductivityArray = Array(0.1, 0.075, 0.073, 0.068, 0.066, 0.064)
        Case "Uconhtf 500"
            TempArray = Array(0, 100, 200, 300, 400, 500)
            FluidThermalConductivityArray = Array(0.1, 0.095, 0.09, 0.0853, 0.08, 0.076)
        Case "Multitherm IG-1"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650, 700, 750, 800)
            FluidThermalConductivityArray = Array(0.0815, 0.0808, 0.0795, 0.0783, 0.077, 0.0757, 0.0745, 0.0733, 0.072, 0.0708, 0.0695, 0.0683, 0.0672)
         Case "Multitherm IG-4"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidThermalConductivityArray = Array(0.0786, 0.0779, 0.0768, 0.0756, 0.0744, 0.0732, 0.0721, 0.0709, 0.0697, 0.0685, 0.0674, 0.06622, 0.065)
        Case "Hydroclear O/S 32"
            FluidThermalConductivityArray = Array(0.0835, 0.0827, 0.0819, 0.0811, 0.0802, 0.0795, 0.0787, 0.0778, 0.077, 0.0762, 0.0754, 0.0746, 0.0738, 0, 0, 0, 0)
        Case "Hydroclear O/S 46"
            FluidThermalConductivityArray = Array(0.0835, 0.0827, 0.0819, 0.0811, 0.0803, 0.0795, 0.0787, 0.0778, 0.077, 0.0762, 0.0754, 0.0746, 0.0738, 0, 0, 0, 0)
        Case "Hydroclear C/S 32"
            FluidThermalConductivityArray = Array(0.0835, 0.0827, 0.0819, 0.0811, 0.0802, 0.0795, 0.0787, 0.0778, 0.077, 0.0762, 0.0754, 0.0746, 0.0738, 0, 0, 0, 0)
        Case "Hydroclear C/S 46"
            FluidThermalConductivityArray = Array(0.0835, 0.0827, 0.0819, 0.0811, 0.0803, 0.0795, 0.0787, 0.0778, 0.077, 0.0762, 0.0754, 0.0746, 0.0738, 0, 0, 0, 0)
        Case "Xceltherm SST"
            FluidThermalConductivityArray = Array(0.068, 0.068, 0.068, 0.067, 0.066, 0.065, 0.064, 0.063, 0.061, 0.059, 0.057, 0.055, 0.053, 0.05, 0.047, 0, 0)
        Case "Paratherm HT"
            TempArray = Array(50, 100, 200, 300, 350, 400, 450, 500, 550)
            FluidThermalConductivityArray = Array(0.0674, 0.066, 0.0631, 0.0602, 0.0587, 0.0573, 0.0558, 0.0544, 0.0529)
        Case "Chevron 22"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 650)
            FluidThermalConductivityArray = Array(0.0748, 0.0723, 0.0698, 0.0673, 0.0648, 0.0623, 0.0598, 0.0573, 0.0548, 0.0523, 0.0498, 0.0473, 0.0448)
        Case "Mobiltherm 605"
            TempArray = Array(50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidThermalConductivityArray = Array(0.0778, 0.0768, 0.0758, 0.0748, 0.0738, 0.0728, 0.0718, 0.0708, 0.0698, 0.0688, 0.0678, 0.0668)
        Case "Duratherm 600"
            TempArray = Array(15, 55, 95, 145, 195, 245, 295, 345, 395, 445, 495, 545, 600)
            FluidThermalConductivityArray = Array(0.082, 0.081, 0.08, 0.08, 0.079, 0.078, 0.077, 0.076, 0.075, 0.074, 0.074, 0.073, 0.072)
        Case "Chevron 46"
            TempArray = Array(32, 104, 122, 212, 302, 392, 482, 572, 662)
            FluidThermalConductivityArray = Array(0.0782, 0.0765, 0.0761, 0.072, 0.074, 0.0698, 0.0677, 0.0656, 0.0634)
        Case "Phillips 66 - OS 32"
            TempArray = Array(60, 100, 320, 550)
            FluidThermalConductivityArray = Array(0.081, 0.079, 0.074, 0.067)
        Case "Phillips 66 - OS 46"
            TempArray = Array(60, 100, 320, 550)
            FluidThermalConductivityArray = Array(0.081, 0.079, 0.074, 0.067)
        Case "Phillips 66 - CS 32"
            TempArray = Array(60, 100, 320, 550)
            FluidThermalConductivityArray = Array(0.081, 0.079, 0.074, 0.067)
        Case "Duratherm HTO"
            TempArray = Array(15, 55, 105, 155, 205, 255, 305, 355, 405, 455, 505, 555, 600)
            FluidThermalConductivityArray = Array(0.08, 0.08, 0.079, 0.078, 0.077, 0.076, 0.075, 0.074, 0.074, 0.073, 0.072, 0.071, 0.07)
        Case "Multitherm IG-4"
            TempArray = Array(20, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600)
            FluidThermalConductivityArray = Array(0.0786, 0.0779, 0.0768, 0.0756, 0.0744, 0.0732, 0.0721, 0.0709, 0.0697, 0.0685, 0.0674, 0.0662, 0.065)
        Case "Duratherm HF"
            TempArray = Array(40, 50, 100, 150, 200, 250, 300, 350, 400, 450, 500, 550, 600, 640)
            FluidThermalConductivityArray = Array(0.086, 0.086, 0.086, 0.086, 0.085, 0.085, 0.085, 0.085, 0.084, 0.084, 0.084, 0.083, 0.083, 0.083)
        Case "Seriola 1510"
            TempArray = Array(32, 50, 59, 68, 86, 104, 122, 140, 158, 176, 194, 212, 230, 248, 266, 284, 302, 320, 338, 356, 374, 392, 410, 428, 446, 464, 482, 500, 518, 536, 554, 572, 590)
            FluidThermalConductivityArray = Array(0.0783, 0.07772, 0.07772, 0.07772, 0.07714, 0.07656, 0.07656, 0.07598, 0.0754, 0.07482, 0.07482, 0.07424, 0.07366, 0.07366, 0.07308, 0.0725, 0.0725, 0.07192, 0.07134, 0.07076, 0.07076, 0.07018, 0.0696, 0.0696, 0.06902, 0.06844, 0.06844, 0.06786, 0.06728, 0.0667, 0.0667, 0.06612, 0.06554)
        Case "Therminol 68"
            TempArray = Array(-14, 0, 20, 40, 60, 80, 100, 120, 140, 160, 180, 200, 220, 240, 260, 280, 300, 320, 340, 360, 380, 400, 420, 440, 460, 480, 500, 520, 540, 560, 580, 600, 620, 640, 660, 680)
            FluidThermalConductivityArray = Array(0.0736, 0.0732, 0.0727, 0.0722, 0.0716, 0.0711, 0.0706, 0.0701, 0.0696, 0.0691, 0.0685, 0.068, 0.0675, 0.067, 0.0665, 0.0659, 0.0654, 0.0649, 0.0644, 0.0639, 0.0634, 0.0628, 0.0623, 0.0618, 0.0613, 0.0608, 0.0602, 0.0597, 0.0592, 0.0587, 0.0582, 0.0577, 0.0571, 0.0566, 0.0561, 0.0556)
        Case "Seriola K 3000"
            TempArray = Array(32, 68, 104, 140, 176, 212, 248, 284, 320, 356, 392, 428, 464, 500, 536, 572)
            FluidThermalConductivityArray = Array(0.081, 0.08, 0.079, 0.079, 0.078, 0.077, 0.076, 0.075, 0.074, 0.073, 0.073, 0.072, 0.071, 0.07, 0.069, 0.068)
    End Select

'interpolation between tabulated temperature values
i = 0
Do While Temperature > TempArray(i)
   If FluidThermalConductivityArray(i + 1) = -1 Then
        FluidThermalConductivity = "Beyond Fluid Limits"
        Exit Function
    End If
    FluidThermalConductivity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (FluidThermalConductivityArray(i + 1) - FluidThermalConductivityArray(i)) + FluidThermalConductivityArray(i)
    i = i + 1
Loop
        
End Function


Function LiquidAmineCp(Temperature As Single, Pressure As Single)
Dim TempArray, cPArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        cPArray = Array(0.9089, 0.9164, 0.9247, 0.9339, 0.939, 0.9401, 0.9389, 0.9373, 0.9356, 0.9338, 0.932, 0.9286, 0.925, 0.9212, 0.9173, 0.9131, 0.9087, 0.9041, 0.8992, 0.8951, 0.8909, 0.8865, 0.8819, 0.8771, 0.873, 0.8688, 0.8644, 0.86, 0.8554, 0.8508)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If cPArray(i + 1) = -1 Then
                LiquidAmineCp = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineCp = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (cPArray(i + 1) - cPArray(i)) + cPArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineCp = "only works at 20 psig fool"
    
End Select

End Function

Function LiquidAmineViscosity(Temperature As Single, Pressure As Single)
Dim TempArray, ViscosityArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        ViscosityArray = Array(1.4045, 1.1914, 1.0235, 0.8893, 0.8314, 0.819, 0.8294, 0.8433, 0.8586, 0.8749, 0.892, 0.9241, 0.9587, 0.9962, 1.0367, 1.0804, 1.1274, 1.178, 1.2321, 1.278, 1.3259, 1.3757, 1.4269, 1.4789, 1.5222, 1.5647, 1.6055, 1.6434, 1.677, 1.7047)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If ViscosityArray(i + 1) = -1 Then
                LiquidAmineViscosity = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineViscosity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (ViscosityArray(i + 1) - ViscosityArray(i)) + ViscosityArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineViscosity = "only works at 20 psig fool"
    
End Select

End Function

Function LiquidAmineDensity(Temperature As Single, Pressure As Single)
Dim TempArray, densityArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        densityArray = Array(60.0733, 59.5882, 59.0724, 58.5247, 58.2163, 58.0904, 58.0661, 58.056, 58.0495, 58.0444, 58.0399, 58.0332, 58.0272, 58.0218, 58.0169, 58.0125, 58.0086, 58.0053, 58.0026, 58.001, 57.9998, 57.9991, 57.9989, 57.9991, 57.9995, 58.0002, 58.001, 58.0018, 58.0022, 58.0022)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If densityArray(i + 1) = -1 Then
                LiquidAmineDensity = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineDensity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (densityArray(i + 1) - densityArray(i)) + densityArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineDensity = "only works at 20 psig fool"
    
End Select

End Function

Function AmineEnthalpy(Temperature As Single, Pressure As Single)
Dim TempArray, LiquidEnthalpyArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        LiquidEnthalpyArray = Array(-476.79, -459.64, -442.34, -424.88, -414.05, -403.23, -392.4, -381.58, -370.75, -359.93, -349.1, -330.16, -311.21, -292.27, -273.33, -254.38, -235.44, -216.5, -197.55, -182.4, -167.24, -152.09, -136.93, -121.78, -109.15, -96.519, -83.89, -71.261, -58.632, -46.002)
        
        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If LiquidEnthalpyArray(i + 1) = -1 Then
                AmineEnthalpy = "Beyond Fluid Limits"
                Exit Function
            End If
            AmineEnthalpy = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (LiquidEnthalpyArray(i + 1) - LiquidEnthalpyArray(i)) + LiquidEnthalpyArray(i)
            i = i + 1
        Loop
    Case Else
        AmineEnthalpy = "only works at 20 psig fool"
    
End Select

End Function

Function LiquidAmineThermalConductivity(Temperature As Single, Pressure As Single)
Dim TempArray, TCArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        TCArray = Array(0.1805, 0.1797, 0.1788, 0.1776, 0.1767, 0.1753, 0.1736, 0.1718, 0.1701, 0.1683, 0.1665, 0.1632, 0.1599, 0.1565, 0.1531, 0.1496, 0.146, 0.1423, 0.1386, 0.1356, 0.1326, 0.1295, 0.1264, 0.1232, 0.1206, 0.1179, 0.1153, 0.1126, 0.11, 0.1073)


        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If TCArray(i + 1) = -1 Then
                LiquidAmineThermalConductivity = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineThermalConductivity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (TCArray(i + 1) - TCArray(i)) + TCArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineThermalConductivity = "only works at 20 psig fool"
    
End Select

End Function

Function LiquidAmineSurfaceTension(Temperature As Single, Pressure As Single)
Dim TempArray, STArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        STArray = Array(56.1684, 54.2749, 52.3397, 50.3645, 49.2932, 48.8075, 48.6378, 48.5081, 48.3851, 48.2618, 48.1358, 47.9054, 47.6593, 47.3948, 47.1093, 46.8, 46.4639, 46.0975, 45.6967, 45.3482, 44.9721, 44.5651, 44.1238, 43.6443, 43.2124, 42.748, 42.2481, 41.7092, 41.1279, 40.5004)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If STArray(i + 1) = -1 Then
                LiquidAmineSurfaceTension = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineSurfaceTension = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (STArray(i + 1) - STArray(i)) + STArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineSurfaceTension = "only works at 20 psig fool"
    
End Select

End Function

Function LiquidAmineCriticalPressure(Temperature As Single, Pressure As Single)
Dim TempArray, cPArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        cPArray = Array(2848.72, 2848.72, 2848.72, 2848.72, 2848.25, 2843.99, 2837.49, 2830.43, 2823.03, 2815.3, 2807.25, 2792.32, 2776.25, 2758.9, 2740.14, 2719.8, 2697.67, 2673.54, 2647.13, 2624.16, 2599.36, 2572.51, 2543.39, 2511.73, 2483.19, 2452.49, 2419.42, 2383.76, 2345.27, 2303.7)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If cPArray(i + 1) = -1 Then
                LiquidAmineCriticalPressure = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineCriticalPressure = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (cPArray(i + 1) - cPArray(i)) + cPArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineCriticalPressure = "only works at 20 psig fool"
    
End Select

End Function

Function LiquidAmineCriticalTemperature(Temperature As Single, Pressure As Single)
Dim TempArray, CTArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        CTArray = Array(712.19, 712.19, 712.19, 712.19, 712.35, 712.5, 712.64, 712.79, 712.95, 713.11, 713.28, 713.59, 713.93, 714.29, 714.68, 715.11, 715.57, 716.08, 716.63, 717.11, 717.63, 718.2, 718.81, 719.47, 720.07, 720.72, 721.42, 722.17, 722.98, 723.86)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If CTArray(i + 1) = -1 Then
                LiquidAmineCriticalTemperature = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineCriticalTemperature = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (CTArray(i + 1) - CTArray(i)) + CTArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineCriticalTemperature = "only works at 20 psig fool"
    
End Select

End Function

Function LiquidAmineMoleWeight(Temperature As Single, Pressure As Single)
Dim TempArray, MWArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(200, 218.793, 237.587, 256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        MWArray = Array(31.39, 31.39, 31.39, 31.39, 31.42, 31.58, 31.83, 32.1, 32.39, 32.68, 32.99, 33.56, 34.18, 34.84, 35.56, 36.34, 37.19, 38.11, 39.13, 40.01, 40.96, 41.99, 43.1, 44.32, 45.41, 46.59, 47.86, 49.23, 50.7, 52.3)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If MWArray(i + 1) = -1 Then
                LiquidAmineMoleWeight = "Beyond Fluid Limits"
                Exit Function
            End If
            LiquidAmineMoleWeight = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (MWArray(i + 1) - MWArray(i)) + MWArray(i)
            i = i + 1
        Loop
    Case Else
        LiquidAmineMoleWeight = "only works at 20 psig fool"
    
End Select

End Function

Function VaporAmineDensity(Temperature As Single, Pressure As Single)
Dim TempArray, densityArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        densityArray = Array(0.1096, 0.0915, 0.0838, 0.0824, 0.082, 0.0817, 0.0816, 0.0814, 0.0813, 0.0812, 0.081, 0.0809, 0.0808, 0.0807, 0.0806, 0.0805, 0.0803, 0.0802, 0.0801, 0.08, 0.0799, 0.0797, 0.0796, 0.0795, 0.0793, 0.0792, 0.0791)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If densityArray(i + 1) = -1 Then
                VaporAmineDensity = "Beyond Fluid Limits"
                Exit Function
            End If
            VaporAmineDensity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (densityArray(i + 1) - densityArray(i)) + densityArray(i)
            i = i + 1
        Loop
    Case Else
        VaporAmineDensity = "only works at 20 psig fool"
    
End Select

End Function

Function VaporAmineViscosity(Temperature As Single, Pressure As Single)
Dim TempArray, cPArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        cPArray = Array(0.0151, 0.0141, 0.0137, 0.0136, 0.0136, 0.0135, 0.0135, 0.0136, 0.0136, 0.0136, 0.0136, 0.0136, 0.0136, 0.0137, 0.0137, 0.0137, 0.0137, 0.0138, 0.0138, 0.0138, 0.0139, 0.0139, 0.0139, 0.014, 0.014, 0.0141, 0.0141)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If cPArray(i + 1) = -1 Then
                VaporAmineViscosity = "Beyond Fluid Limits"
                Exit Function
            End If
            VaporAmineViscosity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (cPArray(i + 1) - cPArray(i)) + cPArray(i)
            i = i + 1
        Loop
    Case Else
        VaporAmineViscosity = "only works at 20 psig fool"
    
End Select

End Function

Function VaporAmineCp(Temperature As Single, Pressure As Single)
Dim TempArray, cPArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        cPArray = Array(0.3638, 0.4202, 0.4514, 0.4574, 0.4594, 0.4604, 0.4609, 0.4613, 0.4618, 0.4621, 0.4623, 0.4624, 0.4626, 0.4627, 0.4628, 0.4629, 0.463, 0.463, 0.4631, 0.4632, 0.4633, 0.4633, 0.4634, 0.4635, 0.4635, 0.4636, 0.4637)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If cPArray(i + 1) = -1 Then
                VaporAmineCp = "Beyond Fluid Limits"
                Exit Function
            End If
            VaporAmineCp = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (cPArray(i + 1) - cPArray(i)) + cPArray(i)
            i = i + 1
        Loop
    Case Else
        VaporAmineCp = "only works at 20 psig fool"
    
End Select

End Function

Function VaporAmineThermalConductivity(Temperature As Single, Pressure As Single)
Dim TempArray, TCArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        TCArray = Array(0.0157, 0.0157, 0.0157, 0.0157, 0.0157, 0.0157, 0.0157, 0.0157, 0.0157, 0.0158, 0.0158, 0.0158, 0.0158, 0.0159, 0.0159, 0.016, 0.016, 0.016, 0.0161, 0.0161, 0.0162, 0.0162, 0.0163, 0.0163, 0.0164, 0.0165, 0.0165)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If TCArray(i + 1) = -1 Then
                VaporAmineThermalConductivity = "Beyond Fluid Limits"
                Exit Function
            End If
            VaporAmineThermalConductivity = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (TCArray(i + 1) - TCArray(i)) + TCArray(i)
            i = i + 1
        Loop
    Case Else
        VaporAmineThermalConductivity = "only works at 20 psig fool"
    
End Select

End Function

Function VaporAmineMoleWeight(Temperature As Single, Pressure As Single)
Dim TempArray, MWArray As Variant

Select Case Pressure
    Case 20
        
        TempArray = Array(256.38, 266.444, 270.593, 271.564, 272.111, 272.563, 272.988, 273.409, 274.163, 274.96, 275.817, 276.743, 277.753, 278.857, 280.071, 281.41, 282.585, 283.864, 285.261, 286.791, 288.473, 290.004, 291.668, 293.481, 295.46, 297.626, 300#)
        
        MWArray = Array(23.93, 20.21, 18.61, 18.33, 18.24, 18.2, 18.17, 18.16, 18.14, 18.13, 18.13, 18.13, 18.13, 18.13, 18.13, 18.14, 18.14, 18.15, 18.16, 18.17, 18.18, 18.19, 18.21, 18.22, 18.24, 18.26, 18.29)

        'interpolation between tabulated temperature values
        i = 0
        Do While Temperature > TempArray(i)
            If MWArray(i + 1) = -1 Then
                VaporAmineMoleWeight = "Beyond Fluid Limits"
                Exit Function
            End If
            VaporAmineMoleWeight = (Temperature - TempArray(i)) / (TempArray(i + 1) - TempArray(i)) * (MWArray(i + 1) - MWArray(i)) + MWArray(i)
            i = i + 1
        Loop
    Case Else
        VaporAmineMoleWeight = "only works at 20 psig fool"
    
End Select

End Function

-------------------------------------------------------------------------------
VBA MACRO PricingFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/PricingFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit
Function FindPrice(STNumber As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

Dim EndDate As Date
EndDate = #1/1/1753#


'Pull cost from material table
SQL = "SELECT [tblPurchasePrice].[Direct Unit Cost]"
'define table
SQL = SQL & " FROM [tblPurchasePrice]"

SQL = SQL & " WHERE [tblPurchasePrice].[Item No_]='" & STNumber & "'"

SQL = SQL & " AND [tblPurchasePrice].[Ending Date]=#" & EndDate & "#"


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPrice = Recordset(0)
    FindPrice = CSng(FindPrice)
    Else: FindPrice = 0
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function DLDPipingPrice(LineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

'Pull cost from material table
SQL = "SELECT [Exp Tank Tower & DLD Piping Pricing Table].[DLD Piping Price]"
'define table
SQL = SQL & " FROM [Exp Tank Tower & DLD Piping Pricing Table]"

SQL = SQL & " WHERE [Exp Tank Tower & DLD Piping Pricing Table].[Nominal Skid Line Size]=" & LineSize


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    DLDPipingPrice = Recordset(0)
    Else: DLDPipingPrice = 0
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function TankTowerPrice(LineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

'Pull cost from material table
SQL = "SELECT [Exp Tank Tower & DLD Piping Pricing Table].[Structure Price]"
'define table
SQL = SQL & " FROM [Exp Tank Tower & DLD Piping Pricing Table]"

SQL = SQL & " WHERE [Exp Tank Tower & DLD Piping Pricing Table].[Nominal Skid Line Size]=" & LineSize


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    TankTowerPrice = Recordset(0)
    Else: TankTowerPrice = 0
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function TankTowerAssemblyPrice(LineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

'Pull cost from material table
SQL = "SELECT [Exp Tank Tower & DLD Piping Pricing Table].[Assembly Price]"
'define table
SQL = SQL & " FROM [Exp Tank Tower & DLD Piping Pricing Table]"

SQL = SQL & " WHERE [Exp Tank Tower & DLD Piping Pricing Table].[Nominal Skid Line Size]=" & LineSize


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    TankTowerAssemblyPrice = Recordset(0)
    Else: TankTowerAssemblyPrice = 0
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function TankTowerPaintPrice(LineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

'Pull cost from material table
SQL = "SELECT [Exp Tank Tower & DLD Piping Pricing Table].[Paint Price]"
'define table
SQL = SQL & " FROM [Exp Tank Tower & DLD Piping Pricing Table]"

SQL = SQL & " WHERE [Exp Tank Tower & DLD Piping Pricing Table].[Nominal Skid Line Size]=" & LineSize


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    TankTowerPaintPrice = Recordset(0)
    Else: TankTowerPaintPrice = 0
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SalesCommission(Profit As Single)

If Profit <= 350000 Then
    SalesCommission = Profit * 0.08
    ElseIf Profit > 350000 And Profit <= 700000 Then
        SalesCommission = 350000 * 0.08 + 0.06 * (Profit - 350000)
    ElseIf Profit > 700000 Then
        SalesCommission = 350000 * 0.08 + 0.06 * 350000 + 0.04 * (Profit - 700000)
End If

 
End Function
Function RepCommission(EquipmentType As String, SalesPrice As Single)
If EquipmentType = "Fired" Then
    If SalesPrice <= 100000 Then
        RepCommission = SalesPrice * 0.07
        ElseIf SalesPrice > 100000 And SalesPrice <= 200000 Then
            RepCommission = 100000 * 0.07 + 0.06 * (SalesPrice - 100000)
        ElseIf SalesPrice > 200000 And SalesPrice <= 500000 Then
            RepCommission = 100000 * 0.07 + 0.06 * 100000 + 0.05 * (SalesPrice - 200000)
        ElseIf SalesPrice > 500000 And SalesPrice <= 1000000 Then
            RepCommission = 100000 * 0.07 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * (SalesPrice - 500000)
        ElseIf SalesPrice > 1000000 And SalesPrice <= 2000000 Then
            RepCommission = 100000 * 0.07 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * 500000 + 0.015 * (SalesPrice - 1000000)
        ElseIf SalesPrice > 2000000 And SalesPrice <= 4000000 Then
            RepCommission = 100000 * 0.07 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * 500000 + 0.015 * 1000000 + 0.01 * (SalesPrice - 2000000)
        ElseIf SalesPrice > 4000000 Then
            RepCommission = 100000 * 0.07 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * 500000 + 0.015 * 1000000 + 0.01 * 2000000 + 0.005 * (SalesPrice - 4000000)
    End If

ElseIf EquipmentType = "Electric" Then
    If SalesPrice <= 50000 Then
        RepCommission = SalesPrice * 0.1
        ElseIf SalesPrice > 50000 And SalesPrice <= 100000 Then
            RepCommission = 50000 * 0.1 + 0.08 * (SalesPrice - 50000)
        
        ElseIf SalesPrice > 100000 And SalesPrice <= 200000 Then
            RepCommission = 50000 * 0.1 + 0.08 * 50000 + 0.06 * (SalesPrice - 100000)
        
        ElseIf SalesPrice > 200000 And SalesPrice <= 500000 Then
            RepCommission = 50000 * 0.1 + 0.08 * 50000 + 0.06 * 100000 + 0.05 * (SalesPrice - 200000)
        
        ElseIf SalesPrice > 500000 And SalesPrice <= 1000000 Then
            RepCommission = 50000 * 0.1 + 0.08 * 50000 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * (SalesPrice - 500000)
        
        ElseIf SalesPrice > 1000000 And SalesPrice <= 2000000 Then
           RepCommission = 50000 * 0.1 + 0.08 * 50000 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * 500000 + 0.02 * (SalesPrice - 1000000)
        
        ElseIf SalesPrice > 2000000 And SalesPrice <= 4000000 Then
           RepCommission = 50000 * 0.1 + 0.08 * 50000 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * 500000 + 0.02 * 1000000 + 0.01 * (SalesPrice - 2000000)
            
        ElseIf SalesPrice > 4000000 Then
           RepCommission = 50000 * 0.1 + 0.08 * 50000 + 0.06 * 100000 + 0.05 * 300000 + 0.03 * 500000 + 0.02 * 1000000 + 0.01 * 2000000 + 0.005 * (SalesPrice - 4000000)
    End If

ElseIf EquipmentType = "Biomass" Then
     If SalesPrice <= 250000 Then
        RepCommission = SalesPrice * 0.05
        ElseIf SalesPrice > 250000 And SalesPrice <= 1000000 Then
            RepCommission = 250000 * 0.05 + 0.03 * (SalesPrice - 250000)
        ElseIf SalesPrice > 1000000 And SalesPrice <= 2000000 Then
            RepCommission = 250000 * 0.05 + 0.03 * 750000 + 0.02 * (SalesPrice - 1000000)
        ElseIf SalesPrice > 2000000 And SalesPrice <= 4000000 Then
            RepCommission = 250000 * 0.05 + 0.03 * 750000 + 0.02 * 1000000 + 0.01 * (SalesPrice - 2000000)
         ElseIf SalesPrice > 2000000 And SalesPrice <= 4000000 Then
            RepCommission = 250000 * 0.05 + 0.03 * 750000 + 0.02 * 1000000 + 0.01 * 2000000 + 0.005 * (SalesPrice - 4000000)
    End If
ElseIf EquipmentType = "Spares" Or EquipmentType = "Service" Then
     If SalesPrice <= 25000 Then
        RepCommission = SalesPrice * 0.1
     ElseIf SalesPrice > 25000 And SalesPrice <= 50000 Then
            RepCommission = 25000 * 0.1 + 0.08 * (SalesPrice - 25000)
     ElseIf SalesPrice > 50000 And SalesPrice <= 100000 Then
            RepCommission = 25000 * 0.1 + 0.08 * 25000 + 0.06 * (SalesPrice - 50000)
     ElseIf SalesPrice > 100000 And SalesPrice <= 250000 Then
            RepCommission = 25000 * 0.1 + 0.08 * 25000 + 0.06 * 50000 + 0.05 * (SalesPrice - 100000)
     ElseIf SalesPrice > 250000 Then
            RepCommission = 25000 * 0.1 + 0.08 * 25000 + 0.06 * 50000 + 0.05 * 150000 + 0.005 * (SalesPrice - 250000)
     End If
ElseIf EquipmentType = "Automation" Then
     If SalesPrice <= 50000 Then
        RepCommission = SalesPrice * 0.08
     ElseIf SalesPrice > 50000 And SalesPrice <= 100000 Then
        RepCommission = 50000 * 0.08 + 0.07 * (SalesPrice - 50000)
     ElseIf SalesPrice > 100000 And SalesPrice <= 150000 Then
        RepCommission = 50000 * 0.08 + 0.07 * 50000 + 0.06 * (SalesPrice - 100000)
     ElseIf SalesPrice > 150000 And SalesPrice <= 200000 Then
        RepCommission = 50000 * 0.08 + 0.07 * 50000 + 0.06 * 50000 + 0.05 * (SalesPrice - 150000)
     ElseIf SalesPrice > 200000 And SalesPrice <= 500000 Then
        RepCommission = 50000 * 0.08 + 0.07 * 50000 + 0.06 * 50000 + 0.05 * 50000 + 0.03 * (SalesPrice - 200000)
     ElseIf SalesPrice > 500000 And SalesPrice <= 1000000 Then
        RepCommission = 50000 * 0.08 + 0.07 * 50000 + 0.06 * 50000 + 0.05 * 50000 + 0.03 * 300000 + 0.02 * (SalesPrice - 500000)
     ElseIf SalesPrice > 1000000 And SalesPrice <= 3000000 Then
        RepCommission = 50000 * 0.08 + 0.07 * 50000 + 0.06 * 50000 + 0.05 * 50000 + 0.03 * 300000 + 0.02 * 500000 + 0.01 * (SalesPrice - 1000000)
     ElseIf SalesPrice > 3000000 Then
        RepCommission = 50000 * 0.08 + 0.07 * 50000 + 0.06 * 50000 + 0.05 * 50000 + 0.03 * 300000 + 0.02 * 500000 + 0.01 * 2000000 + 0.005 * (SalesPrice - 3000000)
    End If
End If
End Function



Function MainSSOVSize(FlowRate As Single)

If FlowRate <= 2000 Then
        MainSSOVSize = 1
    ElseIf FlowRate > 2000 And FlowRate <= 5500 Then
        MainSSOVSize = 1.5
    ElseIf FlowRate > 5500 And FlowRate <= 9000 Then
        MainSSOVSize = 2
    ElseIf FlowRate > 9000 And FlowRate <= 35000 Then
        MainSSOVSize = 2.5
    ElseIf FlowRate > 35000 And FlowRate <= 50000 Then
        MainSSOVSize = 3
    ElseIf FlowRate > 50000 And FlowRate <= 65000 Then
        MainSSOVSize = 4
    ElseIf FlowRate > 65000 And FlowRate <= 125000 Then
        MainSSOVSize = 6
End If


End Function
Function VentValveSize(MainSSOVSize As Single)

Select Case MainSSOVSize
    Case Is <= 1.25
        VentValveSize = 0.5
    Case 1.5
        VentValveSize = 0.75
    Case 2
        VentValveSize = 1
    Case 2.5
        VentValveSize = 1.5
    Case 3
        VentValveSize = 1.5
    Case 4
        VentValveSize = 2
    Case 6
        VentValveSize = 2.5
End Select


End Function
Function BurnerPicLookup(Series As String, ControlMethod As String, PackFan As String)

If Left(Series, 5) = "Ratio" Then
    BurnerPicLookup = "SPPRatio"
ElseIf ControlMethod = "Single Point Positioning" Then
    If PackFan = "Yes" Then
        BurnerPicLookup = "SPPPackage"
    Else: BurnerPicLookup = "SPPRemote"
    End If
Else: BurnerPicLookup = "PPP"
End If

End Function
Function HeaterPicLookup(HeaterConfig As String, FlowConfig As String)
Dim Part1 As String
Dim Part2 As String

If HeaterConfig = "Horizontal" Then
    Part1 = "Horizontal"
ElseIf HeaterConfig = "Vertical Downfired" Then
    Part1 = "VerticalDownfired"
Else: Part1 = "VerticalUpfired"
End If

Part2 = FlowConfig

HeaterPicLookup = Part1 & Part2

End Function
Function FTPicLookup(PilotPressure As Single, RatedPressure As Single)

If PilotPressure > RatedPressure Then
    FTPicLookup = "Upstream"
Else: FTPicLookup = "Downstream"
End If

End Function

Function SSOVPrice(Size As String, AreaClassification As String, Actuation As String, Connection As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlAreaClassification, sqlSize, sqlConnection, sqlActuation  As String

sqlSize = "'" & Size & "'"
sqlAreaClassification = "'" & AreaClassification & "'"
sqlConnection = "'" & Connection & "'"
sqlActuation = "'" & Actuation & "'"

'Pull cost from material table
SQL = "SELECT [SSOVs].Price"
'define table
SQL = SQL & " FROM [SSOVs]"

SQL = SQL & " WHERE [SSOVs].[ValveSize]= " & sqlSize

SQL = SQL & " AND [SSOVs].[AreaClassification]=" & sqlAreaClassification

SQL = SQL & " AND [SSOVs].[Pneumatic/Electric]=" & sqlActuation

SQL = SQL & " AND [SSOVs].[ConnectionType]=" & sqlConnection

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SSOVPrice = Recordset(0)
    Else: SSOVPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassValve(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].ValveST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassValve = Recordset(0)
    Else: SystemBypassValve = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassActuator(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].ActuatorST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassActuator = Recordset(0)
    Else: SystemBypassActuator = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassYoke(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].YokeST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassYoke = Recordset(0)
    Else: SystemBypassYoke = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassFlanges(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].FlangesST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassFlanges = Recordset(0)
    Else: SystemBypassFlanges = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassBlind(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].BlindST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassBlind = Recordset(0)
    Else: SystemBypassBlind = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassGaskets(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].GasketsST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassGaskets = Recordset(0)
    Else: SystemBypassGaskets = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassPositioner(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].PositionerST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassPositioner = Recordset(0)
    Else: SystemBypassPositioner = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SystemBypassBracket(Size As String, AreaClassification As String, MinTemp As Single)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

'Pull cost from material table
SQL = "SELECT [tblSystemBypass].BracketST"
'define table
SQL = SQL & " FROM [tblSystemBypass]"

SQL = SQL & " WHERE [tblSystemBypass].[Size]=" & "'" & Size & "'"

SQL = SQL & " AND [tblSystemBypass].[AreaClass]=" & "'" & AreaClassification & "'"

SQL = SQL & " AND [tblSystemBypass].[MinAmbTemp]<=" & MinTemp

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SystemBypassBracket = Recordset(0)
    Else: SystemBypassBracket = "N/A"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function MainFuelTrainDesignPressure(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        MainFuelTrainDesignPressure = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

'Pull cost from material table
SQL = "SELECT [Main].[Max Inlet Pressure]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
        MainFuelTrainDesignPressure = Recordset(0)
     Else: MainFuelTrainDesignPressure = "Not Found"
End If


If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function MainRegulatorRegistration(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        MainRegulatorRegistration = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

'Pull cost from material table
SQL = "SELECT [Main].[Regulator Registration]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
        MainRegulatorRegistration = Recordset(0)
     Else: MainRegulatorRegistration = "Not Found"
End If


If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function FindMainRegulator(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 15 Then
        FuelPressure = 15
    ElseIf FuelPressure < 15 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainRegulator = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Regulator]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
        FindMainRegulator = Recordset(0)
     Else: FindMainRegulator = "Not Found"
End If


If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainSSOVs(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainSSOVs = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main SSOVs]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainSSOVs = Recordset(0)
    Else: FindMainSSOVs = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainVentValve(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainVentValve = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Vent Valve]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainVentValve = Recordset(0)
    Else: FindMainVentValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindUpstreamPressureGauge(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindUpstreamPressureGauge = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Upstream Pressure Gauge]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindUpstreamPressureGauge = Recordset(0)
    Else: FindUpstreamPressureGauge = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainPressureGauges(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainPressureGauges = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Pressure Gauges]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainPressureGauges = Recordset(0)
    Else: FindMainPressureGauges = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function FindMainPressureSwitches(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainPressureSwitches = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Pressure Switches]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainPressureSwitches = Recordset(0)
    Else: FindMainPressureSwitches = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindUpstreamGaugeValve(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindUpstreamGaugeValve = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Upstream Gauge Valve]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindUpstreamGaugeValve = Recordset(0)
    Else: FindUpstreamGaugeValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainGaugeValves(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainGaugeValves = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Gauge Valves]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainGaugeValves = Recordset(0)
    Else: FindMainGaugeValves = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainInletIsoValve(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer
Dim InFuelPressure As Single

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainInletIsoValve = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Inlet Iso Valve]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainInletIsoValve = Recordset(0)
    Else: FindMainInletIsoValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainOutletIsoValve(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainOutletIsoValve = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Outlet Iso Valve]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainOutletIsoValve = Recordset(0)
    Else: FindMainOutletIsoValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindFuelTrainStrainer(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindFuelTrainStrainer = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Strainer]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindFuelTrainStrainer = Recordset(0)
    Else: FindFuelTrainStrainer = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainSnubbers(FuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, RegulatorPressure As Single, FuelFlowRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

If FuelPressure >= 20 Then
        FuelPressure = 20
    ElseIf FuelPressure < 20 And FuelPressure >= 10 Then
        FuelPressure = 10
    ElseIf FuelPressure < 10 Then
        FindMainSnubbers = "HIGHER INLET FUEL PRESSURE REQUIRED"
        Exit Function
End If

If RegulatorPressure < 16 Then RegulatorPressure = 20

'Pull cost from material table
SQL = "SELECT [Main].[Main Snubbers]"
'define table
SQL = SQL & "FROM [Main]"
'constraint #1
SQL = SQL & "WHERE [Main].[Inlet Fuel Pressure]<=" & FuelPressure
'constraint #2
SQL = SQL & "AND [Main].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Main].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Main].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Main].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Main].[Min Reg Pressure]<=" & RegulatorPressure
'constraint #7
SQL = SQL & "AND [Main].[Max Reg Pressure]>" & RegulatorPressure
'constraint #8
SQL = SQL & "AND [Main].[Min Capacity]<=" & FuelFlowRate
'constraint #9
SQL = SQL & "AND [Main].[Max Capacity]>=" & FuelFlowRate
'constraint #10
SQL = SQL & "AND [Main].[Max Inlet Pressure]>=" & FuelPressure


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainSnubbers = Recordset(0)
    Else: FindMainSnubbers = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotRegulator(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot Regulator]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotRegulator = Recordset(0)
    Else: FindPilotRegulator = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function PilotSSOVSize(AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[SSOV Size]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PilotSSOVSize = Recordset(0)
    Else: PilotSSOVSize = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotSSOVs(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot SSOVs]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotSSOVs = Recordset(0)
    Else: FindPilotSSOVs = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotSSOVManufacturer(AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[SSOV Manufacturer]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotSSOVManufacturer = Recordset(0)
    Else: FindPilotSSOVManufacturer = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotVent(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot Vent Valve]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotVent = Recordset(0)
    Else: FindPilotVent = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotPressureGauge(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot Pressure Gauge]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotPressureGauge = Recordset(0)
    Else: FindPilotPressureGauge = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotGaugeValve(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot Gauge Valve]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotGaugeValve = Recordset(0)
    Else: FindPilotGaugeValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotInletIsoValve(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot Inlet Iso Valve]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotInletIsoValve = Recordset(0)
    Else: FindPilotInletIsoValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotOutletIsoValve(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot Outlet Iso Valve]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotOutletIsoValve = Recordset(0)
    Else: FindPilotOutletIsoValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotNeedleValve(PilotFuelPressure As Single, AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Pilot Needle Valve]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constraint #10
SQL = SQL & "AND [Pilot].[Min Inlet Pressure]<=" & PilotFuelPressure
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotNeedleValve = Recordset(0)
    Else: FindPilotNeedleValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function MinPilotPressure(AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Min Inlet Pressure]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    MinPilotPressure = Recordset(0)
    Else: MinPilotPressure = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function MaxPilotPressure(AreaClass As String, Connect As String, MinTemp As Single, ControlVoltage As String, PilotRegulatorPressure As Single, PilotFuelFlowRate As Single, POC As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant
Dim temp As Variant
Dim i As Integer

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [Pilot].[Design Inlet Pressure]"
'define table
SQL = SQL & "FROM [Pilot]"
'constraint #2
SQL = SQL & "WHERE [Pilot].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [Pilot].[Connection]=" & "'" & Connect & "'"
'constraint #4
SQL = SQL & "AND [Pilot].[Temp Rating]=" & MinTemp
'constraint #5
SQL = SQL & "AND [Pilot].[Control Voltage]=" & "'" & ControlVoltage & "'"
'constraint #6
SQL = SQL & "AND [Pilot].[Min Reg Pressure]<=" & PilotRegulatorPressure
'constraint #7
SQL = SQL & "AND [Pilot].[Max Reg Pressure]>" & PilotRegulatorPressure
'constraint #8
SQL = SQL & "AND [Pilot].[Min Capacity]<=" & PilotFuelFlowRate
'constraint #9
SQL = SQL & "AND [Pilot].[Max Capacity]>=" & PilotFuelFlowRate
'constrain #11
SQL = SQL & "AND [Pilot].[POC]=" & "'" & POC & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    MaxPilotPressure = Recordset(0)
    Else: MaxPilotPressure = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindStackPrice(diameter As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull pricing
SQL = "SELECT [Stack Pricing].[Price]"
'define table
SQL = SQL & "FROM [Stack Pricing]"
'constraint #1
SQL = SQL & "WHERE [Stack Pricing].[Diameter]=" & diameter

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindStackPrice = Recordset(0)
    Else: FindStackPrice = "Not Found"
End If

End Function

Function BurnerPilotCapacity(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Pilot Capacity, MMBTU/hr]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerPilotCapacity = Recordset(0)
    Else: BurnerPilotCapacity = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpModel(Fluid As String, SystemFlowRate As Single, Manufacturer As String, PumpIsoValveType As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Model]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #5
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpModel = Recordset(0)
    Else: FindPumpModel = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindExpansionTank(TankCapacity As Single, StandardAlt As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank].[Tank Subassembly]"
'define table
SQL = SQL & "FROM [Expansion Tank]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank].[Size]=" & TankCapacity
'constraint #2
SQL = SQL & "AND [Expansion Tank].[Standard]=" & "'" & StandardAlt & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindExpansionTank = Recordset(0)
    Else: FindExpansionTank = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindExpTankLevelSwitch(TankCapacity As Single, StandardAlt As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank].[Device Package]"
'define table
SQL = SQL & "FROM [Expansion Tank]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank].[Size]=" & TankCapacity
'constraint #2
SQL = SQL & "AND [Expansion Tank].[Standard]=" & "'" & StandardAlt & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindExpTankLevelSwitch = Recordset(0)
    Else: FindExpTankLevelSwitch = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindLevelGauge(TankCapacity As Single, StandardAlt As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank].[Level Gauge]"
'define table
SQL = SQL & "FROM [Expansion Tank]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank].[Size]=" & TankCapacity
'constraint #2
SQL = SQL & "AND [Expansion Tank].[Standard]=" & "'" & StandardAlt & "'"


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindLevelGauge = Recordset(0)
    Else: FindLevelGauge = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindLevelGaugeIsoValve(TankCapacity As Single, StandardAlt As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank].[Isolation Valves]"
'define table
SQL = SQL & "FROM [Expansion Tank]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank].[Size]=" & TankCapacity
'constraint #2
SQL = SQL & "AND [Expansion Tank].[Standard]=" & "'" & StandardAlt & "'"


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindLevelGaugeIsoValve = Recordset(0)
    Else: FindLevelGaugeIsoValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindExpTankGaugeVentDrainValve(TankCapacity As Single, StandardAlt As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Expansion Tank].[Vent/Drain Valves]"
'define table
SQL = SQL & "FROM [Expansion Tank]"
'constraint #1
SQL = SQL & "WHERE [Expansion Tank].[Size]=" & TankCapacity
'constraint #2
SQL = SQL & "AND [Expansion Tank].[Standard]=" & "'" & StandardAlt & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindExpTankGaugeVentDrainValve = Recordset(0)
    Else: FindExpTankGaugeVentDrainValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function


Function FindPumpInletIsoValve(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Inlet Iso Valve]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpInletIsoValve = Recordset(0)
    Else: FindPumpInletIsoValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpOutletIsoValve(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Outlet Iso Valve]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpOutletIsoValve = Recordset(0)
    Else: FindPumpOutletIsoValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpStrainer(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Strainer]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpStrainer = Recordset(0)
    Else: FindPumpStrainer = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpStrainerDrainValve(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Strainer Drain Valve]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpStrainerDrainValve = Recordset(0)
    Else: FindPumpStrainerDrainValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpCheckValve(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Check Valve]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpCheckValve = Recordset(0)
    Else: FindPumpCheckValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindExpansionBellows(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Expansion Bellows]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindExpansionBellows = Recordset(0)
    Else: FindExpansionBellows = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindSuctionPressureGauge(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Suction Pressure Gauge]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindSuctionPressureGauge = Recordset(0)
    Else: FindSuctionPressureGauge = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindDischargePressureGauge(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Discharge Pressure Gauge]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindDischargePressureGauge = Recordset(0)
    Else: FindDischargePressureGauge = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindSiphons(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Siphons]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindSiphons = Recordset(0)
    Else: FindSiphons = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpSkidDrainValves(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Drain Valves]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpSkidDrainValves = Recordset(0)
    Else: FindPumpSkidDrainValves = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpSkidGaugeValves(SystemFlowRate As Single, PumpIsoValveType As String, PumpMan As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Gauge Valves]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #2
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate
'constraint #3
SQL = SQL & "AND [Pump Skid].[Type of Iso Valve]=" & "'" & PumpIsoValveType & "'"
'constraint #4
SQL = SQL & "AND [Pump Skid].[Pump Manufacturer]=" & "'" & PumpMan & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpSkidGaugeValves = Recordset(0)
    Else: FindPumpSkidGaugeValves = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindHeaterPSV(HeaterModel As String, AreaClass As String, MinTemp As Single, StackTW As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [HC Heaters Instrumentation].[PSV]"
'define table
SQL = SQL & "FROM [HC Heaters Instrumentation]"
'constraint #1
SQL = SQL & "WHERE [HC Heaters Instrumentation].[Temp Rating]=" & MinTemp
'constraint #2
SQL = SQL & "AND [HC Heaters Instrumentation].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [HC Heaters Instrumentation].[HC Heater]=" & "'" & HeaterModel & "'"
'constraint #4
SQL = SQL & "AND [HC Heaters Instrumentation].[Stack Thermowell Used]=" & "'" & StackTW & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindHeaterPSV = Recordset(0)
    Else: FindHeaterPSV = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindCoilTCs(HeaterModel As String, AreaClass As String, MinTemp As Single, StackTW As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [HC Heaters Instrumentation].[Coil Thermocouples]"
'define table
SQL = SQL & "FROM [HC Heaters Instrumentation]"
'constraint #1
SQL = SQL & "WHERE [HC Heaters Instrumentation].[Temp Rating]=" & MinTemp
'constraint #2
SQL = SQL & "AND [HC Heaters Instrumentation].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [HC Heaters Instrumentation].[HC Heater]=" & "'" & HeaterModel & "'"
'constraint #4
SQL = SQL & "AND [HC Heaters Instrumentation].[Stack Thermowell Used]=" & "'" & StackTW & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindCoilTCs = Recordset(0)
    Else: FindCoilTCs = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindCoilTWs(HeaterModel As String, AreaClass As String, MinTemp As Single, StackTW As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [HC Heaters Instrumentation].[Coil Thermowells]"
'define table
SQL = SQL & "FROM [HC Heaters Instrumentation]"
'constraint #1
SQL = SQL & "WHERE [HC Heaters Instrumentation].[Temp Rating]=" & MinTemp
'constraint #2
SQL = SQL & "AND [HC Heaters Instrumentation].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [HC Heaters Instrumentation].[HC Heater]=" & "'" & HeaterModel & "'"
'constraint #4
SQL = SQL & "AND [HC Heaters Instrumentation].[Stack Thermowell Used]=" & "'" & StackTW & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindCoilTWs = Recordset(0)
    Else: FindCoilTWs = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindFlowSwitchAssembly(HeaterModel As String, AreaClass As String, MinTemp As Single, StackTW As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If MinTemp >= -20 Then
        MinTemp = -20
    ElseIf MinTemp < -20 And MinTemp >= -40 Then
        MinTemp = -40
End If

'Pull cost from material table
SQL = "SELECT [HC Heaters Instrumentation].[DPI, 3VM, Flow Switches]"
'define table
SQL = SQL & "FROM [HC Heaters Instrumentation]"
'constraint #1
SQL = SQL & "WHERE [HC Heaters Instrumentation].[Temp Rating]=" & MinTemp
'constraint #2
SQL = SQL & "AND [HC Heaters Instrumentation].[Area Classification]=" & "'" & AreaClass & "'"
'constraint #3
SQL = SQL & "AND [HC Heaters Instrumentation].[HC Heater]=" & "'" & HeaterModel & "'"
'constraint #4
SQL = SQL & "AND [HC Heaters Instrumentation].[Stack Thermowell Used]=" & "'" & StackTW & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindFlowSwitchAssembly = Recordset(0)
    Else: FindFlowSwitchAssembly = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindHCCoil(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Left(HeaterModel, 3) = "HC2" Then

    'Pull cost from material table
    SQL = "SELECT [HC2 Heater Pricing Table].[Coil ST Number]"
    'define table
    SQL = SQL & "FROM [HC2 Heater Pricing Table]"
    'constraint #1
    SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Else
    'Pull cost from material table
    SQL = "SELECT [HC1 Heater Pricing Table].[Coil ST Number]"
    'define table
    SQL = SQL & "FROM [HC1 Heater Pricing Table]"
    'constraint #1
    SQL = SQL & "WHERE [HC1 Heater Pricing Table].[Model No]=" & "'" & HeaterModel & "'"
    
End If


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindHCCoil = Recordset(0)
    Else: FindHCCoil = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilAssemblyPrice(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Coil Assembly]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilAssemblyPrice = Recordset(0)
    Else: CoilAssemblyPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function HeaterShellPrice(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Shell]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    HeaterShellPrice = Recordset(0)
    Else: HeaterShellPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function HeaterLidPrice(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Lid]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    HeaterLidPrice = Recordset(0)
    Else: HeaterLidPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function HeaterSubAssembly(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Heater Sub Assembly]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    HeaterSubAssembly = Recordset(0)
    Else: HeaterSubAssembly = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CoilMaterialAIQualityPrice(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Coil Material and AI]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CoilMaterialAIQualityPrice = Recordset(0)
    Else: CoilMaterialAIQualityPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function LidMaterialPrice(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Lid Material]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    LidMaterialPrice = Recordset(0)
    Else: LidMaterialPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ShellMaterialPrice(HeaterModel As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Shell Material]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ShellMaterialPrice = Recordset(0)
    Else: ShellMaterialPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SkidLaborPrice(HeaterModel As String, Orientation As String, SkidSaddle As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Left(Orientation, 8) = "Vertical" Then
'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Vertical Skid Frame Labor]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"
ElseIf SkidSaddle = "Skid Frame" Then
'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Horizontal Skid Frame Labor]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"
Else:
'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Horizontal Saddle Labor]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"
End If

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SkidLaborPrice = Recordset(0)
    Else: SkidLaborPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function SkidMaterialPrice(HeaterModel As String, Orientation As String, SkidSaddle As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Left(Orientation, 8) = "Vertical" Then
'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Vertical Skid Frame Material]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"
ElseIf SkidSaddle = "Skid Frame" Then
'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Horizontal Skid Frame Material]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"
Else:
'Pull cost from material table
SQL = "SELECT [HC2 Heater Pricing Table].[Horizontal Saddle Material]"
'define table
SQL = SQL & "FROM [HC2 Heater Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [HC2 Heater Pricing Table].[Model Number]=" & "'" & HeaterModel & "'"
End If

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SkidMaterialPrice = Recordset(0)
    Else: SkidMaterialPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function StackBaseHeight(StackDia As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Stack Pricing Table].[Base Height (ft)]"
'define table
SQL = SQL & "FROM [Stack Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Stack Pricing Table].[Diameter (in)]=" & StackDia

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    StackBaseHeight = Recordset(0)
    Else: StackBaseHeight = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function StackBasePrice(StackDia As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Stack Pricing Table].[Price]"
'define table
SQL = SQL & "FROM [Stack Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Stack Pricing Table].[Diameter (in)]=" & StackDia

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    StackBasePrice = Recordset(0)
    Else: StackBasePrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function StackExtraPrice(StackDia As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Stack Pricing Table].[Price per extra 5 ft of height]"
'define table
SQL = SQL & "FROM [Stack Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Stack Pricing Table].[Diameter (in)]=" & StackDia

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    StackExtraPrice = Recordset(0)
    Else: StackExtraPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function PumpMotorRPM(Manufacturer As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Motor RPM]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpMotorRPM = Recordset(0)
    Else: PumpMotorRPM = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function PumpMotorHP(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Motor HP]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpMotorHP = Recordset(0)
    Else: PumpMotorHP = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function PumpMotorFrameSize(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Motor Frame Size]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpMotorFrameSize = Recordset(0)
    Else: PumpMotorFrameSize = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindBarePump(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump ST Number]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindBarePump = Recordset(0)
    Else: FindBarePump = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPumpMotor(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Motor ST Number]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpMotor = Recordset(0)
    Else: FindPumpMotor = "Not Found"
End If

End Function
Function FindPumpCoupling(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Coupling]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpCoupling = Recordset(0)
    Else: FindPumpCoupling = "Not Found"
End If

End Function
Function FindPumpCouplingGuard(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Coupling Guard]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpCouplingGuard = Recordset(0)
    Else: FindPumpCouplingGuard = "Not Found"
End If

End Function
Function FindPumpBasePlate(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Base Plate]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpBasePlate = Recordset(0)
    Else: FindPumpBasePlate = "Not Found"
End If

End Function
Function FindPumpAssembly(Manufacturer As String, SystemFlowRate As Single, Fluid As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Mid(Fluid, 2, 6) <> "Glycol" And Left(Fluid, 3) <> "TEG" Then
    Fluid = "Thermal Oil"
Else: Fluid = "Water/Glycol"
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid].[Pump Assembly]"
'define table
SQL = SQL & "FROM [Pump Skid]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid].[Pump Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Pump Skid].[Fluid Type]=" & "'" & Fluid & "'"
'constraint #3
SQL = SQL & "AND [Pump Skid].[Max Design Flow (GPM)]>=" & SystemFlowRate
'constraint #4
SQL = SQL & "AND [Pump Skid].[Min Design Flow (GPM)]<" & SystemFlowRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPumpAssembly = Recordset(0)
    Else: FindPumpAssembly = "Not Found"
End If

End Function
Function PumpSkidFabPrice(LineSize As Single, PumpQuantity As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid Pricing Table].[Skid Frame Pricing]"
'define table
SQL = SQL & "FROM [Pump Skid Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid Pricing Table].[Nominal Skid Line Size]=" & LineSize
'constraint #2
SQL = SQL & "AND [Pump Skid Pricing Table].[Pump Quantity]=" & PumpQuantity

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpSkidFabPrice = Recordset(0)
    Else: PumpSkidFabPrice = "Not Found"
End If

End Function
Function PumpSkidPipePrice(LineSize As Single, PumpQuantity As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid Pricing Table].[Inlet/Outlet Piping Price]"
'define table
SQL = SQL & "FROM [Pump Skid Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid Pricing Table].[Nominal Skid Line Size]=" & LineSize
'constraint #2
SQL = SQL & "AND [Pump Skid Pricing Table].[Pump Quantity]=" & PumpQuantity

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpSkidPipePrice = Recordset(0)
    Else: PumpSkidPipePrice = "Not Found"
End If

End Function
Function PumpSkidAssemblyPrice(LineSize As Single, PumpQuantity As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid Pricing Table].[Assembly Price]"
'define table
SQL = SQL & "FROM [Pump Skid Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid Pricing Table].[Nominal Skid Line Size]=" & LineSize
'constraint #2
SQL = SQL & "AND [Pump Skid Pricing Table].[Pump Quantity]=" & PumpQuantity

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpSkidAssemblyPrice = Recordset(0)
    Else: PumpSkidAssemblyPrice = "Not Found"
End If

End Function
Function PumpSkidPaintPrice(LineSize As Single, PumpQuantity As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Pump Skid Pricing Table].[Paint Price]"
'define table
SQL = SQL & "FROM [Pump Skid Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid Pricing Table].[Nominal Skid Line Size]=" & LineSize
'constraint #2
SQL = SQL & "AND [Pump Skid Pricing Table].[Pump Quantity]=" & PumpQuantity

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpSkidPaintPrice = Recordset(0)
    Else: PumpSkidPaintPrice = "Not Found"
End If

End Function
Function SSFPipePrice(LineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [SSF Piping Pricing Table].[Piping Price]"
'define table
SQL = SQL & "FROM [SSF Piping Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [SSF Piping Pricing Table].[Nominal Skid Line Size]=" & LineSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SSFPipePrice = Recordset(0)
    Else: SSFPipePrice = "Not Found"
End If

End Function
Function SSFPipeAssemblyPrice(LineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [SSF Piping Pricing Table].[Assembly Price]"
'define table
SQL = SQL & "FROM [SSF Piping Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [SSF Piping Pricing Table].[Nominal Skid Line Size]=" & LineSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SSFPipeAssemblyPrice = Recordset(0)
    Else: SSFPipeAssemblyPrice = "Not Found"
End If

End Function
Function SSFPipePaintPrice(LineSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [SSF Piping Pricing Table].[Paint Price]"
'define table
SQL = SQL & "FROM [SSF Piping Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [SSF Piping Pricing Table].[Nominal Skid Line Size]=" & LineSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    SSFPipePaintPrice = Recordset(0)
    Else: SSFPipePaintPrice = "Not Found"
End If

End Function
Function PumpHeaderPipePrice(LineSize As Single, PumpQuantity As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If PumpQuantity <= 1 Then
    PumpHeaderPipePrice = 0
    Exit Function
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid Header Pricing Table].[Piping Price]"
'define table
SQL = SQL & "FROM [Pump Skid Header Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid Header Pricing Table].[Nominal Skid Line Size]=" & LineSize
'constraint #2
SQL = SQL & "AND [Pump Skid Header Pricing Table].[Pump Quantity]=" & PumpQuantity

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpHeaderPipePrice = Recordset(0)
    Else: PumpHeaderPipePrice = "Not Found"
End If

End Function
Function PumpHeaderPaintPrice(LineSize As Single, PumpQuantity As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If PumpQuantity <= 1 Then
    PumpHeaderPaintPrice = 0
    Exit Function
End If

'Pull cost from material table
SQL = "SELECT [Pump Skid Header Pricing Table].[Paint Price]"
'define table
SQL = SQL & "FROM [Pump Skid Header Pricing Table]"
'constraint #1
SQL = SQL & "WHERE [Pump Skid Header Pricing Table].[Nominal Skid Line Size]=" & LineSize
'constraint #2
SQL = SQL & "AND [Pump Skid Header Pricing Table].[Pump Quantity]=" & PumpQuantity

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PumpHeaderPaintPrice = Recordset(0)
    Else: PumpHeaderPaintPrice = "Not Found"
End If

End Function
Function BurnerMaxFuelInletPressure(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Main natural gas inlet pressure, inWC]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerMaxFuelInletPressure = Recordset(0)
    Else: BurnerMaxFuelInletPressure = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerAirPressureReq(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Combustion air pressure at inlet, inWC]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerAirPressureReq = Recordset(0)
    Else: BurnerAirPressureReq = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerPilotInletPressure(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Pilot natural gas inlet pressure, inWC]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerPilotInletPressure = Recordset(0)
    Else: BurnerPilotInletPressure = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerTurndown(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[TurnDown Ratio]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerTurndown = Recordset(0)
    Else: BurnerTurndown = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerMaxCapacity(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Max Burner Capacity, MMBTU/hr]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerMaxCapacity = Recordset(0)
    Else: BurnerMaxCapacity = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerSTNumber(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[ST Number]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerSTNumber = Recordset(0)
    Else: BurnerSTNumber = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerPackagedFan(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Package Fan]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerPackagedFan = Recordset(0)
    Else: BurnerPackagedFan = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerPackagedFuelValve(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Fuel Control Valve]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerPackagedFuelValve = Recordset(0)
    Else: BurnerPackagedFuelValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerPackagedAirValve(Manufacturer As String, Model As String, Series As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Air Control Valve]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Burner Model]=" & "'" & Model & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Burner Series]=" & "'" & Series & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerPackagedAirValve = Recordset(0)
    Else: BurnerPackagedAirValve = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerSeries(Manufacturer As String, NOx As Single, CO As Single, ControlMethod As String, FiringRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Burner Series]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Manufacturer]=" & "'" & Manufacturer & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Control Method]=" & "'" & ControlMethod & "'"
'constraint #3
SQL = SQL & "AND [Burner Lookup].[Nox, ppm]<=" & NOx
'constraint #4
SQL = SQL & "AND [Burner Lookup].[CO, ppm]<=" & CO
'constraint #4
SQL = SQL & "AND [Burner Lookup].[Max Burner Capacity, MMBtu/hr]>=" & FiringRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerSeries = Recordset(0)
    Else: BurnerSeries = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function BurnerModel(Series As String, DesFiringRate As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Burner Lookup].[Burner Model]"
'define table
SQL = SQL & "FROM [Burner Lookup]"
'constraint #1
SQL = SQL & "WHERE [Burner Lookup].[Burner Series]=" & "'" & Series & "'"
'constraint #2
SQL = SQL & "AND [Burner Lookup].[Max Burner Capacity, MMBTU/hr]>=" & DesFiringRate

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerModel = Recordset(0)
    Else: BurnerModel = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveZeroPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[0% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveZeroPerCv = Recordset(0)
    Else: ControlValveZeroPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveTenPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[10% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveTenPerCv = Recordset(0)
    Else: ControlValveTenPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveTwentyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[20% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveTwentyPerCv = Recordset(0)
    Else: ControlValveTwentyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveThirtyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[30% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveThirtyPerCv = Recordset(0)
    Else: ControlValveThirtyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveFortyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[40% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveFortyPerCv = Recordset(0)
    Else: ControlValveFortyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveFiftyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[50% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveFiftyPerCv = Recordset(0)
    Else: ControlValveFiftyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveSixtyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[60% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveSixtyPerCv = Recordset(0)
    Else: ControlValveSixtyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveSeventyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[70% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveSeventyPerCv = Recordset(0)
    Else: ControlValveSeventyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveEightyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[80% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveEightyPerCv = Recordset(0)
    Else: ControlValveEightyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveNinetyPerCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[90% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveNinetyPerCv = Recordset(0)
    Else: ControlValveNinetyPerCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveFullOpenCv(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[100% Open]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveFullOpenCv = Recordset(0)
    Else: ControlValveFullOpenCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveTorque(ControlValveSize As Single, VBallAngle As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[Torque Required (in-lbs)]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[V-Port Angle (degree)]=" & VBallAngle

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ControlValveTorque = Recordset(0)
    Else: ControlValveTorque = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ControlValveSize(Cv As Single)
Dim valveSize As Variant
Dim i As Single
Dim x As Single

i = 0
valveSize = Array(1, 1.5, 2, 2.5, 3, 4)

Do While i <= 5
    x = valveSize(i)
    If FindVBallAngle(x, Cv) = "Not Found" Then
        i = i + 1
    Else: ControlValveSize = x
    Exit Function
    End If
Loop
End Function
Function FindVBallAngle(ControlValveSize As Single, Cv As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Flo Tite Cvs].[V-Port Angle (degree)]"
'define table
SQL = SQL & "FROM [Flo Tite Cvs]"
'constraint #1
SQL = SQL & "WHERE [Flo Tite Cvs].[Valve Size (inch)]=" & ControlValveSize
'constraint #2
SQL = SQL & "AND [Flo Tite Cvs].[80% Open]>=" & Cv

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindVBallAngle = Recordset(0)
    Else: FindVBallAngle = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindControlValveSize(SSOVSize As Single)
Dim i As Single
Dim Size As Variant

i = 1
Size = Array(0.5, 0.75, 1, 1.5, 2, 2.5, 3, 4, 6, 8)

Do While i < 10
    If Size(i) = SSOVSize Then
        FindControlValveSize = Size(i - 1)
        Exit Function
    End If
    i = i + 1
Loop

End Function
Function FindFTOutletIsoValveCv(valveSize As Single, Connect As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Sharpe FT Iso Valve CVs].[Valve Cv]"
'define table
SQL = SQL & "FROM [Sharpe FT Iso Valve CVs]"
'constraint #1
SQL = SQL & "WHERE [Sharpe FT Iso Valve CVs].[Valve Connection]=" & "'" & Connect & "'"
'constraint #2
SQL = SQL & "AND [Sharpe FT Iso Valve CVs].[Valve Size]=" & valveSize
    
Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindFTOutletIsoValveCv = Recordset(0)
    Else: FindFTOutletIsoValveCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindMainSSOVCv(SSOVSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Maxon Cvs].[Valve Cv]"
'define table
SQL = SQL & "FROM [Maxon Cvs]"
'constraint #1
SQL = SQL & "WHERE [Maxon Cvs].[Valve Size]=" & SSOVSize
    
Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMainSSOVCv = Recordset(0)
    Else: FindMainSSOVCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindSystemBypassValveSize(CvRequired As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Baelz Cv Table].[Valve Size]"
'define table
SQL = SQL & "FROM [Baelz Cv Table]"
'constraint #1
SQL = SQL & "WHERE [Baelz Cv Table].[80%]>" & CvRequired
    
Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindSystemBypassValveSize = Recordset(0)
    Else: FindSystemBypassValveSize = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindPilotSSOVCv(SSOVSize As Single, Manufacturer As String)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

If Left(Manufacturer, 5) = "Maxon" Then
'Pull cost from material table
SQL = "SELECT [Maxon Cvs].[Valve Cv]"
'define table
SQL = SQL & "FROM [Maxon Cvs]"
'constraint #1
SQL = SQL & "WHERE [Maxon Cvs].[Valve Size]=" & SSOVSize
Else:
'Pull cost from material table
SQL = "SELECT [ASCO CVs].[Cv]"
'define table
SQL = SQL & "FROM [ASCO CVs]"
'constraint #1
SQL = SQL & "WHERE [ASCO CVs].[Valve Size]=" & SSOVSize
End If

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindPilotSSOVCv = Recordset(0)
    Else: FindPilotSSOVCv = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function FindNeedleValveCV(ALOSize As Single)

Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant

'Pull cost from material table
SQL = "SELECT [Eclipse ALO CVs].[Cv]"
'define table
SQL = SQL & "FROM [Eclipse ALO CVs]"
'constraint #1
SQL = SQL & "WHERE [Eclipse ALO CVs].[Valve Size]=" & ALOSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindNeedleValveCV = Recordset(0)
    Else: FindNeedleValveCV = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ValvePrice(ValveType As String, Size As String, Material As String, PressureClass As String, Connection As String, Make As String, ServiceType As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlValveType, sqlPressureClass, sqlMaterial, sqlSize, sqlConnection, sqlMake, sqlServiceType As String

sqlValveType = "'" & ValveType & "'"
sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"
sqlConnection = "'" & Connection & "'"
sqlMake = "'" & Make & "'"
sqlServiceType = "'" & ServiceType & "'"

'Pull cost from material table
SQL = "SELECT [Items].Cost"
'define table
SQL = SQL & " FROM [Items]"
'constraint #1
SQL = SQL & " WHERE [Items].[Item Category]='Valves'"
'constraint #2
SQL = SQL & " AND [Items].[Pressure Class]=" & "'" & PressureClass & "'"
'constraint #3
SQL = SQL & " AND [Items].[Material]=" & sqlMaterial
'constraint #4
SQL = SQL & " AND [Items].[Size]=" & sqlSize
'constraint #5
SQL = SQL & " AND [Items].[Connection Type]=" & sqlConnection
'constraint #6
SQL = SQL & " AND [Items].[Product Group]=" & sqlValveType
'constraint #7
SQL = SQL & " AND [Items].[Manufacturer]=" & sqlMake
'constraint #8
SQL = SQL & " AND [Items].[Service Type]=" & "'" & ServiceType & "'"

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ValvePrice = Recordset(0)
    Else: ValvePrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function WeldTime(Material As String, PressureClass As String, Size As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlPressureClass, sqlMaterial, sqlSize As String

sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"

'Pull cost from material table
SQL = "SELECT [Welding].NewTime"
'define table
SQL = SQL & " FROM [Welding]"
'constraint #1
SQL = SQL & " WHERE [Welding].[Pressure Class]=" & sqlPressureClass
'constraint #2
SQL = SQL & " AND [Welding].[Material]=" & sqlMaterial
'constraint #3
SQL = SQL & " AND [Welding].[Pipe Size]=" & sqlSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    WeldTime = Recordset(0)
    Else: WeldTime = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function BellowsPrice(Size As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlSize As String

sqlSize = "'" & Size & "'"


'Pull cost from material table
SQL = "SELECT [Items].Cost"
'define table
SQL = SQL & " FROM [Items]"
'constraint #1
SQL = SQL & " WHERE [Items].[Item Category]='Expansion Joints'"
'constraint #2
SQL = SQL & " AND [Items].[Size]=" & sqlSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BellowsPrice = Recordset(0)
    Else: BellowsPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function BurnerPrice(Size As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlSize As String


sqlSize = "'" & Size & "'"


'Pull cost from material table
SQL = "SELECT [Items].Cost"
'define table
SQL = SQL & " FROM [Items]"
'constraint #1
SQL = SQL & " WHERE [Items].[Item Category]='Burners'"
'constraint #2
SQL = SQL & " AND [Items].[Size]=" & sqlSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    BurnerPrice = Recordset(0)
    Else: BurnerPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function HeaderWeldTime(Material As String, PressureClass As String, Size As String, radius As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlPressureClass, sqlMaterial, sqlSize, sqlRadius As String

sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"
sqlRadius = "'" & radius & "'"

'Pull time from weld table
SQL = "SELECT [Header Welds].NewTime"
'define table
SQL = SQL & " FROM [Header Welds]"
'constraint #1
SQL = SQL & " WHERE [Header Welds].[Pressure Class]=" & sqlPressureClass
'constraint #2
SQL = SQL & " AND [Header Welds].[Material]=" & sqlMaterial
'constraint #3
SQL = SQL & " AND [Header Welds].[Pipe Size]=" & sqlSize
'constraint #4
SQL = SQL & " AND [Header Welds].[Radius]=" & sqlRadius

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    HeaderWeldTime = Recordset(0)
    Else: HeaderWeldTime = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function WeldPrepTime(Material As String, PressureClass As String, Size As String, PrepType As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlPressureClass, sqlMaterial, sqlSize, sqlPrepType As String

sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"
sqlPrepType = "'" & PrepType & "'"

'Pull time from weld table
SQL = "SELECT [Weld Prep].NewTime"
'define table
SQL = SQL & " FROM [Weld Prep]"
'constraint #1
SQL = SQL & " WHERE [Weld Prep].[Pressure Class]=" & sqlPressureClass
'constraint #2
SQL = SQL & " AND [Weld Prep].[Material]=" & sqlMaterial
'constraint #3
SQL = SQL & " AND [Weld Prep].[Pipe Size]=" & sqlSize
'constraint #4
SQL = SQL & " AND [Weld Prep].[Type]=" & sqlPrepType

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    WeldPrepTime = Recordset(0)
    Else: WeldPrepTime = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function PipePrice(Material As String, PressureClass As String, Size As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlPressureClass, sqlMaterial, sqlSize As String

sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"

'Pull cost from material table
SQL = "SELECT [Items].Cost"
'define table
SQL = SQL & " FROM [Items]"
'constraint #1
SQL = SQL & " WHERE [Items].[Product Group]='Pipe'"
'constraint #2
SQL = SQL & " AND [Items].[Pressure Class]=" & sqlPressureClass
'constraint #3
SQL = SQL & " AND [Items].[Material]=" & sqlMaterial
'constraint #4
SQL = SQL & " AND [Items].[Size]=" & sqlSize

Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    PipePrice = Recordset(0)
    Else: PipePrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function ReturnPrice(Material As String, PressureClass As String, Size As String, radius As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlPressureClass, sqlMaterial, sqlSize, sqlRadius As String

sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"
sqlRadius = "'" & radius & "'"

'Pull cost from material table
SQL = "SELECT [Items].Cost"
'define table
SQL = SQL & " FROM [Items]"
'constraint #1
SQL = SQL & " WHERE [Items].[Product Group]='Return'"
'constraint #2
SQL = SQL & " AND [Items].[Pressure Class]=" & sqlPressureClass
'constraint #3
SQL = SQL & " AND [Items].[Material]=" & sqlMaterial
'constraint #4
SQL = SQL & " AND [Items].[Size]=" & sqlSize
'constraint #5
SQL = SQL & " AND [Items].[Configuration]=" & sqlRadius


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    ReturnPrice = Recordset(0)
    Else: ReturnPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function
Function CapPrice(Material As String, PressureClass As String, Size As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlPressureClass, sqlMaterial, sqlSize, sqlRadius As String

sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"

'Pull cost from material table
SQL = "SELECT [Items].Cost"
'define table
SQL = SQL & " FROM [Items]"
'constraint #1
SQL = SQL & " WHERE [Items].[Product Group]='Cap'"
'constraint #2
SQL = SQL & " AND [Items].[Pressure Class]=" & sqlPressureClass
'constraint #3
SQL = SQL & " AND [Items].[Material]=" & sqlMaterial
'constraint #4
SQL = SQL & " AND [Items].[Size]=" & sqlSize


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    CapPrice = Recordset(0)
    Else: CapPrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function FlangePrice(Material As String, PressureClass As String, Size As String, Connection As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Lookup Database.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As String

Dim sqlPressureClass, sqlMaterial, sqlSize, sqlConnection As String

sqlPressureClass = "'" & PressureClass & "'"
sqlMaterial = "'" & Material & "'"
sqlSize = "'" & Size & "'"
sqlConnection = "'" & Connection & "'"

'Pull cost from material table
SQL = "SELECT [Items].Cost"
'define table
SQL = SQL & " FROM [Items]"
'constraint #1
SQL = SQL & " WHERE [Items].[Product Group]='Flange'"
'constraint #2
SQL = SQL & " AND [Items].[Pressure Class]=" & sqlPressureClass
'constraint #3
SQL = SQL & " AND [Items].[Material]=" & sqlMaterial
'constraint #4
SQL = SQL & " AND [Items].[Size]=" & sqlSize
'constraint #5
SQL = SQL & " AND [Items].[Configuration]=" & sqlConnection


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FlangePrice = Recordset(0)
    Else: FlangePrice = "Not Found"
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function

Function TubeStabTime(length As String)

'times are in hours'
'length in feet'

If length = 0 Then
    TubeStabTime = 0
ElseIf length > 0 And length < 12 Then
    TubeStabTime = 0.31
ElseIf length >= 12 And length < 22 Then
    TubeStabTime = 0.42
ElseIf length >= 22 And length < 30 Then
    TubeStabTime = 0.62
ElseIf length >= 30 And length < 40 Then
    TubeStabTime = 0.83
ElseIf length >= 40 And length < 50 Then
    TubeStabTime = 0.94
ElseIf length >= 50 And length < 60 Then
    TubeStabTime = 1.25
ElseIf length >= 60 And length < 70 Then
    TubeStabTime = 1.6
ElseIf length >= 70 And length < 80 Then
    TubeStabTime = 1.9
Else:
    TubeStabTime = 2.5
End If

End Function


Function ShellPlateWelding(length As String, diameter As String, thickness As String)
' length in feet'
' diameter in inches'

Dim circumference As String
Dim circumferencewelds As String
Dim lengthwelds As String
Dim passes As String

circumference = diameter * 3.14159265

If length > 8 Then
    circumferencewelds = 1 + Application.WorksheetFunction.RoundUp(length / 8, 0)
Else:
    circumferencewelds = 2
End If

If circumference > 480 Then
    lengthwelds = Application.WorksheetFunction.RoundUp(circumference / 480, 0)
Else:
    lengthwelds = 1
End If

If thickness = 0.25 Then
    passes = 2
Else:
    passes = 3
End If

ShellPlateWelding = (circumferencewelds * circumference + lengthwelds * length * 12) * passes
' shell plate welding in inches'

End Function

Function SaddleWelding(diameter As String)

Dim circumference, depth, stiffeners As String

circumference = diameter * Application.WorksheetFunction.pi()

If diameter <= 18 Then
    depth = 4
    stiffeners = 2
ElseIf diameter <= 30 Then
    depth = 4
    stiffeners = 3
ElseIf diameter <= 42 Then
    depth = 6
    stiffeners = 3
ElseIf diameter <= 54 Then
    depth = 8
    stiffeners = 3
ElseIf diameter <= 60 Then
    depth = 8
    stiffeners = 5
ElseIf diameter <= 84 Then
    depth = 10
    stiffeners = 5
Else:
    depth = 12
    stiffeners = 5
End If

SaddleWelding = (circumference / 3 + stiffeners * depth) * 2

End Function


Function SkidWelding(beamWidth As String, beamHeight As String, webThickness As String, supports As String)

Dim beamSA As String
'beamSA is beam surface area'
' all units are inches'

beamSA = (beamHeight * 2) + (beamWidth * 4) - (webThickness * 2)

SkidWelding = beamSA * 2 * supports
'result value is in inches'


End Function

Function SkidPainting(beamWidth As String, beamHeight As String, webThickness As String, supports As String, skidWidth As String, skidLength As String)

Dim beamSA As String
'beamSA is beam surface area'
' all units are inches'

beamSA = (beamHeight * 2) + (beamWidth * 4) - (webThickness * 2)
SkidPainting = beamSA * (skidLength * 2 + skidWidth * supports)

End Function


Function materialHandling(weight As String)

'weight in pounds, handling time in hours'

If weight < 20000 Then
    materialHandling = 24
ElseIf weight >= 20000 And weight < 60000 Then
    materialHandling = 0.12 * weight / 100
ElseIf weight >= 60000 And weight < 100000 Then
    materialHandling = 78
Else:
    materialHandling = 0.078 * weight / 100
End If
    

End Function

Function FindMargin(ProductType As String, Application As String, CustType As String, Scope As String)
Const ConnectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=C:\Pricing Database\Target Margins.accdb;Persist Security Info=False"
Dim Recordset As ADODB.Recordset
Dim SQL As Variant


'Pull target margin from access table
SQL = "SELECT [Sheet1].[Avg]"
'define table
SQL = SQL & " FROM [Sheet1]"
SQL = SQL & " WHERE [Sheet1].[Product Type]='" & ProductType & "'"
SQL = SQL & " AND [Sheet1].[Application]='" & Application & "'"
SQL = SQL & " AND [Sheet1].[Customer Type]='" & CustType & "'"
SQL = SQL & " AND [Sheet1].[Scope]='" & Scope & "'"


Set Recordset = New ADODB.Recordset
Call Recordset.Open(SQL, ConnectionString, adOpenForwardOnly, adLockReadOnly)
If Not Recordset.EOF Then
    FindMargin = Recordset(0)
    FindMargin = CSng(FindMargin)
    Else: FindMargin = 0
End If

If (Recordset.State And ObjectStateEnum.adStateOpen) Then Recordset.Close
Set Recordset = Nothing

End Function





-------------------------------------------------------------------------------
VBA MACRO RadiantFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/RadiantFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Function ExchangeFactor(ratio As Single, Emissivity As Single)
   'Exchange factor based on ratio
    
    Dim E1, E2, R1, R2 As Double
    
    
    Select Case ratio
        Case 0 To 0.5
            'E1 = -1.4286 * Emissivity ^ 2 + 1.9657 * Emissivity + 0.1697
            E1 = Emissivity - 0.01
            E2 = -0.375 * Emissivity ^ 2 + 1.2632 * Emissivity + 0.0184
            R1 = 0
            R2 = 0.5
        Case 0.5 To 1
            E1 = -0.375 * Emissivity ^ 2 + 1.2632 * Emissivity + 0.0184
            E2 = -0.7232 * Emissivity ^ 2 + 1.4995 * Emissivity + 0.0548
            R1 = 0.5
            R2 = 1
        Case 1 To 1.5
            E1 = -0.7232 * Emissivity ^ 2 + 1.4995 * Emissivity + 0.0548
            E2 = -0.8839 * Emissivity ^ 2 + 1.6141 * Emissivity + 0.0909
            R1 = 1
            R2 = 1.5
        Case 1.5 To 2
            E1 = -0.8839 * Emissivity ^ 2 + 1.6141 * Emissivity + 0.0909
            E2 = -0.7589 * Emissivity ^ 2 + 1.4402 * Emissivity + 0.1911
            R1 = 1.5
            R2 = 2
        Case 2 To 2.5
            E1 = -0.7589 * Emissivity ^ 2 + 1.4402 * Emissivity + 0.1911
            E2 = -1.3036 * Emissivity ^ 2 + 1.9232 * Emissivity + 0.1357
            R1 = 2
            R2 = 2.5
        Case 2.5 To 3
            E1 = -1.3036 * Emissivity ^ 2 + 1.9232 * Emissivity + 0.1357
            E2 = -1.6786 * Emissivity ^ 2 + 2.2579 * Emissivity + 0.117
            R1 = 2.5
            R2 = 3
        Case 3 To 4
            E1 = -1.6786 * Emissivity ^ 2 + 2.2579 * Emissivity + 0.117
            E2 = -1.0714 * Emissivity ^ 2 + 1.6186 * Emissivity + 0.2849
            R1 = 3
            R2 = 4
        Case 4 To 5
            E1 = -1.0714 * Emissivity ^ 2 + 1.6186 * Emissivity + 0.2849
            E2 = -0.8214 * Emissivity ^ 2 + 1.3264 * Emissivity + 0.3776
            R1 = 4
            R2 = 5
        Case 5 To 6
            E1 = -0.8214 * Emissivity ^ 2 + 1.3264 * Emissivity + 0.3776
            E2 = -0.8857 * Emissivity ^ 2 + 1.3426 * Emissivity + 0.4042
            R1 = 5
            R2 = 6
        Case 6 To 7
            E1 = -0.8857 * Emissivity ^ 2 + 1.3426 * Emissivity + 0.4042
            E2 = -0.7321 * Emissivity ^ 2 + 1.1861 * Emissivity + 0.4517
            R1 = 6
            R2 = 7
End Select
ExchangeFactor = ((ratio - R1) / (R2 - R1)) * (E2 - E1) + E1

End Function
Function GasRadCoeff(TubeWallTemp As Single, AvgGasTemp)
Dim GRC1, GRC2, t1, t2 As Double

Select Case TubeWallTemp
    Case 100 To 400
        GRC1 = 5E-07 * AvgGasTemp ^ 2 + 0.0005 * AvgGasTemp + 0.23
        GRC2 = -2E-07 * AvgGasTemp ^ 2 + 0.0022 * AvgGasTemp - 0.28
        t1 = 100
        t2 = 400
    Case 400 To 600
        GRC1 = -2E-07 * AvgGasTemp ^ 2 + 0.0022 * AvgGasTemp - 0.28
        GRC2 = 2E-08 * AvgGasTemp ^ 2 + 0.002 * AvgGasTemp - 0.03
        t1 = 400
        t2 = 600
    Case 600 To 800
        GRC1 = 2E-08 * AvgGasTemp ^ 2 + 0.002 * AvgGasTemp - 0.03
        GRC2 = -9E-08 * AvgGasTemp ^ 2 + 0.0023 * AvgGasTemp + 0.22
        t1 = 600
        t2 = 800
    Case 800 To 1000
        GRC1 = -9E-08 * AvgGasTemp ^ 2 + 0.0023 * AvgGasTemp + 0.22
        GRC2 = 2E-07 * AvgGasTemp ^ 2 + 0.0012 * AvgGasTemp + 1.72
        t1 = 800
        t2 = 1000
    Case 1000 To 1200
        GRC1 = 2E-07 * AvgGasTemp ^ 2 + 0.0012 * AvgGasTemp + 1.72
        GRC2 = 9E-08 * AvgGasTemp ^ 2 + 0.0013 * AvgGasTemp + 2.046
        t1 = 1000
        t2 = 1200
    Case 1200 To 1400
        GRC1 = 9E-08 * AvgGasTemp ^ 2 + 0.0013 * AvgGasTemp + 2.046
        GRC2 = 2E-07 * AvgGasTemp ^ 2 + 0.0012 * AvgGasTemp + 2.29
        t1 = 1200
        t2 = 1400
    Case 1400 To 1600
        GRC1 = 2E-07 * AvgGasTemp ^ 2 + 0.0012 * AvgGasTemp + 2.29
        GRC2 = -3E-07 * AvgGasTemp ^ 2 + 0.0014 * AvgGasTemp + 4.08
        t1 = 1400
        t2 = 1600
End Select
GasRadCoeff = ((TubeWallTemp - t1) / (t2 - t1)) * (GRC2 - GRC1) + GRC1
End Function
Function GasEmissivity(Pl As Single, GasTemp As Single)
Dim E1, E2, t1, t2 As Double


Select Case GasTemp
    Case 100 To 1400
        E1 = 0.0038 * Pl ^ 3 - 0.0452 * Pl ^ 2 + 0.2027 * Pl + 0.2205
        E2 = 0.0038 * Pl ^ 3 - 0.0452 * Pl ^ 2 + 0.2027 * Pl + 0.2005
        t1 = 200
        t2 = 1400
    Case 1400 To 1800
        E1 = 0.0038 * Pl ^ 3 - 0.0452 * Pl ^ 2 + 0.2027 * Pl + 0.2005
        E2 = 0.0038 * Pl ^ 3 - 0.0452 * Pl ^ 2 + 0.2027 * Pl + 0.1705
        t1 = 1400
        t2 = 1800
    Case 1800 To 2000
        E1 = 0.0038 * Pl ^ 3 - 0.0452 * Pl ^ 2 + 0.2027 * Pl + 0.1705
        E2 = 0.0038 * Pl ^ 3 - 0.0452 * Pl ^ 2 + 0.2027 * Pl + 0.1505
        t1 = 1800
        t2 = 2000
    Case 2000 To 2400
        E1 = 0.0038 * Pl ^ 3 - 0.0452 * Pl ^ 2 + 0.2027 * Pl + 0.1505
        E2 = 0.0017 * Pl ^ 3 - 0.0247 * Pl ^ 2 + 0.1498 * Pl + 0.1239
        t1 = 2000
        t2 = 2400
    Case 2400 To 2800
        E1 = 0.0017 * Pl ^ 3 - 0.0247 * Pl ^ 2 + 0.1498 * Pl + 0.1239
        E2 = 0.0011 * Pl ^ 3 - 0.0168 * Pl ^ 2 + 0.1186 * Pl + 0.1157
        t1 = 2400
        t2 = 2800
    Case 2800 To 3600
        E1 = 0.0011 * Pl ^ 3 - 0.0168 * Pl ^ 2 + 0.1186 * Pl + 0.1157
        E2 = 0.0011 * Pl ^ 3 - 0.0168 * Pl ^ 2 + 0.1186 * Pl + 0.0657
        t1 = 2800
        t2 = 3600
End Select
GasEmissivity = ((GasTemp - t1) / (t2 - t1)) * (E2 - E1) + E1

End Function

-------------------------------------------------------------------------------
VBA MACRO RefpropCode.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/RefpropCode'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

'An environment variable can be set to specify the locations of the fluid files.  This can
'be done under Start/Settings/Control Panel/System/Advanced/Environment Variables.  A new
'variable called RPPrefix should be created with a value of "C:\Program Files\REFPROP" (or
'the directory where Refprop is located).  The Refprop FAQ page shows pictures of these variables.

'Alternatively:
'Change the following to the location of the fluid files and of the mixture files (those with .MIX).
'If the file C:\REFPROP\REFPRP64.DLL is not located in the application workspace, then see note below.
'DO NOT IMPLEMENT BOTH OF THESE METHODS, change either the paths below or modify the environment variables, not both.

Private Const FluidsDirectory As String = "fluids\"
Private Const MixturesDirectory As String = "mixtures\"
'----Usually the files are located here, so comment out the two lines above and use these two:
'Private Const FluidsDirectory As String = "c:\Program Files\REFPROP\fluids\"
'Private Const MixturesDirectory As String = "c:\Program Files\REFPROP\mixtures\"

'REFPROP Excel Functions
'Arguments:  (FluidName, InpCode, Units, Prop1, Prop2)
'FluidName = text,  fluid must be either in Fluids or Mixtures sub directories.
'InpCode   = name and order of Prop1 and Prop2.
'            "TP" would mean Prop1 is Temperature, Prop2 is Pressure (need quotes)
'            Valid InpCodes:  TP,TD,TH,TS,TE,TQ,PD,PH,PS,PE,PQ,DH,DS,DE,HS
'            To define saturated liquid or vapor inputs:  TLIQ, TVAP, PLIQ, PVAP
'            Other:  Crit, Trip, TMelt, PMelt, TSubl, PSubl
'
'            The word "Optional" appears in the argument listings below to indicate that
'            PROP2 is not always required depending on the InpCode argument.
'            In some cases, PROP1 is also optional.
'
'Units     = "SI", "SI with C" (or just "C"), "Molar SI", "E", "molar E", "cgs", "mks", "Mixed" (need quotes).  "SI" is used by default if no input is given. (Unless DefaultUnits changed in VBA code)
'Prop1     = numerical value of the first input property (in the units of the previous line)
'Prop2     = numerical value of the second input property (if required).

'To call the functions from a module located outside the Refprop module, use something like this:
'Den = Application.Run("'REFPROP.XLA'!Density", Fluid, Props, Units, Temp, Pres)


'Function Temperature           (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Pressure              (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Density               (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function LiquidDensity         (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function VaporDensity          (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Volume                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function CompressibilityFactor (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Energy                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Enthalpy              (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function LiquidEnthalpy        (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function VaporEnthalpy         (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Entropy               (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function IsochoricHeatCapacity (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Cv                    (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function IsobaricHeatCapacity  (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Cp                    (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function SpecificHeatInput     (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Csat                  (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function SpeedOfSound          (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Sound                 (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dPdrho                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function d2Pdrho2              (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dPdT                  (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dPdTsat               (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function drhodT                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dHdT_D                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dHdT_P                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dHdD_T                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dHdD_P                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dHdP_T                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function dHdP_D                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Cstar                 (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Quality               (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function QualityMole           (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function QualityMass           (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function LatentHeat            (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function HeatOfVaporization    (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function HeatOfCombustion      (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function GrossHeatingValue     (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function NetHeatingValue       (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function JouleThomson          (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function IsentropicExpansionCoef       (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function IsothermalCompressibility     (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function VolumeExpansivity             (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function AdiabaticCompressibility      (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function AdiabaticBulkModulus          (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function IsothermalExpansionCoef       (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function IsothermalBulkModulus         (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function SecondVirial          (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Viscosity             (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function KinematicViscosity    (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function ThermalConductivity   (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function ThermalDiffusivity    (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function Prandtl               (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function SurfaceTension        (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function DielectricConstant    (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function MolarMass             (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function EOSMax                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function EOSMin                (FluidName, InpCode, Units, Prop1, Optional Prop2)
'Function LiquidFluidString     (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2)
'Function VaporFluidString      (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2)
'Function FluidString           (Nmes, Comps, Optional massmole As String)
'Function PropertyUnits         (InpCode, Units)

'In the following, i is the component number in the mixture (where the maximum value of i can be 20)

'Function ComponentName         (FluidName, i)
'Function MoleFraction          (FluidName, i)
'Function MassFraction          (FluidName, i)
'Function LiquidMoleFraction    (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function VaporMoleFraction     (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function LiquidMassFraction    (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function VaporMassFraction     (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function Activity              (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function ActivityCoefficient   (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function ChemicalPotential     (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function Fugacity              (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function FugacityCoefficient   (FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
'Function Mole2Mass             (FluidName, i, Prop1, Prop2, Optional Prop3, Optional Prop4, Optional Prop5, Optional Prop6, Optional Prop7, Optional Prop8, Optional Prop9, Optional Prop10, Optional Prop11, Optional Prop12, Optional Prop13, Optional Prop14, Optional Prop15, Optional Prop16, Optional Prop17, Optional Prop18, Optional Prop19, Optional Prop20)
'Function Mass2Mole             (FluidName, i, Prop1, Prop2, Optional Prop3, Optional Prop4, Optional Prop5, Optional Prop6, Optional Prop7, Optional Prop8, Optional Prop9, Optional Prop10, Optional Prop11, Optional Prop12, Optional Prop13, Optional Prop14, Optional Prop15, Optional Prop16, Optional Prop17, Optional Prop18, Optional Prop19, Optional Prop20)

'The following functions are for Document Reference
'Function WorkBookName
'Function RefpropXLSVersionNumber
'Function RefpropDLLVersionNumber
'Function WhereAreREFPROPfunctions
'Function WhereIsWorkbook
'Function SeeFileLinkSources
'Function SelectedDefaultUnits
'

'In order for Excel to access the C:\REFPROP\REFPRP64.DLL file, you will need to do one of the items below:
'  -  Place your Excel file in the REFPROP directory
'  -  Change the path and environment variables as described at the top of the "Examples" sheet
'  -  Below, replace all "C:\REFPROP\REFPRP64.DLL" with "C:\Program Files\REFPROP\C:\REFPROP\REFPRP64.DLL" (or the subdirectory where REFPROP was installed)
'  -  Place a copy of C:\REFPROP\REFPRP64.DLL in your directory where your Excel files are located (not preferred)
Private Const MaxComps As Integer = 20

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Sub SETUPdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long, ByVal hfld As String, ByVal hfmix As String, ByVal hrf As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long)
        Private Declare PtrSafe Sub SETREFdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hrf As String, ixflag As Long, x0 As Double, h0 As Double, s0 As Double, t0 As Double, p0 As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare PtrSafe Sub SETMIXdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hmxnme As String, ByVal hfmix As String, ByVal hrf As String, ncc As Long, ByVal hfile As String, x As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long, ln5 As Long)
        Private Declare PtrSafe Sub SETMODdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long, ByVal htype As String, ByVal hmix As String, ByVal hcomp As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long)
        Private Declare PtrSafe Sub SETPATHdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hpath As String, ln As Long)
        Private Declare PtrSafe Sub GERG04dll Lib "C:\REFPROP\REFPRP64.DLL" (nc As Long, iflag As Long, ierr As Long, ByVal herr As String, ln1 As Long)
        Private Declare PtrSafe Sub SETKTVdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, jcomp As Long, ByVal hmodij As String, fij As Double, ByVal hfmix As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare PtrSafe Sub GETKTVdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, jcomp As Long, ByVal hmodij As String, fij As Double, ByVal hfmix As String, ByVal hfij As String, ByVal hbinp As String, ByVal hmxrul As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long, ln5 As Long)
        Private Declare PtrSafe Sub GETFIJdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hmodij As String, fij As Double, ByVal hfij As String, ByVal hmxrul As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare PtrSafe Sub PREOSdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long)
        Private Declare PtrSafe Sub SETAGAdll Lib "C:\REFPROP\REFPRP64.DLL" (ierr As Long, ByVal herr As String, ln1 As Long)
        Private Declare PtrSafe Sub UNSETAGAdll Lib "C:\REFPROP\REFPRP64.DLL" ()
        Private Declare PtrSafe Sub NAMEdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, ByVal hnam As String, ByVal hn80 As String, ByVal hcasn As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare PtrSafe Sub PUREFLDdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long)
        Private Declare PtrSafe Sub SETNCdll Lib "C:\REFPROP\REFPRP64.DLL" (ncomp As Long)
        Private Declare PtrSafe Sub RPVersion Lib "C:\REFPROP\REFPRP64.DLL" (ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub CRITPdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tc As Double, pc As Double, dc As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub MAXTdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tm As Double, pm As Double, dm As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub MAXPdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tm As Double, pm As Double, dm As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub REDXdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, T As Double, d As Double)
        Private Declare PtrSafe Sub THERMdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, hjt As Double)
        Private Declare PtrSafe Sub THERM2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, z As Double, hjt As Double, aH As Double, g As Double, kappa As Double, beta As Double, dPdD As Double, d2PdD2 As Double, dPdT As Double, dDdT As Double, dDdP As Double, d2PT2 As Double, d2PdTD As Double, spare3 As Double, spare4 As Double)
        Private Declare PtrSafe Sub THERM3dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, kappa As Double, beta As Double, isenk As Double, kt As Double, betas As Double, bs As Double, kkt As Double, thrott As Double, pi As Double, spht As Double)
        Private Declare PtrSafe Sub THERM0dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, a As Double, g As Double)
        Private Declare PtrSafe Sub PRESSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double)
        Private Declare PtrSafe Sub ENTHALdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, h As Double)
        Private Declare PtrSafe Sub ENTROdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, s As Double)
        Private Declare PtrSafe Sub CVCPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Cv As Double, Cp As Double)
        Private Declare PtrSafe Sub GIBBSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Ar As Double, Gr As Double)
        Private Declare PtrSafe Sub AGdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, a As Double, g As Double)
        Private Declare PtrSafe Sub PHIXdll Lib "C:\REFPROP\REFPRP64.DLL" (itau As Long, idel As Long, tau As Double, del As Double, x As Double, PHIX As Double)
        Private Declare PtrSafe Sub RESIDUALdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Pr As Double, er As Double, hr As Double, sr As Double, cvr As Double, cpr As Double, Ar As Double, Gr As Double)
        Private Declare PtrSafe Sub CP0dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, Cp As Double)
        Private Declare PtrSafe Sub DPDDdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dPdD As Double)
        Private Declare PtrSafe Sub DPDD2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, d2PdD2 As Double)
        Private Declare PtrSafe Sub DPDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dPdT As Double)
        Private Declare PtrSafe Sub DDDPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dDdP As Double)
        Private Declare PtrSafe Sub DDDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dDdT As Double)
        Private Declare PtrSafe Sub DHD1dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dHdT_D As Double, dHdT_P As Double, dHdD_T As Double, dHdD_P As Double, dHdP_T As Double, dHdP_D As Double)
        Private Declare PtrSafe Sub VIRBdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare PtrSafe Sub VIRCdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, c As Double)
        Private Declare PtrSafe Sub VIRBAdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare PtrSafe Sub VIRCAdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, c As Double)
        Private Declare PtrSafe Sub DBDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, dbt As Double)
        Private Declare PtrSafe Sub B12dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare PtrSafe Sub FGCTYdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, f As Double)
        Private Declare PtrSafe Sub FGCTY2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, f As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub FUGCOFdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, f As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub CHEMPOTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, u As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub ACTVYdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, actv As Double, gamma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub EXCESSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, kph As Long, d As Double, vE As Double, ee As Double, hE As Double, sE As Double, aE As Double, gE As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATSPLNdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATTPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, kph As Long, iGuess As Long, d As Double, rhol As Double, rhov As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, i As Long, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, i As Long, T As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATDdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, x As Double, kph As Long, kr As Long, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATHdll Lib "C:\REFPROP\REFPRP64.DLL" (h As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATEdll Lib "C:\REFPROP\REFPRP64.DLL" (e As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SATSdll Lib "C:\REFPROP\REFPRP64.DLL" (s As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, k3 As Long, t3 As Double, p3 As Double, d3 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub CV2PKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, cv2p As Double, Csat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub CSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, kph As Long, P As Double, rho As Double, Csat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DPTSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, kph As Long, P As Double, rho As Double, Csat As Double, dPdTsat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DLSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DVSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TSATDdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub MELTTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub MLTH2Odll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, p1 As Double, p2 As Double)
        Private Declare PtrSafe Sub MELTPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SUBLTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SUBLPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TPRHOdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, j As Long, i As Long, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TPFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TDFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PDFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, d As Double, x As Double, T As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PHFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, h As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, s As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, e As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub THFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, h As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, s As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, e As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DHFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, h As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, s As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, e As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub HSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (h As Double, s As Double, z As Double, T As Double, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub ESFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (e As Double, s As Double, z As Double, T As Double, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TQFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, q As Double, x As Double, kq As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PQFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, q As Double, x As Double, kq As Long, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, W As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub ABFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (a As Double, b As Double, x As Double, i As Long, ByVal ab As String, dmin As Double, dmax As Double, T As Double, P As Double, d As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare PtrSafe Sub ABFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (a As Double, b As Double, x As Double, kq As Long, ksat As Long, ByVal ab As String, tbub As Double, tdew As Double, pbub As Double, pdew As Double, Dlbub As Double, Dvdew As Double, ybub As Double, xdew As Double, T As Double, P As Double, Dl As Double, Dv As Double, x As Double, y As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long, ln2 As Long)
        Private Declare PtrSafe Sub DBFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, b As Double, x As Double, ByVal ab As String, T As Double, P As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare PtrSafe Sub DBFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, b As Double, x As Double, i As Long, ByVal ab As String, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long, ln2 As Long)
        Private Declare PtrSafe Sub DQFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, q As Double, x As Double, kq As Long, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DSFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, s As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PDFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, d As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PHFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, h As Double, x As Double, kph As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub PSFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, s As Double, x As Double, kph As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TPFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TRNPRPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, eta As Double, tcx As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub DIELECdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, de As Double)
        Private Declare PtrSafe Sub SURFTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, sigma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub SURTENdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, rhol As Double, rhov As Double, xl As Double, xv As Double, sigma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub HEATdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, rho As Double, x As Double, hg As Double, hn As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub CSTARdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, v As Double, x As Double, cs As Double, ts As Double, Ds As Double, ps As Double, ws As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub FPVdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, P As Double, x As Double, f As Double)
        'Private Declare PtrSafe Sub SPECGRdll Lib "C:\REFPROP\REFPRP64.DLL" (t As Double, d As Double, p As Double, Gr As Double)
        Private Declare PtrSafe Sub WMOLdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, wm As Double)
        Private Declare PtrSafe Sub XMASSdll Lib "C:\REFPROP\REFPRP64.DLL" (xmol As Double, xkg As Double, wmix As Double)
        Private Declare PtrSafe Sub XMOLEdll Lib "C:\REFPROP\REFPRP64.DLL" (xkg As Double, xmol As Double, wmix As Double)
        Private Declare PtrSafe Sub QMASSdll Lib "C:\REFPROP\REFPRP64.DLL" (qmol As Double, xl As Double, xv As Double, qkg As Double, xlkg As Double, xvkg As Double, wliq As Double, wvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub QMOLEdll Lib "C:\REFPROP\REFPRP64.DLL" (qkg As Double, xlkg As Double, xvkg As Double, qmol As Double, xl As Double, xv As Double, wliq As Double, wvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub INFOdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, wmm As Double, ttrp As Double, tnbpt As Double, tc As Double, pc As Double, dc As Double, Zc As Double, acf As Double, dip As Double, Rgas As Double)
        Private Declare PtrSafe Sub LIMITSdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, x As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ln As Long)
        Private Declare PtrSafe Sub LIMITXdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, T As Double, d As Double, P As Double, x As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare PtrSafe Sub LIMITKdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, icomp As Long, T As Double, d As Double, P As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare PtrSafe Sub ETAK0dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, eta0 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub ETAK1dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, eta1 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub ETAKRdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, etar As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub ETAKBdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, etab As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub OMEGAdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, epsk As Double, omg As Double)
        Private Declare PtrSafe Sub TCXK0dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, tcx0 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TCXKBdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, tcxb As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Sub TCXKCdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, tcxc As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare PtrSafe Function SetCurDir Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

    #Else
        Private Declare Sub SETUPdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long, ByVal hfld As String, ByVal hfmix As String, ByVal hrf As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long)
        Private Declare Sub SETREFdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hrf As String, ixflag As Long, x0 As Double, h0 As Double, s0 As Double, t0 As Double, p0 As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub SETMIXdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hmxnme As String, ByVal hfmix As String, ByVal hrf As String, ncc As Long, ByVal hfile As String, x As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long, ln5 As Long)
        Private Declare Sub SETMODdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long, ByVal htype As String, ByVal hmix As String, ByVal hcomp As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long)
        Private Declare Sub SETPATHdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hpath As String, ln As Long)
        Private Declare Sub GERG04dll Lib "C:\REFPROP\REFPRP64.DLL" (nc As Long, iflag As Long, ierr As Long, ByVal herr As String, ln1 As Long)
        Private Declare Sub SETKTVdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, jcomp As Long, ByVal hmodij As String, fij As Double, ByVal hfmix As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare Sub GETKTVdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, jcomp As Long, ByVal hmodij As String, fij As Double, ByVal hfmix As String, ByVal hfij As String, ByVal hbinp As String, ByVal hmxrul As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long, ln5 As Long)
        Private Declare Sub GETFIJdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hmodij As String, fij As Double, ByVal hfij As String, ByVal hmxrul As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare Sub PREOSdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long)
        Private Declare Sub SETAGAdll Lib "C:\REFPROP\REFPRP64.DLL" (ierr As Long, ByVal herr As String, ln1 As Long)
        Private Declare Sub UNSETAGAdll Lib "C:\REFPROP\REFPRP64.DLL" ()
        Private Declare Sub NAMEdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, ByVal hnam As String, ByVal hn80 As String, ByVal hcasn As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare Sub PUREFLDdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long)
        Private Declare Sub SETNCdll Lib "C:\REFPROP\REFPRP64.DLL" (ncomp As Long)
        Private Declare Sub RPVersion Lib "C:\REFPROP\REFPRP64.DLL" (ByVal herr As String, ln As Long)
        Private Declare Sub CRITPdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tc As Double, pc As Double, dc As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MAXTdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tm As Double, pm As Double, dm As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MAXPdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tm As Double, pm As Double, dm As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub REDXdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, T As Double, d As Double)
        Private Declare Sub THERMdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, hjt As Double)
        Private Declare Sub THERM2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, z As Double, hjt As Double, aH As Double, g As Double, kappa As Double, beta As Double, dPdD As Double, d2PdD2 As Double, dPdT As Double, dDdT As Double, dDdP As Double, d2PT2 As Double, d2PdTD As Double, spare3 As Double, spare4 As Double)
        Private Declare Sub THERM3dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, kappa As Double, beta As Double, isenk As Double, kt As Double, betas As Double, bs As Double, kkt As Double, thrott As Double, pi As Double, spht As Double)
        Private Declare Sub THERM0dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, A As Double, g As Double)
        Private Declare Sub PRESSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double)
        Private Declare Sub ENTHALdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, h As Double)
        Private Declare Sub ENTROdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, s As Double)
        Private Declare Sub CVCPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Cv As Double, Cp As Double)
        Private Declare Sub GIBBSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Ar As Double, Gr As Double)
        Private Declare Sub AGdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, A As Double, g As Double)
        Private Declare Sub PHIXdll Lib "C:\REFPROP\REFPRP64.DLL" (itau As Long, idel As Long, tau As Double, del As Double, x As Double, PHIX As Double)
        Private Declare Sub RESIDUALdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Pr As Double, er As Double, hr As Double, sr As Double, cvr As Double, cpr As Double, Ar As Double, Gr As Double)
        Private Declare Sub CP0dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, Cp As Double)
        Private Declare Sub DPDDdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dPdD As Double)
        Private Declare Sub DPDD2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, d2PdD2 As Double)
        Private Declare Sub DPDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dPdT As Double)
        Private Declare Sub DDDPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dDdP As Double)
        Private Declare Sub DDDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dDdT As Double)
        Private Declare Sub DHD1dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dHdT_D As Double, dHdT_P As Double, dHdD_T As Double, dHdD_P As Double, dHdP_T As Double, dHdP_D As Double)
        Private Declare Sub VIRBdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare Sub VIRCdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, c As Double)
        Private Declare Sub VIRBAdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare Sub VIRCAdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, c As Double)
        Private Declare Sub DBDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, dbt As Double)
        Private Declare Sub B12dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare Sub FGCTYdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, F As Double)
        Private Declare Sub FGCTY2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, F As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub FUGCOFdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, F As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CHEMPOTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, u As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ACTVYdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, actv As Double, gamma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub EXCESSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, kph As Long, d As Double, vE As Double, ee As Double, hE As Double, sE As Double, aE As Double, gE As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATSPLNdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATTPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, kph As Long, iGuess As Long, d As Double, rhol As Double, rhov As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, i As Long, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, i As Long, T As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATDdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, x As Double, kph As Long, kr As Long, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATHdll Lib "C:\REFPROP\REFPRP64.DLL" (h As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATEdll Lib "C:\REFPROP\REFPRP64.DLL" (e As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATSdll Lib "C:\REFPROP\REFPRP64.DLL" (s As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, k3 As Long, t3 As Double, p3 As Double, d3 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CV2PKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, cv2p As Double, Csat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, kph As Long, P As Double, rho As Double, Csat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DPTSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, kph As Long, P As Double, rho As Double, Csat As Double, dPdTsat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DLSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DVSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TSATDdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MELTTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MLTH2Odll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, p1 As Double, p2 As Double)
        Private Declare Sub MELTPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SUBLTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SUBLPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TPRHOdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, j As Long, i As Long, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TPFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TDFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PDFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, d As Double, x As Double, T As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PHFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, h As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, s As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, e As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub THFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, h As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, s As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, e As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DHFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, h As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, s As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, e As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub HSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (h As Double, s As Double, z As Double, T As Double, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ESFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (e As Double, s As Double, z As Double, T As Double, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TQFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, q As Double, x As Double, kq As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PQFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, q As Double, x As Double, kq As Long, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ABFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (A As Double, b As Double, x As Double, i As Long, ByVal ab As String, dmin As Double, dmax As Double, T As Double, P As Double, d As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub ABFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (A As Double, b As Double, x As Double, kq As Long, ksat As Long, ByVal ab As String, tbub As Double, tdew As Double, pbub As Double, pdew As Double, Dlbub As Double, Dvdew As Double, ybub As Double, xdew As Double, T As Double, P As Double, Dl As Double, Dv As Double, x As Double, y As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long, ln2 As Long)
        Private Declare Sub DBFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, b As Double, x As Double, ByVal ab As String, T As Double, P As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub DBFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, b As Double, x As Double, i As Long, ByVal ab As String, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long, ln2 As Long)
        Private Declare Sub DQFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, q As Double, x As Double, kq As Long, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DSFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, s As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PDFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, d As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PHFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, h As Double, x As Double, kph As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PSFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, s As Double, x As Double, kph As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TPFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TRNPRPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, eta As Double, tcx As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DIELECdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, de As Double)
        Private Declare Sub SURFTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, sigma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SURTENdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, rhol As Double, rhov As Double, xl As Double, xv As Double, sigma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub HEATdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, rho As Double, x As Double, hg As Double, hn As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CSTARdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, v As Double, x As Double, cs As Double, ts As Double, Ds As Double, ps As Double, ws As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub FPVdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, P As Double, x As Double, F As Double)
        'Private Declare  Sub SPECGRdll Lib "C:\REFPROP\REFPRP64.DLL" (t As Double, d As Double, p As Double, Gr As Double)
        Private Declare Sub WMOLdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, wm As Double)
        Private Declare Sub XMASSdll Lib "C:\REFPROP\REFPRP64.DLL" (xmol As Double, xkg As Double, wmix As Double)
        Private Declare Sub XMOLEdll Lib "C:\REFPROP\REFPRP64.DLL" (xkg As Double, xmol As Double, wmix As Double)
        Private Declare Sub QMASSdll Lib "C:\REFPROP\REFPRP64.DLL" (qmol As Double, xl As Double, xv As Double, qkg As Double, xlkg As Double, xvkg As Double, wliq As Double, wvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub QMOLEdll Lib "C:\REFPROP\REFPRP64.DLL" (qkg As Double, xlkg As Double, xvkg As Double, qmol As Double, xl As Double, xv As Double, wliq As Double, wvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub INFOdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, wmm As Double, ttrp As Double, tnbpt As Double, tc As Double, pc As Double, dc As Double, Zc As Double, acf As Double, dip As Double, Rgas As Double)
        Private Declare Sub LIMITSdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, x As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ln As Long)
        Private Declare Sub LIMITXdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, T As Double, d As Double, P As Double, x As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub LIMITKdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, icomp As Long, T As Double, d As Double, P As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub ETAK0dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, eta0 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ETAK1dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, eta1 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ETAKRdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, etar As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ETAKBdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, etab As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub OMEGAdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, epsk As Double, omg As Double)
        Private Declare Sub TCXK0dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, tcx0 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TCXKBdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, tcxb As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TCXKCdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, tcxc As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Function SetCurDir Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
    #End If
#Else
        Private Declare Sub SETUPdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long, ByVal hfld As String, ByVal hfmix As String, ByVal hrf As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long)
        Private Declare Sub SETREFdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hrf As String, ixflag As Long, x0 As Double, h0 As Double, s0 As Double, t0 As Double, p0 As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub SETMIXdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hmxnme As String, ByVal hfmix As String, ByVal hrf As String, ncc As Long, ByVal hfile As String, x As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long, ln5 As Long)
        Private Declare Sub SETMODdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long, ByVal htype As String, ByVal hmix As String, ByVal hcomp As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long)
        Private Declare Sub SETPATHdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hpath As String, ln As Long)
        Private Declare Sub GERG04dll Lib "C:\REFPROP\REFPRP64.DLL" (nc As Long, iflag As Long, ierr As Long, ByVal herr As String, ln1 As Long)
        Private Declare Sub SETKTVdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, jcomp As Long, ByVal hmodij As String, fij As Double, ByVal hfmix As String, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare Sub GETKTVdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, jcomp As Long, ByVal hmodij As String, fij As Double, ByVal hfmix As String, ByVal hfij As String, ByVal hbinp As String, ByVal hmxrul As String, ln1 As Long, ln2 As Long, ln3 As Long, ln4 As Long, ln5 As Long)
        Private Declare Sub GETFIJdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal hmodij As String, fij As Double, ByVal hfij As String, ByVal hmxrul As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare Sub PREOSdll Lib "C:\REFPROP\REFPRP64.DLL" (i As Long)
        Private Declare Sub SETAGAdll Lib "C:\REFPROP\REFPRP64.DLL" (ierr As Long, ByVal herr As String, ln1 As Long)
        Private Declare Sub UNSETAGAdll Lib "C:\REFPROP\REFPRP64.DLL" ()
        Private Declare Sub NAMEdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, ByVal hnam As String, ByVal hn80 As String, ByVal hcasn As String, ln1 As Long, ln2 As Long, ln3 As Long)
        Private Declare Sub PUREFLDdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long)
        Private Declare Sub SETNCdll Lib "C:\REFPROP\REFPRP64.DLL" (ncomp As Long)
        Private Declare Sub RPVersion Lib "C:\REFPROP\REFPRP64.DLL" (ByVal herr As String, ln As Long)
        Private Declare Sub CRITPdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tc As Double, pc As Double, dc As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MAXTdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tm As Double, pm As Double, dm As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MAXPdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, tm As Double, pm As Double, dm As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub REDXdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, T As Double, d As Double)
        Private Declare Sub THERMdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, hjt As Double)
        Private Declare Sub THERM2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, z As Double, hjt As Double, aH As Double, g As Double, kappa As Double, beta As Double, dPdD As Double, d2PdD2 As Double, dPdT As Double, dDdT As Double, dDdP As Double, d2PT2 As Double, d2PdTD As Double, spare3 As Double, spare4 As Double)
        Private Declare Sub THERM3dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, kappa As Double, beta As Double, isenk As Double, kt As Double, betas As Double, bs As Double, kkt As Double, thrott As Double, pi As Double, spht As Double)
        Private Declare Sub THERM0dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, A As Double, g As Double)
        Private Declare Sub PRESSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double)
        Private Declare Sub ENTHALdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, h As Double)
        Private Declare Sub ENTROdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, s As Double)
        Private Declare Sub CVCPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Cv As Double, Cp As Double)
        Private Declare Sub GIBBSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Ar As Double, Gr As Double)
        Private Declare Sub AGdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, A As Double, g As Double)
        Private Declare Sub PHIXdll Lib "C:\REFPROP\REFPRP64.DLL" (itau As Long, idel As Long, tau As Double, del As Double, x As Double, PHIX As Double)
        Private Declare Sub RESIDUALdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, Pr As Double, er As Double, hr As Double, sr As Double, cvr As Double, cpr As Double, Ar As Double, Gr As Double)
        Private Declare Sub CP0dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, Cp As Double)
        Private Declare Sub DPDDdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dPdD As Double)
        Private Declare Sub DPDD2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, d2PdD2 As Double)
        Private Declare Sub DPDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dPdT As Double)
        Private Declare Sub DDDPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dDdP As Double)
        Private Declare Sub DDDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dDdT As Double)
        Private Declare Sub DHD1dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, dHdT_D As Double, dHdT_P As Double, dHdD_T As Double, dHdD_P As Double, dHdP_T As Double, dHdP_D As Double)
        Private Declare Sub VIRBdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare Sub VIRCdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, c As Double)
        Private Declare Sub VIRBAdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare Sub VIRCAdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, c As Double)
        Private Declare Sub DBDTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, dbt As Double)
        Private Declare Sub B12dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, b As Double)
        Private Declare Sub FGCTYdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, F As Double)
        Private Declare Sub FGCTY2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, F As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub FUGCOFdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, F As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CHEMPOTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, u As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ACTVYdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, actv As Double, gamma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub EXCESSdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, kph As Long, d As Double, vE As Double, ee As Double, hE As Double, sE As Double, aE As Double, gE As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATSPLNdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATTPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, kph As Long, iGuess As Long, d As Double, rhol As Double, rhov As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, i As Long, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, i As Long, T As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATDdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, x As Double, kph As Long, kr As Long, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATHdll Lib "C:\REFPROP\REFPRP64.DLL" (h As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATEdll Lib "C:\REFPROP\REFPRP64.DLL" (e As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SATSdll Lib "C:\REFPROP\REFPRP64.DLL" (s As Double, x As Double, kph As Long, nroot As Long, k1 As Long, t1 As Double, p1 As Double, d1 As Double, k2 As Long, t2 As Double, p2 As Double, d2 As Double, k3 As Long, t3 As Double, p3 As Double, d3 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CV2PKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, cv2p As Double, Csat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, kph As Long, P As Double, rho As Double, Csat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DPTSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, kph As Long, P As Double, rho As Double, Csat As Double, dPdTsat As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DLSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DVSATKdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TSATDdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MELTTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub MLTH2Odll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, p1 As Double, p2 As Double)
        Private Declare Sub MELTPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SUBLTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, x As Double, P As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SUBLPdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TPRHOdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, j As Long, i As Long, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TPFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TDFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PDFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, d As Double, x As Double, T As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PHFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, h As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, s As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, e As Double, x As Double, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub THFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, h As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, s As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, e As Double, x As Double, i As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DHFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, h As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, s As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DEFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, e As Double, x As Double, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub HSFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (h As Double, s As Double, z As Double, T As Double, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, e As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ESFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (e As Double, s As Double, z As Double, T As Double, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, h As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TQFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, q As Double, x As Double, kq As Long, P As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PQFLSHdll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, q As Double, x As Double, kq As Long, T As Double, d As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, e As Double, h As Double, s As Double, Cv As Double, Cp As Double, w As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ABFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (A As Double, b As Double, x As Double, i As Long, ByVal ab As String, dmin As Double, dmax As Double, T As Double, P As Double, d As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub ABFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (A As Double, b As Double, x As Double, kq As Long, ksat As Long, ByVal ab As String, tbub As Double, tdew As Double, pbub As Double, pdew As Double, Dlbub As Double, Dvdew As Double, ybub As Double, xdew As Double, T As Double, P As Double, Dl As Double, Dv As Double, x As Double, y As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long, ln2 As Long)
        Private Declare Sub DBFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, b As Double, x As Double, ByVal ab As String, T As Double, P As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub DBFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, b As Double, x As Double, i As Long, ByVal ab As String, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long, ln2 As Long)
        Private Declare Sub DQFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, q As Double, x As Double, kq As Long, T As Double, P As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DSFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (d As Double, s As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PDFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, d As Double, x As Double, T As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PHFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, h As Double, x As Double, kph As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub PSFL1dll Lib "C:\REFPROP\REFPRP64.DLL" (P As Double, s As Double, x As Double, kph As Long, T As Double, d As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TPFL2dll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, x As Double, Dl As Double, Dv As Double, xliq As Double, xvap As Double, q As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TRNPRPdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, eta As Double, tcx As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub DIELECdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, de As Double)
        Private Declare Sub SURFTdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, x As Double, sigma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub SURTENdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, rhol As Double, rhov As Double, xl As Double, xv As Double, sigma As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub HEATdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, rho As Double, x As Double, hg As Double, hn As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub CSTARdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, P As Double, v As Double, x As Double, cs As Double, ts As Double, Ds As Double, ps As Double, ws As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub FPVdll Lib "C:\REFPROP\REFPRP64.DLL" (T As Double, d As Double, P As Double, x As Double, F As Double)
        'Private Declare  Sub SPECGRdll Lib "C:\REFPROP\REFPRP64.DLL" (t As Double, d As Double, p As Double, Gr As Double)
        Private Declare Sub WMOLdll Lib "C:\REFPROP\REFPRP64.DLL" (x As Double, wm As Double)
        Private Declare Sub XMASSdll Lib "C:\REFPROP\REFPRP64.DLL" (xmol As Double, xkg As Double, wmix As Double)
        Private Declare Sub XMOLEdll Lib "C:\REFPROP\REFPRP64.DLL" (xkg As Double, xmol As Double, wmix As Double)
        Private Declare Sub QMASSdll Lib "C:\REFPROP\REFPRP64.DLL" (qmol As Double, xl As Double, xv As Double, qkg As Double, xlkg As Double, xvkg As Double, wliq As Double, wvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub QMOLEdll Lib "C:\REFPROP\REFPRP64.DLL" (qkg As Double, xlkg As Double, xvkg As Double, qmol As Double, xl As Double, xv As Double, wliq As Double, wvap As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub INFOdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, wmm As Double, ttrp As Double, tnbpt As Double, tc As Double, pc As Double, dc As Double, Zc As Double, acf As Double, dip As Double, Rgas As Double)
        Private Declare Sub LIMITSdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, x As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ln As Long)
        Private Declare Sub LIMITXdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, T As Double, d As Double, P As Double, x As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub LIMITKdll Lib "C:\REFPROP\REFPRP64.DLL" (ByVal htyp As String, icomp As Long, T As Double, d As Double, P As Double, tmin As Double, tmax As Double, dmax As Double, pmax As Double, ierr As Long, ByVal herr As String, ln1 As Long, ln2 As Long)
        Private Declare Sub ETAK0dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, eta0 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ETAK1dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, eta1 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ETAKRdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, etar As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub ETAKBdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, etab As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub OMEGAdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, epsk As Double, omg As Double)
        Private Declare Sub TCXK0dll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, tcx0 As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TCXKBdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, tcxb As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Sub TCXKCdll Lib "C:\REFPROP\REFPRP64.DLL" (icomp As Long, T As Double, rho As Double, tcxc As Double, ierr As Long, ByVal herr As String, ln As Long)
        Private Declare Function SetCurDir Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
#End If

'Used to call REFPROP:
Private herr As String * 255, herr2 As String * 255, hfmix As String * 255, hfmix2 As String * 255, hrf As String * 3, htyp As String * 3, hmxnme As String * 255
Private hfld As String * 10000, hfldOld As String * 10000
Private nc As Long, phase As Long, molmass As Long
Private x(1 To MaxComps) As Double, xliq(1 To MaxComps) As Double, xvap(1 To MaxComps) As Double, xmm(1 To MaxComps) As Double, xkg(1 To MaxComps) As Double, xmol(1 To MaxComps) As Double, wmix As Double
Private ierr As Long, ierr2 As Long, kq As Long, kr As Long
Private T As Double, P As Double, d As Double, Dl As Double, Dv As Double, q As Double, wm As Double, tz As Double, pz As Double, dz As Double, dd As Double
Private e As Double, h As Double, s As Double, Cvcalc As Double, Cpcalc As Double, W As Double
Private tmin As Double, tmax As Double, dmax As Double, pmax As Double
Private tc As Double, pc As Double, dc As Double
Private tbub As Double, tdew As Double, pbub As Double, pdew As Double, Dlbub As Double, Dvdew As Double, ybub(1 To MaxComps) As Double, xdew(1 To MaxComps) As Double
Private eta As Double, tcx As Double, sigma As Double, hjt As Double, de As Double
Private wmm As Double, ttrp As Double, tnbpt As Double, Zc As Double, acf As Double, dip As Double, Rgas As Double
Private tUnits(10) As String, taUnits(10) As String, pUnits(10) As String, dUnits(10) As String, vUnits(10) As String, hUnits(10) As String, sUnits(10) As String, wUnits(10) As String, visUnits(10) As String, tcxUnits(10) As String, stUnits(10) As String, tmUnits(10) As String
Private tUnits2 As String, taUnits2 As String, pUnits2 As String, dUnits2 As String, vUnits2 As String, hUnits2 As String, sUnits2 As String, wUnits2 As String, visUnits2 As String, tcxUnits2 As String, stUnits2 As String, tmUnits2 As String, DefUnit As Integer, DefaultUnits As Integer
Private FldOld As String
Private z As Double, aHelm As Double, Gibbs As Double, xkappa As Double, beta As Double
Private dPdD_T As Double, d2PdD2_rho As Double, dPdT_rho As Double, dDdT_P As Double, dDdP_rho As Double
Private d2PT2 As Double, d2PdTD As Double, spare3 As Double, spare4 As Double
Private isenk As Double, kt As Double, betas As Double, bs As Double, kkt As Double, thrott As Double, xpi As Double, spht As Double
Private nroot As Long, k1 As Long, k2 As Long, k3 As Long, t2 As Double, p2 As Double, d2 As Double, t3 As Double, p3 As Double, d3 As Double
Private eta0 As Double, eta1 As Double, etar As Double, etab As Double, epsk As Double, tcx0 As Double, tcxb As Double, tcxc As Double, omg As Double

Private Const Ridgas = 8.3144621
Private Const CtoK As Double = 273.15                  'Exact conversion
Private Const FtoR As Double = 459.67                  'Exact conversion
Private Const RtoK As Double = 5 / 9                   'Exact conversion
Private Const HtoS As Double = 3600                    'Exact conversion
Private Const ATMtoMPa As Double = 0.101325            'Exact conversion
Private Const BARtoMPA As Double = 0.1                 'Exact conversion
Private Const KGFtoN As Double = 98.0665 / 10          'Exact conversion
Private Const INtoM As Double = 0.0254                 'Exact conversion
Private Const FTtoM As Double = 12 * INtoM             'Exact conversion
Private Const LBMtoKG As Double = 0.45359237           'Exact conversion
Private Const CALtoJ As Double = 4.184                 'Exact conversion (tc)
'private Const CALtoJ As Double = 4.1868                'Exact conversion (IT) (Use this one only with pure water)
Private Const MMHGtoMPA As Double = ATMtoMPa / 760     'Exact conversion
Private Const INH2OtoMPA As Double = 0.000249082

Private Const BTUtoKJ As Double = CALtoJ * LBMtoKG * RtoK
Private Const LBFtoN As Double = LBMtoKG * KGFtoN
Private Const IN3toM3 As Double = INtoM * INtoM * INtoM
Private Const FT3toM3 As Double = FTtoM * FTtoM * FTtoM
Private Const GALLONtoM3 As Double = IN3toM3 * 231
Private Const PSIAtoMPA As Double = LBMtoKG / INtoM / INtoM * KGFtoN / 1000000
Private Const FTLBFtoJ As Double = FTtoM * LBFtoN
Private Const HPtoW As Double = 550 * FTLBFtoJ
Private Const BTUtoW As Double = BTUtoKJ * 1000
Private Const LBFTtoNM As Double = LBFtoN / FTtoM
Private CompFlag As Integer

Function Setup(FluidName)
  Dim a As String, ab As String, FluidNme As String, FlNme As String
  Dim i As Integer, sum As Double, sc As Integer, ncc As Integer, nc2 As Long
  Dim Mass As Integer, MixType As Integer
  Dim hRef As Double, sRef As Double, Tref As Double, pref As Double
  Dim htype As String * 3, hmix As String * 3, hcomp As String * 60
  Dim RPPrefix As String, FluidsPrefix As String, MixturesPrefix As String
  Dim xtemp(1 To MaxComps) As Double

  ierr = 0
  herr = ""
  FlNme = FluidName
  If InStr(FluidName, "error") Then Exit Function
  If InStr(FluidName, "Inputs are out of range") Then Exit Function
  If FluidName = FldOld Then Exit Function
  FldOld = ""
  Call CheckName(FluidName)

  RPPrefix = Trim(Environ("RPPrefix"))
  If Right(RPPrefix, 1) = "\" Then RPPrefix = Left(RPPrefix, Len(RPPrefix) - 1)
  If RPPrefix = "" Then
    FluidsPrefix = FluidsDirectory
    MixturesPrefix = MixturesDirectory
  Else
    FluidsPrefix = RPPrefix & "\" & FluidsDirectory
    MixturesPrefix = RPPrefix & "\" & MixturesDirectory
  End If

  hrf = "DEF"
  hfmix = FluidsPrefix & "hmx.bnc"

  On Error GoTo ErrorHandler:
  ChDrive (ThisWorkbook.Path)
  ChDir (ThisWorkbook.Path)
  'ChDrive (Application.ActiveWorkbook.Path)
  'ChDir (Application.ActiveWorkbook.Path)
  On Error GoTo 0

'Use this for UNC drives
' ChDirUNC ThisWorkbook.Path

  a = ""
  For i = 1 To MaxComps: xtemp(i) = 0: Next
  Mass = 0
  MixType = 0
  If InStr(FluidName, "/") <> 0 And InStr(FluidName, "(") <> 0 Then MixType = 2
  If InStr(FluidName, ",") <> 0 Or InStr(FluidName, ";") <> 0 Then MixType = 1

'This section added by J. Newland 9JUL2010 (15MAY2012) to recognize mixed refrigerants
  If (Left(UCase(FluidName), 2) = "R4" Or Left(UCase(FluidName), 2) = "R5") And InStr(UCase(FluidName), ".MIX") = 0 And MixType = 0 Then
    If UCase(FluidName) <> "R40" And UCase(FluidName) <> "R41" Then FluidName = FluidName & ".MIX"
  End If

'Set up predefined mixtures (using *.MIX)
  If InStr(UCase(FluidName), ".MIX") Then
    hmxnme = MixturesPrefix & FluidName
    Call SETMIXdll(hmxnme, hfmix, hrf, nc2, hfld, xtemp(1), ierr, herr, 255&, 255&, 3&, 10000&, 255&)

  ElseIf MixType Then
    FluidNme = Trim(FluidName)
    If UCase(Right(FluidNme, 4)) = "MASS" Then Mass = 1: FluidNme = Trim(Left(FluidNme, Len(FluidNme) - 4))
    nc2 = 0
    If MixType = 1 Then
'Set up user specified mixtures where fluids are separated by "," or ";"
      If InStr(FluidNme, ";") Then sc = 1 Else sc = 0
      Do
        If sc = 0 Then i = InStr(FluidNme, ",") Else i = InStr(FluidNme, ";")
        If i = 0 Then i = Len(FluidNme) + 1
        nc2 = nc2 + 1
        If nc2 > MaxComps Then ierr = 1: herr = Trim2("Too many components"): Exit Function
        ab = Trim(Left(FluidNme, i - 1))
        Call CheckName(ab)
        If InStr(LCase(ab), ".fld") = 0 Then ab = ab + ".fld"
        a = a & FluidsPrefix & ab & "|"
        FluidNme = Mid(FluidNme, i + 1)
        If sc = 0 Then i = InStr(FluidNme, ",") Else i = InStr(FluidNme, ";")
        If i = 0 Then i = Len(FluidNme) + 1
        xtemp(nc2) = CDbl(Left(FluidNme, i - 1))
        FluidNme = Trim(Mid(FluidNme, i + 1))
      Loop Until FluidNme = ""
    ElseIf MixType = 2 Then
'Set up user specified mixtures with "/" as the separator
      Do
        i = InStr(FluidNme, "/")
        If InStr(FluidNme, "(") < i Then i = InStr(FluidNme, "(")
        If i = 0 Then i = Len(FluidNme) + 1
        nc2 = nc2 + 1
        If nc2 > MaxComps Then ierr = 1: herr = Trim2("Too many components"): Exit Function
        ab = Trim(Left(FluidNme, i - 1))
        Call CheckName(ab)
        If InStr(LCase(ab), ".fld") = 0 Then ab = ab + ".fld"
        a = a & FluidsPrefix & ab & "|"
        FluidNme = Trim(Mid(FluidNme, i))
        If Left(FluidNme, 1) = "/" Then FluidNme = Trim(Mid(FluidNme, 2))
      Loop Until Left(FluidNme, 1) = "("
      FluidNme = Mid(FluidNme, 2)
      If Right(FluidNme, 1) = ")" Then FluidNme = Trim(Left(FluidNme, Len(FluidNme) - 1))
      ncc = 0
      Do
        i = InStr(FluidNme, "/")
        If i = 0 Then i = Len(FluidNme) + 1
        ncc = ncc + 1
        If ncc > MaxComps Then ierr = 1: herr = Trim2("Too many components"): Exit Function
        xtemp(ncc) = CDbl(Left(FluidNme, i - 1))
        FluidNme = Mid(FluidNme, i + 1)
      Loop Until FluidNme = ""
    End If

'Common code for mixtures
    sum = 0
    For i = 1 To nc2: sum = sum + xtemp(i): Next
    If sum <= 0 Then ierr = 1: herr = Trim2("Composition not set"): Exit Function
    For i = 1 To nc2: xtemp(i) = xtemp(i) / sum: Next
    hfld = a
    If nc2 < 1 Then ierr = 1: herr = Trim2("Setup failed"): Exit Function
    If hfld <> hfldOld Then
      'To load the GERG-2004 pure fluid equations of state rather than the defaults
      'that come with REFPROP, call the GERG04dll routine with a 1 as the second input.
      'Call GERG04dll(nc2, 1&, ierr, herr, 255&)
      Call SETUPdll(nc2, hfld, hfmix, hrf, ierr, herr, 10000&, 255&, 3&, 255&)
      If ierr > 0 Then hfld = ""
      'To load the AGA8 equation of state, call SETAGA after calling SETUP
      'Call SETAGAdll(ierr, herr, 255&)
    End If

'Set up pure fluids
  Else
    nc2 = 1
    If InStr(LCase(FluidName), ".fld") = 0 And InStr(LCase(FluidName), ".ppf") = 0 Then FluidName = FluidName + ".fld"
    If InStr(FluidName, "\") Then
      hfld = FluidName
    Else
      hfld = FluidsPrefix & FluidName
    End If
    '...Use call to SETMOD to change the equation of state for any of the
    '.....pure components from the default (recommended) values.
    '.....This should only be implemented by an experienced user.
    'If InStr(LCase(hfld), "argon") <> 0 And nc2 = 1 Then
    '  hcomp = "FE1": htype = "EOS": hmix = hcomp
    '  Call SETMODdll(nc2, htype, hmix, hcomp, ierr, herr, 3&, 3&, 60&, 255&)
    'End If
    Call SETUPdll(nc2, hfld, hfmix, hrf, ierr, herr, 10000&, 255&, 3&, 255&)
    If ierr > 0 Then hfld = ""
  End If
  hfldOld = hfld


  'Use call to PREOSdll to change the equation of state to Peng Robinson for all calculations.
  'To revert back to the normal REFPROP EOS and models, use:  Call PREOSdll(0&)
  'Call PREOSdll(2&)

  If Mass Then
    For i = 1 To nc2
      xkg(i) = xtemp(i)
    Next
    Call XMOLEdll(xkg(1), xtemp(1), wmix)
  End If

  If ierr <= 0 Then
    nc = nc2           'If setup was successful, load new values of nc and x()
    For i = 1 To nc
      x(i) = xtemp(i)
    Next
    Setup = FluidName
    FldOld = FlNme
    'Use the following line to activate the call to SATSPLN, which will set up
    'spline curves that represent the saturation states of the mixture, allowing Refprop
    'to know the critical point, the saturation state with the maximum temperature, and
    'the saturation state with the maximum pressure (all of these for the composition given
    'in x).  However, this call is slow and is best used in a separate xls file for a
    'dedicated mixture.
    'Call SATSPLNdll(x(1), ierr, herr, 255)

    'Use the following line to calculate enthalpies and entropies on a reference state
    'based on the currently defined mixture, or to change to some other reference state.
    'The routine does not have to be called, but doing so will cause calculations
    'to be the same as those produced from the graphical interface for mixtures.
    Call SETREFdll(hrf, 2&, x(1), hRef, sRef, Tref, pref, ierr2, herr2, 3&, 255&)
  Else
    Setup = Trim2(herr)
    FldOld = ""
  End If
  Exit Function

ErrorHandler:
  Resume Next
End Function

Sub CheckName(FluidName)
  Dim i As Integer
  FluidName = UCase(FluidName)
  While Left(FluidName, 1) = Chr(34)
    FluidName = Mid(FluidName, 2)
  Wend
  While Right(FluidName, 1) = Chr(34)
    FluidName = Left(FluidName, Len(FluidName) - 1)
  Wend
  Do
    i = InStr(FluidName, " ")
    If i Then FluidName = Left(FluidName, i - 1) + Mid(FluidName, i + 1)
  Loop While i
  Do
    i = InStr(FluidName, "-")
    If i Then FluidName = Left(FluidName, i - 1) + Mid(FluidName, i + 1)
  Loop While i

  If FluidName = "AIR" Then FluidName = "nitrogen;7812;argon;0092;oxygen;2096"
  If FluidName = "BUTENE" Then FluidName = "1BUTENE"
  If FluidName = "CARBONDIOXIDE" Then FluidName = "CO2"
  If FluidName = "CARBONMONOXIDE" Then FluidName = "CO"
  If FluidName = "CARBONYLSULFIDE" Then FluidName = "COS"
  If FluidName = "CIS-BUTENE" Then FluidName = "C2BUTENE"
  If FluidName = "CYCLOHEXANE" Then FluidName = "CYCLOHEX"
  If FluidName = "CYCLOPENTANE" Then FluidName = "CYCLOPEN"
  If FluidName = "CYCLOPROPANE" Then FluidName = "CYCLOPRO"
  If FluidName = "DEUTERIUM" Then FluidName = "D2"
  If FluidName = "DIMETHYLCARBONATE" Then FluidName = "DMC"
  If FluidName = "DIMETHYLETHER" Then FluidName = "DME"
  If FluidName = "DIETHYLETHER" Then FluidName = "DEE"
  If FluidName = "DODECANE" Then FluidName = "C12"
  If FluidName = "ETHYLBENZENE" Then FluidName = "EBENZENE"
  If FluidName = "HEAVYWATER" Then FluidName = "D2O"
  If FluidName = "HYDROGENCHLORIDE" Then FluidName = "HCL"
  If FluidName = "HYDROGENSULFIDE" Then FluidName = "H2S"
  If FluidName = "IBUTANE" Then FluidName = "ISOBUTAN"
  If FluidName = "ISOBUTANE" Then FluidName = "ISOBUTAN"
  If FluidName = "ISOBUTENE" Then FluidName = "IBUTENE"
  If FluidName = "ISOHEXANE" Then FluidName = "IHEXANE"
  If FluidName = "ISOPENTANE" Then FluidName = "IPENTANE"
  If FluidName = "ISOOCTANE" Then FluidName = "IOCTANE"
  If FluidName = "METHYLCYCLOHEXANE" Then FluidName = "C1CC6"
  If FluidName = "METHYLLINOLENATE" Then FluidName = "MLINOLEN"
  If FluidName = "METHYLLINOLEATE" Then FluidName = "MLINOLEA"
  If FluidName = "METHYLOLEATE" Then FluidName = "MOLEATE"
  If FluidName = "METHYLPALMITATE" Then FluidName = "MPALMITA"
  If FluidName = "METHYLSTEARATE" Then FluidName = "MSTEARAT"
  If FluidName = "NEOPENTANE" Then FluidName = "NEOPENTN"
  If FluidName = "NITROGENTRIFLUORIDE" Then FluidName = "NF3"
  If FluidName = "NITROUSOXIDE" Then FluidName = "N2O"
  If FluidName = "ORTHOHYDROGEN" Then FluidName = "ORTHOHYD"
  If FluidName = "PARAHYDROGEN" Then FluidName = "PARAHYD"
  If FluidName = "PERFLUOROBUTANE" Then FluidName = "C4F10"
  If FluidName = "PERFLUOROPENTANE" Then FluidName = "C5F12"
  If FluidName = "PROPYLCYCLOHEXANE" Then FluidName = "C3CC6"
  If FluidName = "PROPYLENE" Then FluidName = "PROPYLEN"
  If FluidName = "SULFUR HEXAFLUORIDE" Then FluidName = "SF6"
  If FluidName = "TRANS-BUTENE" Then FluidName = "T2BUTENE"
  If FluidName = "TRIFLUOROIODOMETHANE" Then FluidName = "CF3I"
  If FluidName = "SULFURDIOXIDE" Then FluidName = "SO2"
  If FluidName = "SULFURHEXAFLUORIDE" Then FluidName = "SF6"
  If FluidName = "UNDECANE" Then FluidName = "C11"
End Sub

Sub CalcSetup(FluidName, InpCode, Units, Prop1, Prop2)
  If Trim(FluidName) = "" Then herr = Trim2("Invalid inputs"): Exit Sub
  Call Setup(FluidName)
  If ierr > 0 Then Exit Sub
  Call ConvertUnits(InpCode, Units, Prop1, Prop2)
  If ierr = 1 Then herr = Trim2(herr): Exit Sub
  herr = ""
  q = 0: T = 0: P = 0: d = 0: Dl = 0: Dv = 0: e = 0: h = 0: s = 0: Cvcalc = 0: Cpcalc = 0: W = 0
End Sub

Sub CalcProp(FluidName, InpCode, ByVal Units, ByVal Prop1, ByVal Prop2)
  Dim iflag1 As Integer, iflag2 As Integer
  ThisWorkbook.Activate
  q = 0: T = 0: P = 0: d = 0: Dl = 0: Dv = 0: e = 0: h = 0: s = 0: Cvcalc = 0: Cpcalc = 0: W = 0

  If IsMissing(Prop1) Then iflag1 = 1
  If iflag1 = 0 Then
    If Len(Trim(Prop1)) = 0 Then iflag1 = 2
    If iflag1 = 0 Then If CDbl(Prop1) = 0 And Prop1 <> "0" Then ierr = 1: herr = Trim2("Invalid input: ") + Prop1: Exit Sub
  End If

  If IsMissing(Prop2) Then iflag2 = 1
  If iflag2 = 0 Then
    If Len(Trim(Prop2)) = 0 Then iflag2 = 2
    If iflag2 = 0 Then If CDbl(Prop2) = 0 And Prop2 <> "0" Then ierr = 1: herr = Trim2("Invalid input: ") + Prop2: Exit Sub
  End If

  If IsMissing(InpCode) Then InpCode = ""
  Call CalcSetup(FluidName, InpCode, Units, Prop1, Prop2)
  If UCase(Left(InpCode, 4)) = "CRIT" Then
    Call CRITPdll(x(1), T, P, d, ierr, herr, 255&)
    If ierr = 0 Then Call THERMdll(T, d, x(1), pc, e, h, s, Cvcalc, Cpcalc, W, hjt)
    Exit Sub
  ElseIf UCase(Left(InpCode, 7)) = "SATMAXT" Then
    Call MAXTdll(x(1), T, P, d, ierr, herr, 255&)
    If ierr = 0 Then Call THERMdll(T, d, x(1), pc, e, h, s, Cvcalc, Cpcalc, W, hjt)
    Exit Sub
  ElseIf UCase(Left(InpCode, 7)) = "SATMAXP" Then
    Call MAXPdll(x(1), T, P, d, ierr, herr, 255&)
    If ierr = 0 Then Call THERMdll(T, d, x(1), pc, e, h, s, Cvcalc, Cpcalc, W, hjt)
    Exit Sub
  ElseIf UCase(Left(InpCode, 4)) = "TRIP" Then
    If nc <> 1 Then ierr = 1: herr = Trim2("Can only return triple point for a pure fluid"): Exit Sub
    Call INFOdll(1, wmm, T, tnbpt, tc, pc, dc, Zc, acf, dip, Rgas)
    Call SATTdll(T, x(1), 1, P, d, Dv, xliq(1), xvap(1), ierr, herr, 255&)
    If ierr = 0 Then Call THERMdll(T, d, x(1), pc, e, h, s, Cvcalc, Cpcalc, W, hjt)
    Exit Sub
  End If

  If iflag1 Then ierr = 1: herr = Trim2("Inputs are missing"): Exit Sub
  If ierr > 0 Then Exit Sub
  If InpCode <> "" Then Call Calc(InpCode, Prop1, Prop2, iflag1, iflag2)
End Sub

Sub Calc(InputCode, Prop1, Prop2, iflag1, iflag2)
  Dim a As String, Input1 As String, Input2 As String, InpCode, i As Integer, pp As Double
  Dim xlkg(1 To MaxComps) As Double, xvkg(1 To MaxComps) As Double, xlj(1 To MaxComps) As Double, xvj(1 To MaxComps) As Double, qmol As Double, wliq As Double, wvap As Double
  ierr = 0
  herr = ""
  InpCode = Trim(UCase(InputCode))
  Input2 = ""
  Input1 = Left(InpCode, 1)
  If Len(InpCode) = 2 Then Input2 = Mid(InpCode, 2, 1)
  If Len(InpCode) = 3 Then
    a = Right(InpCode, 1)
    If a = "&" Or a = "<" Or a = ">" Then Input2 = Mid(InpCode, 2, 1)
  End If
  If Left(InpCode, 2) = "TP" Or Left(InpCode, 2) = "PT" Then Input2 = Mid(InpCode, 2, 1)

  If Input1 = "T" Then T = Prop1: If iflag1 >= 1 Then GoTo Error1
  If Input1 = "P" Then P = Prop1: If iflag1 >= 1 Then GoTo Error1
  If Input1 = "D" Then d = Prop1: If iflag1 >= 1 Then GoTo Error1
  If Input1 = "V" And Prop1 <> 0 And Len(InpCode) = 2 Then d = 1 / Prop1: Mid(InpCode, 1, 1) = "D": If iflag1 >= 1 Then GoTo Error1
  If Input1 = "E" Then e = Prop1: If iflag1 >= 1 Then GoTo Error1
  If Input1 = "H" Then h = Prop1: If iflag1 >= 1 Then GoTo Error1
  If Input1 = "S" Then s = Prop1: If iflag1 >= 1 Then GoTo Error1
  If Input1 = "Q" Then q = Prop1: If iflag1 >= 1 Then GoTo Error1
  If Input2 = "T" Then T = Prop2: If iflag2 >= 1 Then GoTo Error2
  If Input2 = "P" Then P = Prop2: If iflag2 >= 1 Then GoTo Error2
  If Input2 = "D" Then d = Prop2: If iflag2 >= 1 Then GoTo Error2
  If Input2 = "V" And Prop2 <> 0 And Len(InpCode) = 2 Then d = 1 / Prop2: Mid(InpCode, 2, 1) = "D": If iflag2 >= 1 Then GoTo Error2
  If Input2 = "E" Then e = Prop2: If iflag2 >= 1 Then GoTo Error2
  If Input2 = "H" Then h = Prop2: If iflag2 >= 1 Then GoTo Error2
  If Input2 = "S" Then s = Prop2: If iflag2 >= 1 Then GoTo Error2
  If Input2 = "Q" Then q = Prop2: If iflag2 >= 1 Then GoTo Error2

  phase = 2
  If Len(InpCode) > 1 Then If UCase(Mid(InpCode, 2, 1)) = "L" Then phase = 1

  For i = 1 To nc
    xliq(i) = 0: xvap(i) = 0
  Next
  If Left(InpCode, 1) = "T" And T <= 0 Then herr = Trim2("Input temperature is zero"): Exit Sub
  'Calculate saturation values given temperature
  If InpCode = "TL" Or InpCode = "TLIQ" Or InpCode = "TVAP" Then
    Call SATTdll(T, x(1), phase, P, Dl, Dv, xliq(1), xvap(1), ierr, herr, 255&)
    If (P = 0 Or Dl = 0) And ierr = 0 Then ierr = 1: herr = Trim2("Inputs are out of range"): Exit Sub
    d = Dl: q = 0
    If phase = 2 Then d = Dv: q = 1
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
  'Calculate saturation values given pressure
  ElseIf InpCode = "PL" Or InpCode = "PLIQ" Or InpCode = "PVAP" Then
    Call SATPdll(P, x(1), phase, T, Dl, Dv, xliq(1), xvap(1), ierr, herr, 255&)
    If (P = 0 Or Dl = 0) And ierr = 0 Then ierr = 1: herr = Trim2("Inputs are out of range"): Exit Sub
    d = Dl: q = 0
    If phase = 2 Then d = Dv: q = 1
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
  'Calculate saturation values given density
  ElseIf InpCode = "DSAT" Or InpCode = "DL" Or InpCode = "DLIQ" Or InpCode = "DVAP" Then
    Call SATDdll(d, x(1), 1&, kr, T, P, Dl, Dv, xliq(1), xvap(1), ierr, herr, 255&)
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    q = kr - 1
  'Calculate properties given the enthalpy along the saturation line
  ElseIf InpCode = "HSAT" Or InpCode = "HSAT1" Then
    Call SATHdll(h, x(1), 0&, nroot, k1, T, P, d, k2, t2, p2, d2, ierr, herr, 255&)
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    q = k1 - 1
    If nroot = 0 Then ierr = 1: herr = Trim2("Inputs are out of range"): Exit Sub
  'Calculate properties given the enthalpy along the saturation line for the second possible root
  ElseIf InpCode = "HSAT2" Then
    Call SATHdll(h, x(1), 0&, nroot, k2, t2, p2, d2, k1, T, P, d, ierr, herr, 255&)
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    q = k1 - 1
    If nroot < 2 Then ierr = 1: herr = Trim2("Inputs are out of range or there is no second root"): Exit Sub
  'Calculate properties given the entropy along the saturation line
  ElseIf InpCode = "SSAT" Or InpCode = "SSAT1" Then
    Call SATSdll(s, x(1), 0&, nroot, k1, T, P, d, k2, t2, p2, d2, k3, t3, p3, d3, ierr, herr, 255&)
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    q = k1 - 1
    If nroot = 0 Then ierr = 1: herr = Trim2("Inputs are out of range"): Exit Sub
  'Calculate properties given the entropy along the saturation line for the second possible root
  ElseIf InpCode = "SSAT2" Then
    Call SATSdll(s, x(1), 0&, nroot, k2, t2, p2, d2, k1, T, P, d, k3, t3, p3, d3, ierr, herr, 255&)
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    q = k1 - 1
    If nroot < 2 Then ierr = 1: herr = Trim2("Inputs are out of range or there is no second root"): Exit Sub
  'Calculate properties given the entropy along the saturation line for the third possible root
  ElseIf InpCode = "SSAT3" Then
    Call SATSdll(s, x(1), 0&, nroot, k3, t3, p3, d3, k2, t2, p2, d2, k1, T, P, d, ierr, herr, 255&)
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    q = k1 - 1
    If nroot < 3 Then ierr = 1: herr = Trim2("Inputs are out of range or there is no third root"): Exit Sub
  ElseIf InpCode = "TPL" Or InpCode = "PTL" Then
    Call TPRHOdll(T, P, x(1), 1&, 0&, d, ierr, herr, 255&)
    Dl = d: Dv = d: q = 990
    Call THERMdll(T, d, x(1), pp, e, h, s, Cvcalc, Cpcalc, W, hjt)
  ElseIf InpCode = "TPV" Or InpCode = "PTV" Then
    Call TPRHOdll(T, P, x(1), 2&, 0&, d, ierr, herr, 255&)
    Dl = d: Dv = d: q = 990
    Call THERMdll(T, d, x(1), pp, e, h, s, Cvcalc, Cpcalc, W, hjt)
  ElseIf InpCode = "TP" Or InpCode = "PT" Then
    Call TPFLSHdll(T, P, x(1), d, Dl, Dv, xliq(1), xvap(1), q, e, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
    If ierr = -16 And nc = 1 Then
      herr2 = herr
      'Point is below triple point, call SATT to see if a sublimation line is available
      'and to check if the input p is less than the sublimation pressure.
      Call SATTdll(T, x(1), 2, pp, Dl, Dv, xliq(1), xvap(1), ierr, herr, 255&)
      If P < pp And ierr <= 0 Then
        Call TPRHOdll(T, P, x(1), 2&, 0&, d, ierr, herr, 255&)
      Else
        ierr = -16: herr = herr2
      End If
      q = 998
    End If
  ElseIf InpCode = "TD" Or InpCode = "DT" Then
    Call TDFLSHdll(T, d, x(1), P, Dl, Dv, xliq(1), xvap(1), q, e, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "TD&" Or InpCode = "DT&" Then
    'Do not perform any flash calculation here
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    q = 990
  ElseIf Left(InpCode, 2) = "TH" Or Left(InpCode, 2) = "HT" Then
    If InStr(InpCode, "<") Then
      Call THFLSHdll(T, h, x(1), 1&, P, d, Dl, Dv, xliq(1), xvap(1), q, e, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
    Else
      Call THFLSHdll(T, h, x(1), 2&, P, d, Dl, Dv, xliq(1), xvap(1), q, e, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
    End If
  ElseIf InpCode = "TS" Or InpCode = "ST" Then
    Call TSFLSHdll(T, s, x(1), 1&, P, d, Dl, Dv, xliq(1), xvap(1), q, e, h, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "TE" Or InpCode = "ET" Then
    Call TEFLSHdll(T, e, x(1), 2&, P, d, Dl, Dv, xliq(1), xvap(1), q, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "TQ" Or InpCode = "QT" Then
    Call TQFLSHdll(T, q, x(1), molmass, P, d, Dl, Dv, xliq(1), xvap(1), e, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
    If molmass = 2 Then
      Call XMASSdll(xliq(1), xlkg(1), wmix)
      Call XMASSdll(xvap(1), xvkg(1), wmix)
      Call QMOLEdll(q, xlkg(1), xvkg(1), qmol, xlj(1), xvj(1), wliq, wvap, ierr2, herr2, 255&)
      q = qmol
    End If
  ElseIf InpCode = "PD" Or InpCode = "DP" Then
    Call PDFLSHdll(P, d, x(1), T, Dl, Dv, xliq(1), xvap(1), q, e, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "PH" Or InpCode = "HP" Then
    Call PHFLSHdll(P, h, x(1), T, d, Dl, Dv, xliq(1), xvap(1), q, e, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "PS" Or InpCode = "SP" Then
    Call PSFLSHdll(P, s, x(1), T, d, Dl, Dv, xliq(1), xvap(1), q, e, h, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "PE" Or InpCode = "EP" Then
    Call PEFLSHdll(P, e, x(1), T, d, Dl, Dv, xliq(1), xvap(1), q, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "PQ" Or InpCode = "QP" Then
    Call PQFLSHdll(P, q, x(1), molmass, T, d, Dl, Dv, xliq(1), xvap(1), e, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
    If molmass = 2 Then
      Call XMASSdll(xliq(1), xlkg(1), wmix)
      Call XMASSdll(xvap(1), xvkg(1), wmix)
      Call QMOLEdll(q, xlkg(1), xvkg(1), qmol, xlj(1), xvj(1), wliq, wvap, ierr2, herr2, 255&)
      q = qmol
    End If
  ElseIf InpCode = "DH" Or InpCode = "HD" Then
    Call DHFLSHdll(d, h, x(1), T, P, Dl, Dv, xliq(1), xvap(1), q, e, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "DS" Or InpCode = "SD" Then
    Call DSFLSHdll(d, s, x(1), T, P, Dl, Dv, xliq(1), xvap(1), q, e, h, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "DE" Or InpCode = "ED" Then
    Call DEFLSHdll(d, e, x(1), T, P, Dl, Dv, xliq(1), xvap(1), q, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "HS" Or InpCode = "SH" Then
    Call HSFLSHdll(h, s, x(1), T, P, d, Dl, Dv, xliq(1), xvap(1), q, e, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "TMELT" Then
    Call MELTTdll(T, x(1), P, ierr, herr, 255&)
    If ierr = 0 Then Call TPFLSHdll(T, P, x(1), d, Dl, Dv, xliq(1), xvap(1), q, e, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "PMELT" Then
    If P = 0 Then ierr = 1: herr = Trim2("Input pressure is zero"): Exit Sub
    Call MELTPdll(P, x(1), T, ierr, herr, 255&)
    If ierr = 0 Then Call TPFLSHdll(T, P, x(1), d, Dl, Dv, xliq(1), xvap(1), q, e, h, s, Cvcalc, Cpcalc, W, ierr, herr, 255&)
  ElseIf InpCode = "TSUBL" Then
    Call SUBLTdll(T, x(1), P, ierr, herr, 255&)
    If ierr = 0 And P = 0 Then ierr = 1: herr = Trim2("No sublimation line available")
    If ierr = 0 Then
      q = 1
      d = P / Ridgas / T
      Call TPRHOdll(T, P, x(1), 2&, 1&, d, ierr, herr, 255&)
      Call THERMdll(T, d, x(1), pp, e, h, s, Cvcalc, Cpcalc, W, hjt)
    End If
  ElseIf InpCode = "PSUBL" Then
    If P = 0 Then ierr = 1: herr = Trim2("Input pressure is zero"): Exit Sub
    Call SUBLPdll(P, x(1), T, ierr, herr, 255&)
    If ierr = 0 And T = 0 Then ierr = 1: herr = Trim2("No sublimation line available")
    If ierr = 0 Then
      q = 1
      d = P / Ridgas / T
      Call TPRHOdll(T, P, x(1), 2&, 1&, d, ierr, herr, 255&)
      Call THERMdll(T, d, x(1), pp, e, h, s, Cvcalc, Cpcalc, W, hjt)
    End If
  Else
    ierr = 1: herr = Trim2("Invalid input code")
  End If
  'This line has been removed because there are valid q's less than 0.000001, is it still needed?
  'If (q <= 0.000001 Or q >= 0.999999) And Cvcalc = -9999980 Then Call THERMdll(t, d, x(1), pp, e, h, s, Cvcalc, Cpcalc, w, hjt)
  Exit Sub

Error1:
  ierr = 1: herr = Trim2("First property missing"): Exit Sub
Error2:
  ierr = 1: herr = Trim2("Second property missing"): Exit Sub

End Sub

Function Temperature(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Temperature = ConvertUnits("-T", Units, T, 0)
  If ierr > 0 And LCase(InpCode) = "trip" Then Temperature = T
End Function

Function Pressure(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Pressure = ConvertUnits("-P", Units, P, 0)
End Function

Function Density(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Density = ConvertUnits("-D", Units, d, 0)
End Function

Function CompressibilityFactor(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Call INFOdll(1, wmm, ttrp, tnbpt, tc, pc, dc, Zc, acf, dip, Rgas)
  If ierr > 0 Then CompressibilityFactor = Trim2(herr): Exit Function
  CompressibilityFactor = P / d / T / Rgas
End Function

Function LiquidDensity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then LiquidDensity = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    LiquidDensity = Trim2("Inputs are single phase")
  Else
    CompFlag = 1
    LiquidDensity = ConvertUnits("-D", Units, Dl, 0)
    CompFlag = 0
  End If
End Function

Function VaporDensity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then VaporDensity = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    VaporDensity = Trim2("Inputs are single phase")
  Else
    CompFlag = 2
    VaporDensity = ConvertUnits("-D", Units, Dv, 0)
    CompFlag = 0
  End If
End Function

Function LiquidEnthalpy(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then LiquidEnthalpy = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    'LiquidEnthalpy = Trim2("Inputs are single phase")
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    LiquidEnthalpy = ConvertUnits("-H", Units, h, 0)
  Else
    Call THERMdll(T, Dl, xliq(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    CompFlag = 1
    LiquidEnthalpy = ConvertUnits("-H", Units, h, 0)
    CompFlag = 0
  End If
End Function

Function VaporEnthalpy(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then VaporEnthalpy = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    'VaporEnthalpy = Trim2("Inputs are single phase")
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    VaporEnthalpy = ConvertUnits("-H", Units, h, 0)
  Else
    Call THERMdll(T, Dv, xvap(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    CompFlag = 2
    VaporEnthalpy = ConvertUnits("-H", Units, h, 0)
    CompFlag = 0
  End If
End Function

Function LiquidEntropy(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then LiquidEntropy = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    'LiquidEntropy = Trim2("Inputs are single phase")
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    LiquidEntropy = ConvertUnits("-S", Units, s, 0)
  Else
    Call THERMdll(T, Dl, xliq(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    CompFlag = 1
    LiquidEntropy = ConvertUnits("-S", Units, s, 0)
    CompFlag = 0
  End If
End Function

Function VaporEntropy(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then VaporEntropy = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    'VaporEntropy = Trim2("Inputs are single phase")
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    VaporEntropy = ConvertUnits("-S", Units, s, 0)
  Else
    Call THERMdll(T, Dv, xvap(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    CompFlag = 2
    VaporEntropy = ConvertUnits("-S", Units, s, 0)
    CompFlag = 0
  End If
End Function

Function LiquidCp(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then LiquidCp = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    LiquidCp = ConvertUnits("-S", Units, Cpcalc, 0)

  Else
    Call THERMdll(T, Dl, xliq(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    CompFlag = 1
    LiquidCp = ConvertUnits("-S", Units, Cpcalc, 0)
    CompFlag = 0
  End If
End Function

Function VaporCp(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then VaporCp = Trim2(herr): Exit Function
  If q < 0 Or q > 1 Then
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    VaporCp = ConvertUnits("-S", Units, Cpcalc, 0)
  Else
    Call THERMdll(T, Dv, xvap(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    CompFlag = 2
    VaporCp = ConvertUnits("-S", Units, Cpcalc, 0)
    CompFlag = 0
  End If
End Function

Function Volume(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim v As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Volume = 0
  If d <= 0 Then Volume = Trim2("Density is zero"): Exit Function
  v = 1 / d
  Volume = ConvertUnits("-V", Units, v, 0)
End Function

Function Energy(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Energy = ConvertUnits("-H", Units, e, 0)
End Function

Function Enthalpy(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Enthalpy = ConvertUnits("-H", Units, h, 0)
End Function

Function Entropy(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Entropy = ConvertUnits("-S", Units, s, 0)
End Function

Function IsochoricHeatCapacity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  IsochoricHeatCapacity = ConvertUnits("-S", Units, Cvcalc, 0)
End Function

Function Cv(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Cv = ConvertUnits("-S", Units, Cvcalc, 0)
End Function

Function IsobaricHeatCapacity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  IsobaricHeatCapacity = ConvertUnits("-S", Units, Cpcalc, 0)
End Function

Function Cp(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Cp = ConvertUnits("-S", Units, Cpcalc, 0)
End Function

Function Csat(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If nc <> 1 Then Csat = "Csat can only be called for a pure fluid": Exit Function
  Call CSATKdll(1, T, 1, P, d, Csat, ierr, herr, 255&)
  Csat = ConvertUnits("-S", Units, Csat, 0)
End Function

Function SpeedOfSound(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  SpeedOfSound = ConvertUnits("-W", Units, W, 0)
End Function

Function Sound(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Sound = ConvertUnits("-W", Units, W, 0)
End Function

Function LatentHeat(FluidName, InpCode, Optional Units, Optional ByVal Prop1, Optional ByVal Prop2)
  Dim hl As Double, hv As Double
  InpCode = Trim(UCase(InpCode))
  If Left(InpCode, 1) = "T" Then
    Call CalcSetup(FluidName, "T", Units, Prop1, Prop2)
    If ierr > 0 Then LatentHeat = Trim2(herr): Exit Function
    If nc <> 1 Then LatentHeat = Trim2("Can only be calculated for pure fluids"): Exit Function
    Call INFOdll(1, wmm, ttrp, tnbpt, tc, pc, dc, Zc, acf, dip, Rgas)
    T = Prop1
    If T <= 0 Then LatentHeat = Trim2("Input temperature is zero"): Exit Function
    If T > tc Then LatentHeat = Trim2("Temperature is greater than the critical point temperture"): Exit Function
    Call SATTdll(T, x(1), 1&, P, Dl, Dv, xliq(1), xvap(1), ierr, herr, 255&)
    If (P = 0 Or Dl = 0) And ierr = 0 Then ierr = 1: LatentHeat = Trim2("Inputs are out of range"): Exit Function
  ElseIf Left(InpCode, 1) = "P" Then
    Call CalcSetup(FluidName, "P", Units, Prop1, Prop2)
    If ierr > 0 Then LatentHeat = Trim2(herr): Exit Function
    If nc <> 1 Then LatentHeat = Trim2("Can only be calculated for pure fluids"): Exit Function
    Call INFOdll(1, wmm, ttrp, tnbpt, tc, pc, dc, Zc, acf, dip, Rgas)
    P = Prop1
    If P <= 0 Then LatentHeat = Trim2("Input pressure is zero"): Exit Function
    If P > pc Then LatentHeat = Trim2("Pressure is greater than the critical point pressure"): Exit Function
    Call SATPdll(P, x(1), 1&, T, Dl, Dv, xliq(1), xvap(1), ierr, herr, 255&)
    If (T = 0 Or Dl = 0) And ierr = 0 Then ierr = 1: LatentHeat = Trim2("Inputs are out of range"): Exit Function
  Else
    LatentHeat = Trim2("Valid inputs are only 'T' or 'P'"): Exit Function
  End If
  If ierr > 0 Then LatentHeat = Trim2(herr): Exit Function
  Call THERMdll(T, Dl, x(1), P, e, hl, s, Cvcalc, Cpcalc, W, hjt)
  Call THERMdll(T, Dv, x(1), P, e, hv, s, Cvcalc, Cpcalc, W, hjt)
  LatentHeat = ConvertUnits("-H", Units, hv - hl, 0)
End Function

Function HeatOfVaporization(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  HeatOfVaporization = LatentHeat(FluidName, InpCode, Units, Prop1, Prop2)
End Function

Function HeatOfCombustion(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim hg As Double, hn As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Call HEATdll(T, d, x(1), hg, hn, ierr, herr, 255&)
  HeatOfCombustion = ConvertUnits("-H", Units, hg, 0)
End Function

Function GrossHeatingValue(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim hg As Double, hn As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Call HEATdll(T, d, x(1), hg, hn, ierr, herr, 255&)
  GrossHeatingValue = ConvertUnits("-H", Units, hg, 0)
End Function

Function NetHeatingValue(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim hg As Double, hn As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Call HEATdll(T, d, x(1), hg, hn, ierr, herr, 255&)
  NetHeatingValue = ConvertUnits("-H", Units, hn, 0)
End Function

Function JouleThomson(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then JouleThomson = Trim2("Inputs are 2-phase"): Exit Function
  Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
  JouleThomson = ConvertUnits("-J", Units, hjt, 0)
End Function

Function IsentropicExpansionCoef(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then IsentropicExpansionCoef = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM3dll(T, d, x(1), xkappa, beta, isenk, kt, betas, bs, kkt, thrott, xpi, spht)
  If ierr > 0 Then IsentropicExpansionCoef = Trim2(herr): Exit Function
  IsentropicExpansionCoef = isenk
End Function

Function IsothermalCompressibility(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then IsothermalCompressibility = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM2dll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, z, hjt, aHelm, Gibbs, xkappa, beta, dPdD_T, d2PdD2_rho, dPdT_rho, dDdT_P, dDdP_rho, d2PT2, d2PdTD, spare3, spare4)
  If ierr > 0 Then IsothermalCompressibility = Trim2(herr): Exit Function
  IsothermalCompressibility = Trim2("Infinite")
  If d > 1E-20 And Not (xkappa = -9999990 Or xkappa > 1E+15) Then IsothermalCompressibility = 1 / ConvertUnits("-P", Units, 1 / xkappa, 0)
End Function

Function VolumeExpansivity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then VolumeExpansivity = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM2dll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, z, hjt, aHelm, Gibbs, xkappa, beta, dPdD_T, d2PdD2_rho, dPdT_rho, dDdT_P, dDdP_rho, d2PT2, d2PdTD, spare3, spare4)
  If ierr > 0 Then VolumeExpansivity = Trim2(herr): Exit Function
  VolumeExpansivity = 1 / ConvertUnits("-A", Units, 1 / beta, 0)
End Function

Function AdiabaticCompressibility(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then AdiabaticCompressibility = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM3dll(T, d, x(1), xkappa, beta, isenk, kt, betas, bs, kkt, thrott, xpi, spht)
  If ierr > 0 Then AdiabaticCompressibility = Trim2(herr): Exit Function
  AdiabaticCompressibility = Trim2("Infinite")
  If d > 1E-20 Then AdiabaticCompressibility = 1 / ConvertUnits("-P", Units, 1 / betas, 0)
End Function

Function AdiabaticBulkModulus(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then AdiabaticBulkModulus = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM3dll(T, d, x(1), xkappa, beta, isenk, kt, betas, bs, kkt, thrott, xpi, spht)
  AdiabaticBulkModulus = ConvertUnits("-P", Units, bs, 0)
End Function

Function IsothermalExpansionCoef(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then IsothermalExpansionCoef = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM3dll(T, d, x(1), xkappa, beta, isenk, kt, betas, bs, kkt, thrott, xpi, spht)
  If ierr > 0 Then IsothermalExpansionCoef = Trim2(herr): Exit Function
  IsothermalExpansionCoef = kt
End Function

Function IsothermalBulkModulus(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then IsothermalBulkModulus = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM3dll(T, d, x(1), xkappa, beta, isenk, kt, betas, bs, kkt, thrott, xpi, spht)
  IsothermalBulkModulus = ConvertUnits("-P", Units, kkt, 0)
End Function

Function SpecificHeatInput(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then SpecificHeatInput = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM3dll(T, d, x(1), xkappa, beta, isenk, kt, betas, bs, kkt, thrott, xpi, spht)
  SpecificHeatInput = ConvertUnits("-H", Units, spht, 0)
End Function

Function SecondVirial(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim b As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  Call VIRBdll(T, x(1), b)
  SecondVirial = ConvertUnits("-V", Units, b, 0)
End Function

Function dPdrho(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dPdrho = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM2dll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, z, hjt, aHelm, Gibbs, xkappa, beta, dPdD_T, d2PdD2_rho, dPdT_rho, dDdT_P, dDdP_rho, d2PT2, d2PdTD, spare3, spare4)
  If ierr > 0 Then dPdrho = Trim2(herr): Exit Function
  dPdrho = ConvertUnits("-P", Units, 1 / ConvertUnits("-D", Units, 1 / dPdD_T, 0), 0)
End Function

Function d2Pdrho2(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then d2Pdrho2 = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM2dll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, z, hjt, aHelm, Gibbs, xkappa, beta, dPdD_T, d2PdD2_rho, dPdT_rho, dDdT_P, dDdP_rho, d2PT2, d2PdTD, spare3, spare4)
  If ierr > 0 Then d2Pdrho2 = Trim2(herr): Exit Function
  d2Pdrho2 = ConvertUnits("-P", Units, 1 / ConvertUnits("-D", Units, ConvertUnits("-D", Units, 1 / d2PdD2_rho, 0), 0), 0)
End Function

Function dPdT(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dPdT = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM2dll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, z, hjt, aHelm, Gibbs, xkappa, beta, dPdD_T, d2PdD2_rho, dPdT_rho, dDdT_P, dDdP_rho, d2PT2, d2PdTD, spare3, spare4)
  If ierr > 0 Then dPdT = Trim2(herr): Exit Function
  dPdT = 0
  If dPdT_rho <> 0 Then dPdT = ConvertUnits("-P", Units, 1 / ConvertUnits("-A", Units, 1 / dPdT_rho, 0), 0)
End Function

Function dPdTsat(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim Cst As Double, dPT As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If nc <> 1 Then dPdTsat = "Functions work only for pure fluids"
  If ierr > 0 Then dPdTsat = Trim2(herr): Exit Function
  Call DPTSATKdll(1, T, 1&, P, d, Cst, dPT, ierr, herr, 255&)
  If ierr > 0 Then dPdTsat = Trim2(herr): Exit Function
  If dPT = 0 Or ierr > 0 Then dPdTsat = herr: Exit Function
  dPdTsat = ConvertUnits("-P", Units, 1 / ConvertUnits("-A", Units, 1 / dPT, 0), 0)
End Function

Function dHdT_D(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim dHdT_Dcalc As Double, dHdT_Pcalc As Double, dHdD_Tcalc As Double, dHdD_Pcalc As Double, dHdP_Tcalc As Double, dHdP_Dcalc As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dHdT_D = Trim2("Inputs are 2-phase"): Exit Function
  Call DHD1dll(T, d, x(1), dHdT_Dcalc, dHdT_Pcalc, dHdD_Tcalc, dHdD_Pcalc, dHdP_Tcalc, dHdP_Dcalc)
  dHdT_D = ConvertUnits("-S", Units, dHdT_Dcalc, 0)
End Function

Function dHdT_P(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim dHdT_Dcalc As Double, dHdT_Pcalc As Double, dHdD_Tcalc As Double, dHdD_Pcalc As Double, dHdP_Tcalc As Double, dHdP_Dcalc As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dHdT_P = Trim2("Inputs are 2-phase"): Exit Function
  Call DHD1dll(T, d, x(1), dHdT_Dcalc, dHdT_Pcalc, dHdD_Tcalc, dHdD_Pcalc, dHdP_Tcalc, dHdP_Dcalc)
  dHdT_P = ConvertUnits("-S", Units, dHdT_Pcalc, 0)
End Function

Function dHdD_T(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim dHdT_Dcalc As Double, dHdT_Pcalc As Double, dHdD_Tcalc As Double, dHdD_Pcalc As Double, dHdP_Tcalc As Double, dHdP_Dcalc As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dHdD_T = Trim2("Inputs are 2-phase"): Exit Function
  Call DHD1dll(T, d, x(1), dHdT_Dcalc, dHdT_Pcalc, dHdD_Tcalc, dHdD_Pcalc, dHdP_Tcalc, dHdP_Dcalc)
  If ierr > 0 Then dHdD_T = Trim2(herr): Exit Function
  dHdD_T = ConvertUnits("-H", Units, 1 / ConvertUnits("-D", Units, 1 / dHdD_Tcalc, 0), 0)
End Function

Function dHdD_P(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim dHdT_Dcalc As Double, dHdT_Pcalc As Double, dHdD_Tcalc As Double, dHdD_Pcalc As Double, dHdP_Tcalc As Double, dHdP_Dcalc As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dHdD_P = Trim2("Inputs are 2-phase"): Exit Function
  Call DHD1dll(T, d, x(1), dHdT_Dcalc, dHdT_Pcalc, dHdD_Tcalc, dHdD_Pcalc, dHdP_Tcalc, dHdP_Dcalc)
  If ierr > 0 Then dHdD_P = Trim2(herr): Exit Function
  dHdD_P = ConvertUnits("-H", Units, 1 / ConvertUnits("-D", Units, 1 / dHdD_Pcalc, 0), 0)
End Function

Function dHdP_T(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim dHdT_Dcalc As Double, dHdT_Pcalc As Double, dHdD_Tcalc As Double, dHdD_Pcalc As Double, dHdP_Tcalc As Double, dHdP_Dcalc As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dHdP_T = Trim2("Inputs are 2-phase"): Exit Function
  Call DHD1dll(T, d, x(1), dHdT_Dcalc, dHdT_Pcalc, dHdD_Tcalc, dHdD_Pcalc, dHdP_Tcalc, dHdP_Dcalc)
  If ierr > 0 Then dHdP_T = Trim2(herr): Exit Function
  dHdP_T = ConvertUnits("-H", Units, 1 / ConvertUnits("-P", Units, 1 / dHdP_Tcalc, 0), 0)
End Function

Function dHdP_D(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim dHdT_Dcalc As Double, dHdT_Pcalc As Double, dHdD_Tcalc As Double, dHdD_Pcalc As Double, dHdP_Tcalc As Double, dHdP_Dcalc As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then dHdP_D = Trim2("Inputs are 2-phase"): Exit Function
  Call DHD1dll(T, d, x(1), dHdT_Dcalc, dHdT_Pcalc, dHdD_Tcalc, dHdD_Pcalc, dHdP_Tcalc, dHdP_Dcalc)
  If ierr > 0 Then dHdP_D = Trim2(herr): Exit Function
  dHdP_D = ConvertUnits("-H", Units, 1 / ConvertUnits("-P", Units, 1 / dHdP_Dcalc, 0), 0)
End Function

Function drhodT(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If q > 0 And q < 1 Then drhodT = Trim2("Inputs are 2-phase"): Exit Function
  Call THERM2dll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, z, hjt, aHelm, Gibbs, xkappa, beta, dPdD_T, d2PdD2_rho, dPdT_rho, dDdT_P, dDdP_rho, d2PT2, d2PdTD, spare3, spare4)
  If ierr > 0 Then drhodT = Trim2(herr): Exit Function
  drhodT = 0
  If dDdT_P <> 0 Then drhodT = ConvertUnits("-D", Units, 1 / ConvertUnits("-A", Units, 1 / dDdT_P, 0), 0)
End Function

Function Cstar(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim v As Double, cs As Double, ts As Double, Ds As Double, ps As Double, ws As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  v = 0
  On Error GoTo ErrorHandler:
  Call CSTARdll(T, P, v, x(1), cs, ts, Ds, ps, ws, ierr, herr, 255&)
  If ierr > 0 Then Cstar = Trim2(herr): Exit Function
  Cstar = cs
  Return

ErrorHandler:
  'Call CCRITdll(t, p, v, x(1), cs, ts, Ds, ps, ws, ierr, herr, 255&)   'Old format
  Cstar = cs
End Function

Function Quality(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim xlkg(1 To MaxComps) As Double, xvkg(1 To MaxComps) As Double, qkg As Double, wliq As Double, wvap As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Quality = Trim2(herr): Exit Function
  Quality = q
  If q = 990 Then Quality = Trim2("Not calculated")
  If q = 998 Or q > 1 Then Quality = Trim2("Superheated vapor")
  If q = 999 Then Quality = Trim2("Supercritical state (T>Tc, p>pc)")
  If q < 0 Then Quality = Trim2("Subcooled liquid")
  If q = -998 Then Quality = Trim2("Subcooled liquid with p>pc")
  If q > 0 And q < 1 And molmass = 2 Then
    Call QMASSdll(q, xliq(1), xvap(1), qkg, xlkg(1), xvkg(1), wliq, wvap, ierr2, herr2, 255&)
    Quality = qkg
  End If
End Function

Function QualityMole(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim xlkg(1 To MaxComps) As Double, xvkg(1 To MaxComps) As Double, qkg As Double, wliq As Double, wvap As Double
  QualityMole = Quality(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Exit Function
  If q > 0 And q < 1 Then QualityMole = q
End Function

Function QualityMass(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim xlkg(1 To MaxComps) As Double, xvkg(1 To MaxComps) As Double, qkg As Double, wliq As Double, wvap As Double
  QualityMass = Quality(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Exit Function
  If q > 0 And q < 1 Then
    Call QMASSdll(q, xliq(1), xvap(1), qkg, xlkg(1), xvkg(1), wliq, wvap, ierr2, herr2, 255&)
    QualityMass = qkg
  End If
End Function

Function LiquidMoleFraction(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then LiquidMoleFraction = Trim2(herr): Exit Function
  If IsMissing(i) Then LiquidMoleFraction = Trim2("Component number is missing"): Exit Function
  If i < 1 Or i > nc Then LiquidMoleFraction = Trim2("Constituent number out of range"): Exit Function
  If q < 0 Or q > 1 Then
    LiquidMoleFraction = x(i)
  Else
    LiquidMoleFraction = xliq(i)
  End If
  If nc = 1 Then LiquidMoleFraction = Trim2("Not applicable for a pure fluid")
End Function

Function VaporMoleFraction(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then VaporMoleFraction = Trim2(herr): Exit Function
  If IsMissing(i) Then VaporMoleFraction = Trim2("Component number is missing"): Exit Function
  If i < 1 Or i > nc Then VaporMoleFraction = Trim2("Constituent number out of range"): Exit Function
  If q < 0 Or q > 1 Then
    VaporMoleFraction = x(i)
  Else
    VaporMoleFraction = xvap(i)
  End If

  If nc = 1 Then VaporMoleFraction = Trim2("Not applicable for a pure fluid")
End Function

Function LiquidMassFraction(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then LiquidMassFraction = Trim2(herr): Exit Function
  If IsMissing(i) Then LiquidMassFraction = Trim2("Component number is missing"): Exit Function
  If i < 1 Or i > nc Then LiquidMassFraction = Trim2("Constituent number out of range"): Exit Function
  If q < 0 Or q > 1 Then
    Call XMASSdll(x(1), xmm(1), wm)
    LiquidMassFraction = xmm(i)
  Else
    Call XMASSdll(xliq(1), xmm(1), wm)
    LiquidMassFraction = xmm(i)
  End If
  If nc = 1 Then LiquidMassFraction = Trim2("Not applicable for a pure fluid")
End Function

Function VaporMassFraction(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then VaporMassFraction = Trim2(herr): Exit Function
  If IsMissing(i) Then VaporMassFraction = Trim2("Component number is missing"): Exit Function
  If i < 1 Or i > nc Then VaporMassFraction = Trim2("Constituent number out of range"): Exit Function
  If q < 0 Or q > 1 Then
    Call XMASSdll(x(1), xmm(1), wm)
    VaporMassFraction = xmm(i)
  Else
    Call XMASSdll(xvap(1), xmm(1), wm)
    VaporMassFraction = xmm(i)
  End If

  If nc = 1 Then VaporMassFraction = Trim2("Not applicable for a pure fluid")
End Function

Function Fugacity(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Dim f(MaxComps) As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Fugacity = Trim2(herr): Exit Function
  If IsMissing(i) Then Fugacity = Trim2("Component number is missing "): Exit Function
  If i < 1 Or i > nc Then Fugacity = Trim2("Constituent number out of Range "): Exit Function
  If q > 0 And q < 1 Then Fugacity = Trim2("Inputs are 2-phase"): Exit Function
  Call FGCTYdll(T, d, x(1), f(1))
  Fugacity = ConvertUnits("-P", Units, f(i), 0)
End Function

Function FugacityCoefficient(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Dim f(MaxComps) As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then FugacityCoefficient = Trim2(herr): Exit Function
  If IsMissing(i) Then FugacityCoefficient = Trim2("Component number is missing "): Exit Function
  If i < 1 Or i > nc Then FugacityCoefficient = Trim2("Constituent number out of Range "): Exit Function
  If q > 0 And q < 1 Then FugacityCoefficient = Trim2("Inputs are 2-phase"): Exit Function
  Call FUGCOFdll(T, d, x(1), f(1), ierr, herr, 255&)
  If ierr > 0 Then FugacityCoefficient = Trim2(herr): Exit Function
  FugacityCoefficient = f(i)
End Function

Function ChemicalPotential(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Dim f(MaxComps) As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then ChemicalPotential = Trim2(herr): Exit Function
  If IsMissing(i) Then ChemicalPotential = Trim2("Component number is missing "): Exit Function
  If i < 1 Or i > nc Then ChemicalPotential = Trim2("Constituent number out of Range "): Exit Function
  If q > 0 And q < 1 Then ChemicalPotential = Trim2("Inputs are 2-phase"): Exit Function
  Call CHEMPOTdll(T, d, x(1), f(1), ierr, herr, 255&)
  If ierr > 0 Then ChemicalPotential = Trim2(herr): Exit Function
  ChemicalPotential = ConvertUnits("-H", Units, f(i), 0)
End Function

Function Activity(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Dim f(MaxComps) As Double, gamma(MaxComps) As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Activity = Trim2(herr): Exit Function
  If IsMissing(i) Then Activity = Trim2("Component number is missing "): Exit Function
  If i < 1 Or i > nc Then Activity = Trim2("Constituent number out of Range "): Exit Function
  If q > 0 And q < 1 Then Activity = Trim2("Inputs are 2-phase"): Exit Function
  Call ACTVYdll(T, d, x(1), f(1), gamma(1), ierr, herr, 255&)
  If ierr > 0 Then Activity = Trim2(herr): Exit Function
  Activity = f(i)
End Function

Function ActivityCoefficient(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2, Optional i)
  Dim f(MaxComps) As Double, gamma(MaxComps) As Double
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then ActivityCoefficient = Trim2(herr): Exit Function
  If IsMissing(i) Then ActivityCoefficient = Trim2("Component number is missing "): Exit Function
  If i < 1 Or i > nc Then ActivityCoefficient = Trim2("Constituent number out of Range "): Exit Function
  If q > 0 And q < 1 Then ActivityCoefficient = Trim2("Inputs are 2-phase"): Exit Function
  Call ACTVYdll(T, d, x(1), f(1), gamma(1), ierr, herr, 255&)
  If ierr > 0 Then ActivityCoefficient = Trim2(herr): Exit Function
  ActivityCoefficient = gamma(i)
End Function

Function Viscosity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Viscosity = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then Viscosity = Trim2("Inputs out of range"): Exit Function
  Call TRNPRPdll(T, d, x(1), eta, tcx, ierr2, herr2, 255&)
  If q > 0 And q < 1 Then eta = -9999999
  Viscosity = ConvertUnits("-U", Units, eta, 0)
  If eta = 0 Then Viscosity = Trim2("Unable to calculate property")
End Function

Function ThermalConductivity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then ThermalConductivity = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then ThermalConductivity = Trim2("Inputs out of range"): Exit Function
  Call TRNPRPdll(T, d, x(1), eta, tcx, ierr2, herr2, 255&)
  If q > 0 And q < 1 Then tcx = -9999999
  ThermalConductivity = ConvertUnits("-K", Units, tcx, 0)
  If tcx = 0 Then ThermalConductivity = Trim2("Unable to calculate property")
End Function

Function KinematicViscosity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then KinematicViscosity = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then KinematicViscosity = Trim2("Inputs out of range"): Exit Function
  Call TRNPRPdll(T, d, x(1), eta, tcx, ierr2, herr2, 255&)
  If eta = 0 Then KinematicViscosity = Trim2("Unable to calculate property"): Exit Function
  If q > 0 And q < 1 Then
    eta = -9999999
  Else
    eta = eta / d / wm / 100  'cm^2/s
  End If
  KinematicViscosity = ConvertUnits("-I", Units, eta, 0)
End Function

Function ThermalDiffusivity(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then ThermalDiffusivity = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then ThermalDiffusivity = Trim2("Inputs out of range"): Exit Function
  Call TRNPRPdll(T, d, x(1), eta, tcx, ierr2, herr2, 255&)
  If tcx = 0 Then ThermalDiffusivity = Trim2("Unable to calculate property"): Exit Function
  If q > 0 And q < 1 Then
    tcx = -9999999
  Else
    Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
    tcx = tcx / d / Cpcalc * 10  'cm^2/s
  End If
  ThermalDiffusivity = ConvertUnits("-I", Units, tcx, 0)
End Function

Function Prandtl(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Prandtl = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then Prandtl = Trim2("Inputs out of range"): Exit Function
  Call TRNPRPdll(T, d, x(1), eta, tcx, ierr2, herr2, 255&)
  If q > 0 And q < 1 Then Prandtl = Trim2("Undefined"): Exit Function
  If tcx = 0 Or eta = 0 Then Prandtl = Trim2("Unable to calculate property")
  Call THERMdll(T, d, x(1), P, e, h, s, Cvcalc, Cpcalc, W, hjt)
  Call WMOLdll(x(1), wm)
  Prandtl = eta * Cpcalc / tcx / wm / 1000
End Function

Function SurfaceTension(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then SurfaceTension = Trim2(herr): Exit Function
  If T = 0 Then SurfaceTension = Trim2("Input temperature is zero"): Exit Function
  If q >= 0 And q <= 1 Then
    Call SURFTdll(T, Dl, xliq(1), sigma, ierr2, herr2, 255&)
  Else
    Call SURFTdll(T, d, x(1), sigma, ierr2, herr2, 255&)
  End If
  SurfaceTension = ConvertUnits("-N", Units, sigma, 0)
  If sigma = 0 Or ierr2 > 0 Then SurfaceTension = Trim2("Inputs out of range")
End Function

Function DielectricConstant(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then DielectricConstant = Trim2(herr): Exit Function
  If q > 0 And q < 1 Then DielectricConstant = Trim2("Undefined"): Exit Function
  If T = 0 Then DielectricConstant = Trim2("Inputs out of range"): Exit Function
  Call DIELECdll(T, d, x(1), de)
  DielectricConstant = de
End Function

Function MolarMass(FluidName, Optional InpCode, Optional Units, Optional ByVal Prop1, Optional ByVal Prop2)
  Call CalcSetup(FluidName, "", Units, Prop1, Prop2)
  If ierr > 0 Then MolarMass = Trim2(herr): Exit Function
  Call WMOLdll(x(1), wm)
  MolarMass = wm
End Function

Function MoleFraction(FluidName, i)
  Call CalcProp(FluidName, "", "", 0, 0)
  If ierr > 0 Then MoleFraction = Trim2(herr): Exit Function
  If IsMissing(i) Then MoleFraction = Trim2("Component number is missing"): Exit Function
  If i < 1 Or i > nc Then MoleFraction = Trim2("Constituent number out of range"): Exit Function
  MoleFraction = x(i)
  If nc = 1 Then MoleFraction = Trim2("Not applicable for a pure fluid")
End Function

Function MassFraction(FluidName, i)
  Call CalcProp(FluidName, "", "", 0, 0)
  If ierr > 0 Then MassFraction = Trim2(herr): Exit Function
  If IsMissing(i) Then MassFraction = Trim2("Component number is missing"): Exit Function
  If i < 1 Or i > nc Then MassFraction = Trim2("Constituent number out of range"): Exit Function
  Call XMASSdll(x(1), xmm(1), wm)
  MassFraction = xmm(i)
  If nc = 1 Then MassFraction = Trim2("Not applicable for a pure fluid")
End Function

Function ComponentName(FluidName, i)
  Dim a As String, j As Integer, k As Integer
  Call CalcProp(FluidName, "", "", 0, 0)
  If ierr > 0 Then ComponentName = Trim2(herr): Exit Function
  If IsMissing(i) Then ComponentName = Trim2("Component number is missing"): Exit Function
  If i < 1 Or i > nc Then ComponentName = Trim2("Constituent number out of range"): Exit Function
  a = hfld
  If nc > 1 Then
    For k = 1 To i - 1
      j = InStr(a, "|")
      If j Then a = Mid(a, j + 1)
    Next
    j = InStr(a, "|")
    If j Then a = Left(a, j - 1)
  End If
  Do
    j = InStr(a, "\")
    If j Then a = Mid(a, j + 1)
  Loop While j <> 0
  j = InStr(a, ".")
  If j Then a = Left(a, j - 1)
  ComponentName = a
End Function

Function LiquidFluidString(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim a As String, i As Integer
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then LiquidFluidString = Trim2(herr): Exit Function
  For i = 1 To nc
    a = a + ";" + ComponentName(FluidName, i)
    If q < 0 Or q > 1 Then
      a = a + ";" + Trim(Str(x(i)))
    Else
      a = a + ";" + Trim(Str(xliq(i)))
    End If
  Next
  LiquidFluidString = Mid(a, 2)
  If nc = 1 Then LiquidFluidString = ComponentName(FluidName, i)
End Function

Function VaporFluidString(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Dim a As String, i As Integer
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then VaporFluidString = Trim2(herr): Exit Function
  For i = 1 To nc
    a = a + ";" + ComponentName(FluidName, i)
    If q < 0 Or q > 1 Then
      a = a + ";" + Trim(Str(x(i)))
    Else
      a = a + ";" + Trim(Str(xvap(i)))
    End If
  Next
  VaporFluidString = Mid(a, 2)
  If nc = 1 Then VaporFluidString = ComponentName(FluidName, i)
End Function

'Change molar composition to mass composition
'Prop1 - Prop20 are the molar values for the components in the mixture.
'i specifies which component's mole fraction is returned.  If zero, the molar mass is returned
Function Mole2Mass(FluidName, i, Prop1, Prop2, Optional Prop3, Optional Prop4, Optional Prop5, Optional Prop6, Optional Prop7, Optional Prop8, Optional Prop9, Optional Prop10, Optional Prop11, Optional Prop12, Optional Prop13, Optional Prop14, Optional Prop15, Optional Prop16, Optional Prop17, Optional Prop18, Optional Prop19, Optional Prop20)
Dim j As Integer, xkg2(1 To MaxComps) As Double, xmol2(1 To MaxComps) As Double, wmix2 As Double, sum As Double
For j = 1 To MaxComps: xmol2(j) = 0: Next
xmol2(1) = Prop1
xmol2(2) = Prop2
If IsMissing(Prop3) = False And IsNumeric(Prop3) = True Then xmol2(3) = Prop3
If IsMissing(Prop4) = False And IsNumeric(Prop4) = True Then xmol2(4) = Prop4
If IsMissing(Prop5) = False And IsNumeric(Prop5) = True Then xmol2(5) = Prop5
If IsMissing(Prop6) = False And IsNumeric(Prop6) = True Then xmol2(6) = Prop6
If IsMissing(Prop7) = False And IsNumeric(Prop7) = True Then xmol2(7) = Prop7
If IsMissing(Prop8) = False And IsNumeric(Prop8) = True Then xmol2(8) = Prop8
If IsMissing(Prop9) = False And IsNumeric(Prop9) = True Then xmol2(9) = Prop9
If IsMissing(Prop10) = False And IsNumeric(Prop10) = True Then xmol2(10) = Prop10
If IsMissing(Prop11) = False And IsNumeric(Prop11) = True Then xmol2(11) = Prop11
If IsMissing(Prop12) = False And IsNumeric(Prop12) = True Then xmol2(12) = Prop12
If IsMissing(Prop13) = False And IsNumeric(Prop13) = True Then xmol2(13) = Prop13
If IsMissing(Prop14) = False And IsNumeric(Prop14) = True Then xmol2(14) = Prop14
If IsMissing(Prop15) = False And IsNumeric(Prop15) = True Then xmol2(15) = Prop15
If IsMissing(Prop16) = False And IsNumeric(Prop16) = True Then xmol2(16) = Prop16
If IsMissing(Prop17) = False And IsNumeric(Prop17) = True Then xmol2(17) = Prop17
If IsMissing(Prop18) = False And IsNumeric(Prop18) = True Then xmol2(18) = Prop18
If IsMissing(Prop19) = False And IsNumeric(Prop19) = True Then xmol2(19) = Prop19
If IsMissing(Prop20) = False And IsNumeric(Prop20) = True Then xmol2(20) = Prop20
Call CalcSetup(FluidName, "", "", 0, 0)
If ierr > 0 Then Mole2Mass = Trim2(herr): Exit Function
If i < 0 Or i > nc Then Mole2Mass = Trim2("Index out of Range (greater than number of components in mixture)"):  Exit Function
sum = 0
For j = 1 To nc
  sum = sum + xmol2(j)
Next
If Abs(sum - 1) > 0.0001 Then Mole2Mass = Trim2("Composition does not sum to 1"): Exit Function
Call XMASSdll(xmol2(1), xkg2(1), wmix2)
If i = 0 Then  'Molar mass of mixture
  Mole2Mass = wmix2
Else               'Mass fraction
  Mole2Mass = xkg2(i)
End If
End Function

'Change mass composition to molar composition
'Prop1 - Prop20 are the mass values for the components in the mixture.
'i specifies which component's mass fraction is returned.  If zero, the molar mass is returned
Function Mass2Mole(FluidName, i, Prop1, Prop2, Optional Prop3, Optional Prop4, Optional Prop5, Optional Prop6, Optional Prop7, Optional Prop8, Optional Prop9, Optional Prop10, Optional Prop11, Optional Prop12, Optional Prop13, Optional Prop14, Optional Prop15, Optional Prop16, Optional Prop17, Optional Prop18, Optional Prop19, Optional Prop20)
Dim j As Integer, xkg2(1 To MaxComps) As Double, xmol2(1 To MaxComps) As Double, wmix2 As Double, sum As Double
For j = 1 To MaxComps: xkg2(j) = 0: Next
xkg2(1) = Prop1
xkg2(2) = Prop2
If IsMissing(Prop3) = False And IsNumeric(Prop3) = True Then xkg2(3) = Prop3
If IsMissing(Prop4) = False And IsNumeric(Prop4) = True Then xkg2(4) = Prop4
If IsMissing(Prop5) = False And IsNumeric(Prop5) = True Then xkg2(5) = Prop5
If IsMissing(Prop6) = False And IsNumeric(Prop6) = True Then xkg2(6) = Prop6
If IsMissing(Prop7) = False And IsNumeric(Prop7) = True Then xkg2(7) = Prop7
If IsMissing(Prop8) = False And IsNumeric(Prop8) = True Then xkg2(8) = Prop8
If IsMissing(Prop9) = False And IsNumeric(Prop9) = True Then xkg2(9) = Prop9
If IsMissing(Prop10) = False And IsNumeric(Prop10) = True Then xkg2(10) = Prop10
If IsMissing(Prop11) = False And IsNumeric(Prop11) = True Then xkg2(11) = Prop11
If IsMissing(Prop12) = False And IsNumeric(Prop12) = True Then xkg2(12) = Prop12
If IsMissing(Prop13) = False And IsNumeric(Prop13) = True Then xkg2(13) = Prop13
If IsMissing(Prop14) = False And IsNumeric(Prop14) = True Then xkg2(14) = Prop14
If IsMissing(Prop15) = False And IsNumeric(Prop15) = True Then xkg2(15) = Prop15
If IsMissing(Prop16) = False And IsNumeric(Prop16) = True Then xkg2(16) = Prop16
If IsMissing(Prop17) = False And IsNumeric(Prop17) = True Then xkg2(17) = Prop17
If IsMissing(Prop18) = False And IsNumeric(Prop18) = True Then xkg2(18) = Prop18
If IsMissing(Prop19) = False And IsNumeric(Prop19) = True Then xkg2(19) = Prop19
If IsMissing(Prop20) = False And IsNumeric(Prop20) = True Then xkg2(20) = Prop20
Call CalcSetup(FluidName, "", "", 0, 0)
If ierr > 0 Then Mass2Mole = Trim2(herr): Exit Function
If i < 0 Or i > nc Then Mass2Mole = Trim2("Index out of Range (greater than number of components in mixture)"):  Exit Function
sum = 0
For j = 1 To nc
  sum = sum + xkg2(j)
Next
If Abs(sum - 1) > 0.0001 Then Mass2Mole = Trim2("Composition does not sum to 1"): Exit Function
Call XMOLEdll(xkg2(1), xmol2(1), wmix2)
If i = 0 Then  'Molar mass of mixture
  Mass2Mole = wmix2
Else               'Mole fraction
  Mass2Mole = xmol2(i)
End If
End Function

Function EOSMax(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcSetup(FluidName, "", Units, Prop1, Prop2)
  If nc > 1 Then
    Call LIMITXdll("EOS", 300#, 0#, 0#, x(1), tmin, tmax, dmax, pmax, ierr2, herr2, 3&, 255&)
  Else
    Call LIMITKdll("EOS", 1, 300#, 0#, 0#, tmin, tmax, dmax, pmax, ierr2, herr2, 3&, 255&)
  End If
  If IsMissing(InpCode) Then InpCode = ""
  If InpCode = "P" Or InpCode = "p" Then
    EOSMax = ConvertUnits("-P", Units, pmax, 0)
  ElseIf InpCode = "D" Or InpCode = "d" Then
    EOSMax = ConvertUnits("-D", Units, dmax, 0)
  Else
    EOSMax = ConvertUnits("-T", Units, tmax, 0)
  End If
End Function

Function EOSMin(FluidName, Optional InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcSetup(FluidName, "", Units, Prop1, Prop2)
  If nc > 1 Then
    Call LIMITXdll("EOS", 300#, 0#, 0#, x(1), tmin, tmax, dmax, pmax, ierr2, herr2, 3&, 255&)
  Else
    Call LIMITKdll("EOS", 1, 300#, 0#, 0#, tmin, tmax, dmax, pmax, ierr2, herr2, 3&, 255&)
  End If
  If IsMissing(InpCode) Then InpCode = ""
  If InpCode = "P" Or InpCode = "p" Then
    EOSMin = 0
  ElseIf InpCode = "D" Or InpCode = "d" Then
    EOSMin = 0
  Else
    EOSMin = ConvertUnits("-T", Units, tmin, 0)
  End If
End Function

Function ErrorCode(InputCell)
  ErrorCode = ierr
End Function

Function ErrorString(InputCell)
  ErrorString = Trim2(herr)
End Function

Function Trim2(a)
'All error messages call this routine to add the pound sign (#) to the beginning of the line.
'If you do not want this error code, simply remove the ["#" +] piece below.
'It can also be changed to any other symbol(s) you desire.
  If Left(a, 1) <> "#" Then
    Trim2 = "#" + Trim(a)
  Else
    Trim2 = Trim(a)
  End If
End Function

Function Viscosity_ETAK0(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Viscosity_ETAK0 = Trim2(herr): Exit Function
  If T = 0 Then Viscosity_ETAK0 = Trim2("Inputs out of range"): Exit Function
  Call ETAK0dll(1&, T, eta0, ierr, herr, 255&)
  If q > 0 And q < 1 Then eta0 = -9999999
  Viscosity_ETAK0 = ConvertUnits("-U", Units, eta0, 0)
  If eta0 = 0 Then Viscosity_ETAK0 = Trim2("Unable to calculate property")
End Function

Function Viscosity_ETAK1(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Viscosity_ETAK1 = Trim2(herr): Exit Function
  If T = 0 Then Viscosity_ETAK1 = Trim2("Inputs out of range"): Exit Function
  Call ETAK1dll(1&, T, eta1, ierr, herr, 255&)
  If q > 0 And q < 1 Then eta1 = -9999999
  Viscosity_ETAK1 = ConvertUnits("-U", Units, eta1, 0)
  If eta1 = 0 Then Viscosity_ETAK1 = Trim2("Unable to calculate property")
End Function

Function Viscosity_ETAKR(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Viscosity_ETAKR = Trim2(herr): Exit Function
  If T = 0 Then Viscosity_ETAKR = Trim2("Inputs out of range"): Exit Function
  Call ETAKRdll(1&, T, d, etar, ierr, herr, 255&)
  If q > 0 And q < 1 Then etar = -9999999
  Viscosity_ETAKR = ConvertUnits("-U", Units, etar, 0)
  If etar = 0 Then Viscosity_ETAKR = Trim2("Unable to calculate property")
End Function

Function Viscosity_ETAKB(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Viscosity_ETAKB = Trim2(herr): Exit Function
  If T = 0 Then Viscosity_ETAKB = Trim2("Inputs out of range"): Exit Function
  Call ETAKBdll(1&, T, d, etab, ierr, herr, 255&)
  If q > 0 And q < 1 Then etab = -9999999
  Viscosity_ETAKB = ConvertUnits("-U", Units, etab, 0)
  If etab = 0 Then Viscosity_ETAKB = Trim2("Unable to calculate property")
End Function

Function Transport_Omega(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then Transport_Omega = Trim2(herr): Exit Function
  If T = 0 Then Transport_Omega = Trim2("Inputs out of range"): Exit Function
  'epsk is not read from the FLD file yet !!
  epsk = 280.51
  Call OMEGAdll(1, T, epsk, omg)
  If q > 0 And q < 1 Then omg = -9999999
  Transport_Omega = ConvertUnits("-U", Units, omg, 0)
  If omg = 0 Then Transport_Omega = Trim2("Unable to calculate property")
End Function

Function ThermalConductivity_TCXK0(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then ThermalConductivity_TCXK0 = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then ThermalConductivity_TCXK0 = Trim2("Inputs out of range"): Exit Function
  Call TCXK0dll(1&, T, tcx0, ierr, herr, 255&)
  If q > 0 And q < 1 Then tcx0 = -9999999
  ThermalConductivity_TCXK0 = ConvertUnits("-K", Units, tcx0, 0)
  If tcx0 = 0 Then ThermalConductivity_TCXK0 = Trim2("Unable to calculate property")
End Function

Function ThermalConductivity_TCXKB(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then ThermalConductivity_TCXKB = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then ThermalConductivity_TCXKB = Trim2("Inputs out of range"): Exit Function
  Call TCXKBdll(1&, T, d, tcxb, ierr, herr, 255&)
  If q > 0 And q < 1 Then tcxb = -9999999
  ThermalConductivity_TCXKB = ConvertUnits("-K", Units, tcxb, 0)
  If tcxb = 0 Then ThermalConductivity_TCXKB = Trim2("Unable to calculate property")
End Function

Function ThermalConductivity_TCXKC(FluidName, InpCode, Optional Units, Optional Prop1, Optional Prop2)
  Call CalcProp(FluidName, InpCode, Units, Prop1, Prop2)
  If ierr > 0 Then ThermalConductivity_TCXKC = Trim2(herr): Exit Function
  If T = 0 Or d = 0 Then ThermalConductivity_TCXKC = Trim2("Inputs out of range"): Exit Function
  Call TCXKCdll(1&, T, d, tcxc, ierr, herr, 255&)
  If q > 0 And q < 1 Then tcxc = -9999999
  ThermalConductivity_TCXKC = ConvertUnits("-K", Units, tcxc, 0)
  If tcxc = 0 Then ThermalConductivity_TCXKC = Trim2("Unable to calculate property")
End Function



Function UnitConvert(InputValue, UnitType As String, OldUnits As String, NewUnits As String)

'InputValue is the value to be converted from OldUnits to NewUnits
'UnitType is one of the following letters (one character only in most cases):
'UnitType     Unit name                          SI units
'  T         Temperature                            K
'  P         Pressure                               Pa
'  D         Density or specific volume         mol/m^3 or kg/m^3 (or m^3/mol or m^3/kg)
'  H         Enthalpy or specific energy        J/mol or J/kg
'  S         Entropy or heat capacity           J/mol-K or J/kg-K
'  W         Speed of sound                         m/s
'  U         Viscosity                              Pa-s
'  K         Thermal conductivity                   W/m-K
'  JT        Joule Thomson                          K/Pa
'  L         Length                                 m
'  A         Area                                   m^2
'  V         Volume                                 m^3
'  M         Mass                                   kg
'  F         Force                                  N
'  E         Energy                                 J
'  Q         Power                                  W
'  N         Surface tension                        N/m
' Gage pressures can be used by adding "_g" to the unit, e.g., "MPa_g"
' The different inputs for OldUnits and NewUnits can be found scattered in the text below.
' Several examples are given below:
' T=UnitConvert(323.15,"T","K","F")
' P=UnitConvert(1.01325,"P","bar","mmHg")
' V=UnitConvert(1000.,"D","kg/m^3","cm^3/mol")


Dim Value As Double, Tpe As String, Unit1 As String, Unit2 As String
Dim Drct As Integer, Gage As Integer, Vacm As Integer
Dim MolWt As Double

If Not IsNumeric(InputValue) Then UnitConvert = InputValue: Exit Function
If NewUnits = "" Then UnitConvert = InputValue: Exit Function
Value = InputValue
Tpe = UCase(Trim(UnitType))
Unit1 = UCase(Trim(OldUnits))
Unit2 = UCase(Trim(NewUnits))

Call WMOLdll(x(1), wm)
If CompFlag = 1 Then Call WMOLdll(xliq(1), wm)
If CompFlag = 2 Then Call WMOLdll(xvap(1), wm)
MolWt = wm
If MolWt = 0 Then MolWt = 1

For Drct = 1 To -1 Step -2
'-----------------------------------------------------------------------
'   Temperature Conversion
'-----------------------------------------------------------------------
  If Tpe = "T" Then
    If Unit1 = "K" Then
    ElseIf Unit1 = "C" Then
      Value = Value + Drct * CtoK
    ElseIf Unit1 = "R" Then
      Value = Value * RtoK ^ Drct
    ElseIf Unit1 = "F" Then
      If Drct = 1 Then
        'Value = RtoK * (Value + FtoR)    'Does not give exactly zero at 32 F
        Value = (Value - 32) * RtoK + CtoK
      Else
        'Value = Value / RtoK - FtoR      'Does not give exactly 32 at 273.15 K
        Value = (Value - CtoK) / RtoK + 32
      End If
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Pressure Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "P" Then
    Gage = InStr(Unit1, "GAGE")
    Vacm = InStr(Unit1, "VACM")
    If Gage = 0 Then Gage = InStr(Unit1, "_G")
    If Vacm = 0 Then Vacm = InStr(Unit1, "_V")
    If Gage <> 0 And Drct = -1 Then Value = Value - ATMtoMPa
    If Vacm <> 0 And Drct = -1 Then Value = ATMtoMPa - Value
    If Gage <> 0 Then Unit1 = Trim(Left(Unit1, Gage - 1))
    If Vacm <> 0 Then Unit1 = Trim(Left(Unit1, Vacm - 1))
    If Unit1 = "PA" Then
      Value = Value / 1000000 ^ Drct
    ElseIf Unit1 = "KPA" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "MPA" Then
      Value = Value
    ElseIf Unit1 = "GPA" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "BAR" Then
      Value = Value * BARtoMPA ^ Drct
    ElseIf Unit1 = "KBAR" Then
      Value = Value * (BARtoMPA * 1000) ^ Drct
    ElseIf Unit1 = "ATM" Then
      Value = Value * ATMtoMPa ^ Drct
    ElseIf Unit1 = "KGF/CM^2" Or Unit1 = "KG/CM^2" Or Unit1 = "ATA" Or Unit1 = "AT" Or Unit1 = "ATMA" Then
      Value = Value * (KGFtoN / 100) ^ Drct
    ElseIf Unit1 = "PSI" Or Unit1 = "PSIA" Then
      Value = Value * PSIAtoMPA ^ Drct
    ElseIf Unit1 = "PSF" Then
      Value = Value * (PSIAtoMPA / 144) ^ Drct
    ElseIf Unit1 = "MMHG" Or Unit1 = "TORR" Then
      Value = Value * MMHGtoMPA ^ Drct
    ElseIf Unit1 = "CMHG" Then
      Value = Value * (MMHGtoMPA * 10) ^ Drct
    ElseIf Unit1 = "INHG" Then
      Value = Value * (MMHGtoMPA * INtoM * 1000) ^ Drct
    ElseIf Unit1 = "INH2O" Then
      Value = Value * INH2OtoMPA ^ Drct
    ElseIf Unit1 = "PSIG" Then
      If Drct = 1 Then
        Value = PSIAtoMPA * Value + ATMtoMPa
      Else
        Value = (Value - ATMtoMPa) / PSIAtoMPA
      End If
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If
    If Gage <> 0 And Drct = 1 Then Value = Value + ATMtoMPa
    If Vacm <> 0 And Drct = 1 Then Value = ATMtoMPa - Value

'-----------------------------------------------------------------------
'   Density Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "D" Then
    If Value = 0 Then Value = 1E-50
    If Unit1 = "MOL/DM^3" Or Unit1 = "MOL/L" Or Unit1 = "KMOL/M^3" Then
    ElseIf Unit1 = "MOL/CM^3" Or Unit1 = "MOL/CC" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "MOL/M^3" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "KG/M^3" Then
      Value = Value / MolWt ^ Drct
    ElseIf Unit1 = "KG/DM^3" Or Unit1 = "KG/L" Then
      Value = Value * (1000 / MolWt) ^ Drct
    ElseIf Unit1 = "G/DM^3" Or Unit1 = "G/L" Then
      Value = Value * (1 / MolWt) ^ Drct
    ElseIf Unit1 = "G/CC" Or Unit1 = "G/CM^3" Or Unit1 = "G/ML" Then
      Value = Value * (1000 / MolWt) ^ Drct
    ElseIf Unit1 = "G/DM^3" Then
      Value = Value * (1 / MolWt) ^ Drct
    ElseIf Unit1 = "LBM/FT^3" Or Unit1 = "LB/FT^3" Then
      Value = Value * (LBMtoKG / FT3toM3 / MolWt) ^ Drct
    ElseIf Unit1 = "LBMOL/FT^3" Then
      Value = Value * (LBMtoKG / FT3toM3) ^ Drct
    ElseIf Unit1 = "SLUG/FT^3" Then
      Value = Value * (LBMtoKG / FT3toM3 / MolWt * KGFtoN / FTtoM) ^ Drct
    ElseIf Unit1 = "LBMOL/GAL" Then
      Value = Value * (LBMtoKG / GALLONtoM3) ^ Drct
    ElseIf Unit1 = "LB/GAL" Or Unit1 = "LBM/GAL" Then
      Value = Value * (LBMtoKG / GALLONtoM3 / MolWt) ^ Drct

'-----------------------------------------------------------------------
'   Specific Volume Conversion
'-----------------------------------------------------------------------
    ElseIf Unit1 = "DM^3/MOL" Or Unit1 = "L/MOL" Or Unit1 = "M^3/KMOL" Then
      Value = 1 / Value
    ElseIf Unit1 = "CM^3/MOL" Or Unit1 = "CC/MOL" Or Unit1 = "ML/MOL" Then
      Value = 1000 / Value
    ElseIf Unit1 = "M^3/MOL" Then
      Value = 1 / Value / 1000
    ElseIf Unit1 = "M^3/KG" Then
      Value = 1 / Value / MolWt
    ElseIf Unit1 = "DM^3/KG" Or Unit1 = "L/KG" Then
      Value = 1000 / Value / MolWt
    ElseIf Unit1 = "CC/G" Or Unit1 = "CM^3/G" Or Unit1 = "ML/G" Then
      Value = 1000 / Value / MolWt
    ElseIf Unit1 = "DM^3/G" Then
      Value = 1 / Value / MolWt
    ElseIf Unit1 = "FT^3/LBM" Or Unit1 = "FT^3/LB" Then
      Value = 1 / Value * (LBMtoKG / FT3toM3 / MolWt)
    ElseIf Unit1 = "FT^3/LBMOL" Then
      Value = 1 / Value * (LBMtoKG / FT3toM3)
    ElseIf Unit1 = "FT^3/SLUG" Then
      Value = 1 / Value * (LBMtoKG / FT3toM3 / MolWt * KGFtoN / FTtoM)
    ElseIf Unit1 = "GAL/LBMOL" Then
      Value = 1 / Value * (LBMtoKG / GALLONtoM3)
    ElseIf Unit1 = "GAL/LB" Or Unit1 = "GAL/LBM" Then
      Value = 1 / Value * (LBMtoKG / GALLONtoM3 / MolWt)
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If
    If Abs(Value) < 1E-30 Then Value = 0

'-----------------------------------------------------------------------
'   Specific Energy and Enthalpy Conversions
'-----------------------------------------------------------------------
  ElseIf Tpe = "H" Then
    If Unit1 = "J/MOL" Or Unit1 = "KJ/KMOL" Then
    ElseIf Unit1 = "KJ/MOL" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "MJ/MOL" Then
      Value = Value * 1000000 ^ Drct
    ElseIf Unit1 = "KJ/KG" Or Unit1 = "J/G" Then
      Value = MolWt ^ Drct * Value
    ElseIf Unit1 = "J/KG" Then
      Value = (MolWt / 1000) ^ Drct * Value
    ElseIf Unit1 = "M^2/S^2" Then
      Value = (MolWt / 1000) ^ Drct * Value
    ElseIf Unit1 = "FT^2/S^2" Then
      Value = (MolWt / 1000 * FTtoM ^ 2) ^ Drct * Value
    ElseIf Unit1 = "CAL/MOL" Or Unit1 = "KCAL/KMOL" Then
      Value = CALtoJ ^ Drct * Value
    ElseIf Unit1 = "CAL/G" Or Unit1 = "KCAL/KG" Then
      Value = (CALtoJ * MolWt) ^ Drct * Value
    ElseIf Unit1 = "BTU/LBM" Or Unit1 = "BTU/LB" Then
      Value = (BTUtoKJ / LBMtoKG * MolWt) ^ Drct * Value
    ElseIf Unit1 = "BTU/LBMOL" Then
      Value = (BTUtoKJ / LBMtoKG) ^ Drct * Value
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Entropy and Heat Capacity Conversions
'-----------------------------------------------------------------------
  ElseIf Tpe = "S" Then
    If Unit1 = "J/MOL-K" Or Unit1 = "KJ/KMOL-K" Then
      Value = Value
    ElseIf Unit1 = "KJ/MOL-K" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "KJ/KG-K" Or Unit1 = "J/G-K" Then
      Value = MolWt ^ Drct * Value
    ElseIf Unit1 = "J/KG-K" Then
      Value = (MolWt / 1000) ^ Drct * Value
    ElseIf Unit1 = "BTU/LBM-R" Or Unit1 = "BTU/LB-R" Then
      Value = (BTUtoKJ / LBMtoKG / RtoK * MolWt) ^ Drct * Value
    ElseIf Unit1 = "BTU/LBMOL-R" Then
      Value = (BTUtoKJ / LBMtoKG / RtoK) ^ Drct * Value
    ElseIf Unit1 = "CAL/G-K" Or Unit1 = "CAL/G-C" Or Unit1 = "KCAL/KG-K" Or Unit1 = "KCAL/KG-C" Then
      Value = (CALtoJ * MolWt) ^ Drct * Value
    ElseIf Unit1 = "CAL/MOL-K" Or Unit1 = "CAL/MOL-C" Then
      Value = CALtoJ ^ Drct * Value
    ElseIf Unit1 = "FT-LBF/LBMOL-R" Then
      Value = (FTLBFtoJ / LBMtoKG / RtoK / 1000) ^ Drct * Value
    ElseIf Unit1 = "CP/R" Then
      Value = Ridgas ^ Drct * Value * 1000
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Speed of Sound Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "W" Then
    If Unit1 = "M/S" Then
    ElseIf Unit1 = "M^2/S^2" Then
      Value = Sqr(Value)
    ElseIf Unit1 = "CM/S" Then
      Value = Value / 100 ^ Drct
    ElseIf Unit1 = "KM/H" Then
      Value = Value * (1000 / HtoS) ^ Drct
    ElseIf Unit1 = "FT/S" Then
      Value = Value * FTtoM ^ Drct
    ElseIf Unit1 = "IN/S" Then
      Value = Value * INtoM ^ Drct
    ElseIf Unit1 = "MILE/H" Or Unit1 = "MPH" Then
      Value = Value * (INtoM * 63360 / HtoS) ^ Drct
    ElseIf Unit1 = "KNOT" Then
      Value = Value * 0.5144444444 ^ Drct
    ElseIf Unit1 = "MACH" Then
      Value = Value * Sqr(1.4 * 298.15 * 8314.51 / 28.95853816) ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Viscosity Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "U" Then
    If Unit1 = "PA-S" Or Unit1 = "KG/M-S" Then
    ElseIf Unit1 = "MPA-S" Then      'Note:  This is milliPa-s, not MPa-s
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "UPA-S" Then
      Value = Value / 1000000 ^ Drct
    ElseIf Unit1 = "G/CM-S" Or Unit1 = "POISE" Then
      Value = Value / 10 ^ Drct
    ElseIf Unit1 = "CENTIPOISE" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "MILLIPOISE" Or Unit1 = "MPOISE" Then
      Value = Value / 10000 ^ Drct
    ElseIf Unit1 = "MICROPOISE" Or Unit1 = "UPOISE" Then
      Value = Value / 10000000 ^ Drct
    ElseIf Unit1 = "LBM/FT-S" Or Unit1 = "LB/FT-S" Then
      Value = Value * (LBMtoKG / FTtoM) ^ Drct
    ElseIf Unit1 = "LBF-S/FT^2" Then
      Value = Value * (LBFtoN / FTtoM ^ 2) ^ Drct
    ElseIf Unit1 = "LBM/FT-H" Or Unit1 = "LB/FT-H" Then
      Value = Value * (LBMtoKG / FTtoM / HtoS) ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Thermal Conductivity Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "K" Then
    If Unit1 = "MW/M-K" Then
    ElseIf Unit1 = "W/M-K" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "G-CM/S^3-K" Then
      Value = Value / 100 ^ Drct
    ElseIf Unit1 = "KG-M/S^3-K" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "CAL/S-CM-K" Then
      Value = Value * (CALtoJ * 100000) ^ Drct
    ElseIf Unit1 = "KCAL/HR-M-K" Then
      Value = Value * (CALtoJ * 100000 * 1000 / 100 / 3600) ^ Drct
    ElseIf Unit1 = "LBM-FT/S^3-F" Or Unit1 = "LB-FT/S^3-F" Then
      Value = Value * (1000 * LBMtoKG * FTtoM / RtoK) ^ Drct
    ElseIf Unit1 = "LBF/S-F" Then
      Value = Value * (1000 * LBFtoN / RtoK) ^ Drct
    ElseIf Unit1 = "BTU/H-FT-F" Then
      Value = Value * (1000 * BTUtoW / HtoS / FTtoM / RtoK) ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Joule-Thomson Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "JT" Then
    If Unit1 = "K/MPA" Or Unit1 = "C/MPA" Then
    ElseIf Unit1 = "K/KPA" Or Unit1 = "C/KPA" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "K/PA" Or Unit1 = "C/PA" Then
      Value = Value * 1000000 ^ Drct
    ElseIf Unit1 = "C/ATM" Then
      Value = Value / ATMtoMPa ^ Drct
    ElseIf Unit1 = "C/BAR" Then
      Value = Value / BARtoMPA ^ Drct
    ElseIf Unit1 = "K/PSI" Or Unit1 = "K/PSIA" Then
      Value = Value / PSIAtoMPA ^ Drct
    ElseIf Unit1 = "F/PSI" Or Unit1 = "F/PSIA" Or Unit1 = "R/PSIA" Then
      Value = Value / (PSIAtoMPA / RtoK) ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Length Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "L" Then
    If Unit1 = "METER" Or Unit1 = "M" Then
    ElseIf Unit1 = "DM" Then
      Value = Value / 10 ^ Drct
    ElseIf Unit1 = "CM" Then
      Value = Value / 100 ^ Drct
    ElseIf Unit1 = "MM" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "KM" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "INCH" Or Unit1 = "IN" Then
      Value = Value * INtoM ^ Drct
    ElseIf Unit1 = "FOOT" Or Unit1 = "FT" Then
      Value = Value * FTtoM ^ Drct
    ElseIf Unit1 = "YARD" Or Unit1 = "YD" Then
      Value = Value * (INtoM * 36) ^ Drct
    ElseIf Unit1 = "MILE" Or Unit1 = "MI" Then
      Value = Value * (INtoM * 63360) ^ Drct
    ElseIf Unit1 = "LIGHT YEAR" Then
      Value = Value * 9.46055E+15 ^ Drct
    ElseIf Unit1 = "ANGSTROM" Then
      Value = Value / 10000000000# ^ Drct
    ElseIf Unit1 = "FATHOM" Then
      Value = Value * (FTtoM * 6) ^ Drct
    ElseIf Unit1 = "MIL" Then
      Value = Value * (INtoM / 1000) ^ Drct
    ElseIf Unit1 = "ROD" Then
      Value = Value * (INtoM * 16.5 * 12) ^ Drct
    ElseIf Unit1 = "PARSEC" Then
      Value = Value * (30837400000000# * 1000) ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Area Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "A" Then
    If Unit1 = "METER^2" Or Unit1 = "M^2" Then
    ElseIf Unit1 = "CM^2" Then
      Value = Value / 10000 ^ Drct
    ElseIf Unit1 = "MM^2" Then
      Value = Value / 1000000 ^ Drct
    ElseIf Unit1 = "KM^2" Then
      Value = Value * 1000000 ^ Drct
    ElseIf Unit1 = "INCH^2" Or Unit1 = "IN^2" Then
      Value = Value * (INtoM ^ 2) ^ Drct
    ElseIf Unit1 = "FOOT^2" Or Unit1 = "FT^2" Then
      Value = Value * (FTtoM ^ 2) ^ Drct
    ElseIf Unit1 = "YARD^2" Or Unit1 = "YD^2" Then
      Value = Value * ((INtoM * 36) ^ 2) ^ Drct
    ElseIf Unit1 = "MILE^2" Or Unit1 = "MI^2" Then
      Value = Value * ((INtoM * 63360) ^ 2) ^ Drct
    ElseIf Unit1 = "ACRE" Then
      Value = Value * ((INtoM * 36) ^ 2 * 4840) ^ Drct
    ElseIf Unit1 = "BARN" Then
      Value = Value * 1E-28 ^ Drct
    ElseIf Unit1 = "HECTARE" Then
      Value = Value * 10000 ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Volume Conversion (Note: not specific volume)
'-----------------------------------------------------------------------
  ElseIf Tpe = "V" Then
    If Unit1 = "METER^3" Or Unit1 = "M^3" Then
    ElseIf Unit1 = "CM^3" Then
      Value = Value / 1000000 ^ Drct
    ElseIf Unit1 = "LITER" Or Unit1 = "L" Or Unit1 = "DM^3" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "INCH^3" Or Unit1 = "IN^3" Then
      Value = Value * IN3toM3 ^ Drct
    ElseIf Unit1 = "FOOT^3" Or Unit1 = "FT^3" Then
      Value = Value * (IN3toM3 * 12 ^ 3) ^ Drct
    ElseIf Unit1 = "YARD^3" Or Unit1 = "YD^3" Then
      Value = Value * (IN3toM3 * 36 ^ 3) ^ Drct
    ElseIf Unit1 = "GALLON" Or Unit1 = "GAL" Then
      Value = Value * GALLONtoM3 ^ Drct
    ElseIf Unit1 = "QUART" Or Unit1 = "QT" Then
      Value = Value * (GALLONtoM3 / 4) ^ Drct
    ElseIf Unit1 = "PINT" Or Unit1 = "PT" Then
      Value = Value * (GALLONtoM3 / 8) ^ Drct
    ElseIf Unit1 = "CUP" Then
      Value = Value * (GALLONtoM3 / 16) ^ Drct
    ElseIf Unit1 = "OUNCE" Then
      Value = Value * (GALLONtoM3 / 128) ^ Drct
    ElseIf Unit1 = "TABLESPOON" Or Unit1 = "TBSP" Then
      Value = Value * (GALLONtoM3 / 256) ^ Drct
    ElseIf Unit1 = "TEASPOON" Or Unit1 = "TSP" Then
      Value = Value * (GALLONtoM3 / 768) ^ Drct
    ElseIf Unit1 = "CORD" Then
      Value = Value * (FT3toM3 * 128) ^ Drct
    ElseIf Unit1 = "BARREL" Then
      Value = Value * (GALLONtoM3 * 42) ^ Drct
    ElseIf Unit1 = "BOARD FOOT" Then
      Value = Value * (IN3toM3 * 144) ^ Drct
    ElseIf Unit1 = "BUSHEL" Then
      Value = Value * 0.03523907016688 ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Mass Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "M" Then
    If Unit1 = "KG" Then
    ElseIf Unit1 = "G" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "MG" Then            'milligram
      Value = Value / 1000000 ^ Drct
    ElseIf Unit1 = "LBM" Or Unit1 = "LB" Then
      Value = Value * LBMtoKG ^ Drct
    ElseIf Unit1 = "GRAIN" Then
      Value = Value * (LBMtoKG / 7000) ^ Drct
    ElseIf Unit1 = "SLUG" Then
      Value = Value * (KGFtoN * LBMtoKG / FTtoM) ^ Drct
    ElseIf Unit1 = "TON" Then
      Value = Value * (LBMtoKG * 2000) ^ Drct
    ElseIf Unit1 = "TONNE" Then
      Value = Value * 1000 ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Force Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "F" Then
    If Unit1 = "NEWTON" Or Unit1 = "N" Then
    ElseIf Unit1 = "MN" Then 'milliNewtons
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "KGF" Then
      Value = Value * KGFtoN ^ Drct
    ElseIf Unit1 = "DYNE" Then
      Value = Value / 100000 ^ Drct
    ElseIf Unit1 = "LBF" Then
      Value = Value * LBFtoN ^ Drct
    ElseIf Unit1 = "POUNDAL" Then
      Value = Value * (LBMtoKG * FTtoM) ^ Drct
    ElseIf Unit1 = "OZF" Then
      Value = Value * (LBFtoN / 16) ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Energy Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "E" Then
    If Unit1 = "JOULE" Or Unit1 = "J" Then
    ElseIf Unit1 = "KJ" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "MJ" Then
      Value = Value * 1000000 ^ Drct
    ElseIf Unit1 = "KW-H" Then
      Value = Value * (HtoS * 1000) ^ Drct
    ElseIf Unit1 = "CAL" Then
      Value = CALtoJ ^ Drct * Value
    ElseIf Unit1 = "KCAL" Then
      Value = Value * (CALtoJ * 1000) ^ Drct
    ElseIf Unit1 = "ERG" Then
      Value = Value / 10000000 ^ Drct
    ElseIf Unit1 = "BTU" Then
      Value = Value * (BTUtoKJ * 1000) ^ Drct
    ElseIf Unit1 = "FT-LBF" Then
      Value = Value * FTLBFtoJ ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Power Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "Q" Then
    If Unit1 = "WATT" Or Unit1 = "W" Then
    ElseIf Unit1 = "KWATT" Or Unit1 = "KW" Then
      Value = Value * 1000 ^ Drct
    ElseIf Unit1 = "BTU/S" Then
      Value = Value * BTUtoW ^ Drct
    ElseIf Unit1 = "BTU/MIN" Then
      Value = Value * (BTUtoW / 60) ^ Drct
    ElseIf Unit1 = "BTU/H" Then
      Value = Value * (BTUtoW / HtoS) ^ Drct
    ElseIf Unit1 = "CAL/S" Then
      Value = Value * CALtoJ ^ Drct
    ElseIf Unit1 = "KCAL/S" Then
      Value = Value * (CALtoJ * 1000) ^ Drct
    ElseIf Unit1 = "CAL/MIN" Then
      Value = Value * (CALtoJ / 60) ^ Drct
    ElseIf Unit1 = "KCAL/MIN" Then
      Value = Value * (CALtoJ / 60 * 1000) ^ Drct
    ElseIf Unit1 = "FT-LBF/S" Then
      Value = Value * FTLBFtoJ ^ Drct
    ElseIf Unit1 = "FT-LBF/MIN" Then
      Value = Value * (FTLBFtoJ / 60) ^ Drct
    ElseIf Unit1 = "FT-LBF/H" Then
      Value = Value * (FTLBFtoJ / HtoS) ^ Drct
    ElseIf Unit1 = "HP" Then
      Value = Value * HPtoW ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If

'-----------------------------------------------------------------------
'   Surface Tension Conversion
'-----------------------------------------------------------------------
  ElseIf Tpe = "N" Then
    If Unit1 = "N/M" Then
    ElseIf Unit1 = "MN/M" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "DYNE/CM" Or Unit1 = "DYN/CM" Then
      Value = Value / 1000 ^ Drct
    ElseIf Unit1 = "LBF/FT" Then
      Value = Value * LBFTtoNM ^ Drct
    Else
      UnitConvert = Trim2("Undefined input unit"): Exit Function
    End If
  End If
  Unit1 = Unit2
Next Drct
UnitConvert = Value
End Function

Sub SetupUnits(i)

'Warning:  If any of these are changed (to make them the default) after the program has run,
'  you will need to exit Excel and restart it so that it reinitializes

'REFPROP Units
  tUnits2 = "K"
  taUnits2 = "K"
  pUnits2 = "kPa"
  dUnits2 = "mol/dm^3"
  vUnits2 = "dm^3/mol"
  hUnits2 = "J/mol"
  sUnits2 = "J/mol-K"
  wUnits2 = "m/s"
  visUnits2 = "uPa-s"
  tcxUnits2 = "W/m-K"
  stUnits2 = "N/m"
  tmUnits2 = "s"
'Default units: (SI)
  tUnits(0) = "K"
  taUnits(0) = "K"
  pUnits(0) = "MPa"
  dUnits(0) = "kg/m^3"
  vUnits(0) = "m^3/kg"
  hUnits(0) = "kJ/kg"
  sUnits(0) = "kJ/kg-K"
  wUnits(0) = "m/s"
  visUnits(0) = "uPa-s"
  tcxUnits(0) = "mW/m-K"
  stUnits(0) = "mN/m"
  tmUnits(0) = "s"
'Default units but with K switched to C (SI with C) (SIwithC) or (C)
  tUnits(5) = "C"
  taUnits(5) = "K"
  pUnits(5) = "MPa"
  dUnits(5) = "kg/m^3"
  vUnits(5) = "m^3/kg"
  hUnits(5) = "kJ/kg"
  sUnits(5) = "kJ/kg-K"
  wUnits(5) = "m/s"
  visUnits(5) = "uPa-s"
  tcxUnits(5) = "mW/m-K"
  stUnits(5) = "mN/m"
  tmUnits(5) = "s"
'Default units on a molar basis (Molar SI)
  tUnits(6) = "K"
  taUnits(6) = "K"
  pUnits(6) = "MPa"
  dUnits(6) = "mol/dm^3"
  vUnits(6) = "dm^3/mol"
  hUnits(6) = "J/mol"
  sUnits(6) = "J/mol-K"
  wUnits(6) = "m/s"
  visUnits(6) = "uPa-s"
  tcxUnits(6) = "mW/m-K"
  stUnits(6) = "mN/m"
  tmUnits(6) = "s"
'mks (mks)
  tUnits(1) = "K"
  taUnits(1) = "K"
  pUnits(1) = "kPa"
  dUnits(1) = "kg/m^3"
  vUnits(1) = "m^3/kg"
  hUnits(1) = "kJ/kg"
  sUnits(1) = "kJ/kg-K"
  wUnits(1) = "m/s"
  visUnits(1) = "uPa-s"
  tcxUnits(1) = "W/m-K"
  stUnits(1) = "mN/m"
  tmUnits(1) = "s"
'cgs (cgs)
  tUnits(2) = "K"
  taUnits(2) = "K"
  pUnits(2) = "MPa"
  dUnits(2) = "g/cm^3"
  vUnits(2) = "cm^3/g"
  hUnits(2) = "J/g"
  sUnits(2) = "J/g-K"
  wUnits(2) = "cm/s"
  visUnits(2) = "uPa-s"
  tcxUnits(2) = "mW/m-K"
  stUnits(2) = "dyn/cm"
  tmUnits(2) = "s"
'English (E)
  tUnits(3) = "F"             'See comments above
  taUnits(3) = "R"
  pUnits(3) = "psia"
  dUnits(3) = "lbm/ft^3"
  vUnits(3) = "ft^3/lbm"
  hUnits(3) = "Btu/lbm"
  sUnits(3) = "Btu/lbm-R"
  wUnits(3) = "ft/s"
  visUnits(3) = "lbm/ft-s"
  tcxUnits(3) = "Btu/h-ft-F"
  stUnits(3) = "lbf/ft"
  tmUnits(3) = "s"
'Molar English (molar E)
  tUnits(7) = "F"
  taUnits(7) = "R"
  pUnits(7) = "psia"
  dUnits(7) = "lbmol/ft^3"
  vUnits(7) = "ft^3/lbmol"
  hUnits(7) = "Btu/lbmol"
  sUnits(7) = "Btu/lbmol-R"
  wUnits(7) = "ft/s"
  visUnits(7) = "lbm/ft-s"
  tcxUnits(7) = "Btu/h-ft-F"
  stUnits(7) = "lbf/ft"
  tmUnits(7) = "s"
'Mixed (Mixed)
  tUnits(4) = "K"
  taUnits(4) = "K"
  pUnits(4) = "psia"
  dUnits(4) = "g/cm^3"
  vUnits(4) = "cm^3/g"
  hUnits(4) = "J/g"
  sUnits(4) = "J/g-K"
  wUnits(4) = "m/s"
  visUnits(4) = "uPa-s"
  tcxUnits(4) = "mW/m-K"
  stUnits(4) = "mN/m"
  tmUnits(4) = "s"
'Mechanical Engineering units (MEUNITS)
  tUnits(9) = "C"
  taUnits(9) = "K"
  pUnits(9) = "bar"
  dUnits(9) = "g/cm^3"
  vUnits(9) = "cm^3/g"
  hUnits(9) = "J/g"
  sUnits(9) = "J/g-K"
  wUnits(9) = "cm/s"
  visUnits(9) = "centipoise"
  tcxUnits(9) = "mW/m-K"
  tmUnits(9) = "s"

End Sub

Function ConvertUnits(InpCode, Units, Prop1, Prop2)
Dim i As Integer, at As String, bt As String, tConv As Double

If IsMissing(InpCode) Then InpCode = ""
If IsMissing(Units) Then Units = ""
If IsMissing(Prop1) Then Prop1 = 0
If IsMissing(Prop2) Then Prop2 = 0
If ierr > 0 Or ierr = -16 Then ConvertUnits = Trim2(herr): Exit Function
If tUnits2 = "" Then
  Call SetupUnits(0)  'If Default units are changed, this needs to be called again.  Normally it is skipped after the first entry
End If


'Change the 0 in the following line to 3 for default English units, 1 for mks, or 2 for cgs, etc.

DefaultUnits = 0
i = -1

'Do not change the order of the next block of statements
If UCase(Units) = "SI" Then i = 0                       'SI
If UCase(Units) = "SI WITH C" Or UCase(Units) = "SIWITHC" Then i = 5    'SI with C
If UCase(Units) = "C" Then i = 5                                 'SI with C
If Left(UCase(Units), 3) = "MIX" Then i = 4                      'Mixed
If UCase(Units) = "MOLAR SI" Then i = 6                          'Molar SI
If UCase(Units) = "MOLARSI" Then i = 6                           'Molar SI
If UCase(Units) = "MKS" Then i = 1                               'mks
If UCase(Units) = "CGS" Then i = 2                               'cgs
If Left(UCase(Units), 1) = "E" Then i = 3                        'English
If UCase(Units) = "MOLAR E" Then i = 7                           'Molar English
If UCase(Units) = "MEUNITS" Then i = 9                           'Mechanical Engineering units

If i = -1 Then
  If Units = "" Then
    i = DefaultUnits
  Else
    ierr = 1
    herr = "Invalid Units"
    Exit Function
  End If
End If
DefUnit = i


at = UCase(Left(InpCode, 1))
bt = UCase(Mid(InpCode, 2, 1))
molmass = 1
If InStr(LCase(dUnits(i)), "mol") = 0 Then molmass = 2

If at = "-" Then
  ConvertUnits = Prop1
  If Prop1 >= -9999999 And Prop1 <= -9999900 Then
    If Prop1 = CLng(Prop1) Then
      ConvertUnits = Trim2("Undefined")
      Exit Function
    End If
  End If
  'If Len(Trim(Prop1)) > 0 Then
    If bt = "T" Then ConvertUnits = UnitConvert(Prop1, "T", tUnits2, tUnits(i))
    If bt = "A" Then ConvertUnits = UnitConvert(Prop1, "T", taUnits2, taUnits(i))
    If bt = "P" Then ConvertUnits = UnitConvert(Prop1, "P", pUnits2, pUnits(i))
    If bt = "D" Then ConvertUnits = UnitConvert(Prop1, "D", dUnits2, dUnits(i))
    If bt = "V" Then ConvertUnits = UnitConvert(Prop1, "D", vUnits2, vUnits(i))
    If bt = "H" Or bt = "E" Then ConvertUnits = UnitConvert(Prop1, "H", hUnits2, hUnits(i))
    If bt = "S" Then ConvertUnits = UnitConvert(Prop1, "S", sUnits2, sUnits(i))
    If bt = "W" Then ConvertUnits = UnitConvert(Prop1, "W", wUnits2, wUnits(i))
    If bt = "U" Then ConvertUnits = UnitConvert(Prop1, "U", visUnits2, visUnits(i))
    If bt = "K" Then ConvertUnits = UnitConvert(Prop1, "K", tcxUnits2, tcxUnits(i))
    If bt = "N" Then ConvertUnits = UnitConvert(Prop1, "N", stUnits2, stUnits(i))
  'End If
  If bt = "J" Then
    tConv = 1
    If tUnits(i) = "R" Or tUnits(i) = "F" Then tConv = 1 / RtoK
    ConvertUnits = Prop1 * tConv / UnitConvert(1, "P", "kPa", pUnits(i))
  End If
  If bt = "I" Then
    'convert cm^2/s to ft^2/s
    If i = 3 Or i = 7 Then
      ConvertUnits = Prop1 * UnitConvert(1, "A", "cm^2", "ft^2")
    End If
  End If
Else
  If Len(Trim(Prop1)) > 0 Then
    If at = "T" Then Prop1 = UnitConvert(Prop1, "T", tUnits(i), tUnits2)
    If at = "A" Then Prop1 = UnitConvert(Prop1, "T", taUnits(i), taUnits2)
    If at = "P" Then Prop1 = UnitConvert(Prop1, "P", pUnits(i), pUnits2)
    If at = "D" Then Prop1 = UnitConvert(Prop1, "D", dUnits(i), dUnits2)
    If at = "V" Then Prop1 = UnitConvert(Prop1, "D", vUnits(i), vUnits2)
    If at = "H" Or at = "E" Then Prop1 = UnitConvert(Prop1, "H", hUnits(i), hUnits2)
    If at = "S" Then Prop1 = UnitConvert(Prop1, "S", sUnits(i), sUnits2)
    If at = "W" Then Prop1 = UnitConvert(Prop1, "W", wUnits(i), wUnits2)
    If at = "U" Then Prop1 = UnitConvert(Prop1, "U", visUnits(i), visUnits2)
    If at = "K" Then Prop1 = UnitConvert(Prop1, "K", tcxUnits(i), tcxUnits2)
    If at = "N" Then Prop1 = UnitConvert(Prop1, "N", stUnits(i), stUnits2)
  End If

  If Len(Trim(Prop2)) > 0 Then
    If bt = "T" Then Prop2 = UnitConvert(Prop2, "T", tUnits(i), tUnits2)
    If bt = "A" Then Prop2 = UnitConvert(Prop2, "T", taUnits(i), taUnits2)
    If bt = "P" Then Prop2 = UnitConvert(Prop2, "P", pUnits(i), pUnits2)
    If bt = "D" Then Prop2 = UnitConvert(Prop2, "D", dUnits(i), dUnits2)
    If bt = "V" Then Prop2 = UnitConvert(Prop2, "D", vUnits(i), vUnits2)
    If bt = "H" Or bt = "E" Then Prop2 = UnitConvert(Prop2, "H", hUnits(i), hUnits2)
    If bt = "S" Then Prop2 = UnitConvert(Prop2, "S", sUnits(i), sUnits2)
    If bt = "W" Then Prop2 = UnitConvert(Prop2, "W", wUnits(i), wUnits2)
    If bt = "U" Then Prop2 = UnitConvert(Prop2, "U", visUnits(i), visUnits2)
    If bt = "K" Then Prop2 = UnitConvert(Prop2, "K", tcxUnits(i), tcxUnits2)
    If bt = "N" Then Prop2 = UnitConvert(Prop2, "N", stUnits(i), stUnits2)
  End If
End If
End Function

Function FluidString(Nmes, Comps, Optional massmole As String) As String
  Dim a As String, i As Integer, ncalc As Integer, sum As Double
  ncalc = 0
  If Nmes.Count <> Comps.Count Then
    If Nmes.Count <> Comps.Count + 1 Then
      FluidString = "Number of fluid names and compositions not the same": Exit Function
    Else
      ncalc = 1  'Calculate missing composition (in last spot only)
    End If
  End If
  a = ""
  sum = 0
  For i = 1 To Nmes.Count
    If Nmes(i) <> "" Then
      If i = Nmes.Count And ncalc = 1 Then
        If sum < 0 Or sum > 1 Then FluidString = "Sum must be less than 1 to calculate final composition.": Exit Function
        a = a & Nmes(i) & ";" & (1 - sum) & ";"   'Compositions must be given in fractions, not percent
      ElseIf Comps(i) > 0 Then
        sum = sum + Comps(i)
        a = a & Nmes(i) & ";" & Comps(i) & ";"
      End If
    End If
  Next
  Do While Right(a, 1) = ";"
    a = Left(a, Len(a) - 1)
  Loop
  FluidString = a
  If massmole = "" Or LCase(Left(massmole, 3)) = "mol" Or LCase(Right(massmole, 3)) = "mol" Then
    FluidString = a
  Else
    FluidString = a & " mass"
  End If
End Function

Sub ChDirUNC(ByVal sPath As String)
  If Left(sPath, 2) = "\\" Then
    'Change to a UNC Directory
    Dim lReturn As Long
    'Call the API function to set the current directory
    lReturn = SetCurDir(sPath)
    'A zero return value means an error
    If lReturn = 0 Then
        Err.Raise vbObjectError + 1, "Error setting path. In Excel, under Tools, Options, General, change your 'Default File Location' to a local directory."
        Exit Sub
    End If
  Else
    ChDrive sPath
    ChDir sPath
  End If
End Sub

Function WorkBookName()
  WorkBookName = ThisWorkbook.FullName       'returns Path+Name of this file
End Function

Function WhereIsWorkbook()
    WhereIsWorkbook = ActiveWorkbook.FullName   'returns Path+Name of "Present and Active" File.
End Function

Function WhereAreREFPROPfunctions()
    WhereAreREFPROPfunctions = "Using REFPROP.xla functions"
End Function

Function SeeFileLinkSources()
    SeeFileLinkSources = ActiveWorkbook.LinkSources
End Function

Function PropertyUnits(InpCode, Units)
  Dim i As Integer
  Dim d1 As Double, d2 As Double
  Dim mmunits As String, molunits As String, massunits As String, kvunits As String
  Dim eunits As String, volunits As String, distanceunits As String
  PropertyUnits = ""
  ierr = 0
  d1 = 1
  d2 = 1
  Call ConvertUnits("PD", Units, d1, d2)
  If UCase(InpCode) = "T" Then PropertyUnits = tUnits(DefUnit)
  If UCase(InpCode) = "P" Then PropertyUnits = pUnits(DefUnit)
  If UCase(InpCode) = "D" Then PropertyUnits = dUnits(DefUnit)
  If UCase(InpCode) = "V" Then PropertyUnits = vUnits(DefUnit)
  If UCase(InpCode) = "H" Then PropertyUnits = hUnits(DefUnit)
  If UCase(InpCode) = "E" Then PropertyUnits = hUnits(DefUnit)
  If UCase(InpCode) = "S" Then PropertyUnits = sUnits(DefUnit)
  If UCase(InpCode) = "W" Then PropertyUnits = wUnits(DefUnit)
  If UCase(InpCode) = "U" Then PropertyUnits = visUnits(DefUnit)
  If UCase(InpCode) = "K" Then PropertyUnits = tcxUnits(DefUnit)
  If UCase(InpCode) = "N" Then PropertyUnits = stUnits(DefUnit)
  If UCase(InpCode) = "Z" Then PropertyUnits = "-"
  If UCase(InpCode) = "B" Then PropertyUnits = vUnits(DefUnit)

  massunits = "g"
  molunits = "mol"
  kvunits = "cm^2/s"
  If Left(dUnits(DefUnit), 2) = "lb" Then
    massunits = "lbm"
    molunits = "lbmol"
    kvunits = "ft^2/s"
  End If
  mmunits = molunits
  If molmass = 2 Then mmunits = massunits
  i = InStr(hUnits(DefUnit), "/")
  If i Then eunits = Left(hUnits(DefUnit), i - 1)
  i = InStr(vUnits(DefUnit), "/")
  If i Then volunits = Left(vUnits(DefUnit), i - 1)
  i = InStr(volunits, "^")
  If i Then distanceunits = Left(volunits, i - 1)
  If UCase(InpCode) = "TEMPERATURE" Then PropertyUnits = tUnits(DefUnit)
  If UCase(InpCode) = "PRESSURE" Then PropertyUnits = pUnits(DefUnit)
  If UCase(InpCode) = "DENSITY" Then PropertyUnits = dUnits(DefUnit)
  If UCase(InpCode) = "VOLUME" Then PropertyUnits = vUnits(DefUnit)
  If UCase(InpCode) = "ENTHALPY" Then PropertyUnits = hUnits(DefUnit)
  If UCase(InpCode) = "ENERGY" Then PropertyUnits = hUnits(DefUnit)
  If UCase(InpCode) = "ENTROPY" Then PropertyUnits = sUnits(DefUnit)
  If UCase(InpCode) = "SPEED OF SOUND" Then PropertyUnits = wUnits(DefUnit)
  If UCase(InpCode) = "SECOND VIRIAL" Then PropertyUnits = vUnits(DefUnit)
  If UCase(InpCode) = "VISCOSITY" Then PropertyUnits = visUnits(DefUnit)
  If UCase(InpCode) = "THERMAL CONDUCTIVITY" Then PropertyUnits = tcxUnits(DefUnit)
  If UCase(InpCode) = "KINEMATIC VISCOSITY" Then PropertyUnits = kvunits
  If UCase(InpCode) = "THERMAL DIFFUSIVITY" Then PropertyUnits = kvunits
  If UCase(InpCode) = "SURFACE TENSION" Then PropertyUnits = stUnits(DefUnit)
  If UCase(InpCode) = "EXPANSIVITY" Then PropertyUnits = "1/" + taUnits(DefUnit)
  If UCase(InpCode) = "COMPRESSIBILITY" Then PropertyUnits = "1/" + pUnits(DefUnit)
  If UCase(InpCode) = "JOULE THOMSON" Then PropertyUnits = taUnits(DefUnit) + "/" + pUnits(DefUnit)
  If UCase(InpCode) = "DPDD" Then PropertyUnits = pUnits(DefUnit) + "/(" + dUnits(DefUnit) + ")"
  If UCase(InpCode) = "DPDD2" Then PropertyUnits = pUnits(DefUnit) + "/(" + dUnits(DefUnit) + ")^2"
  If UCase(InpCode) = "DPDT" Then PropertyUnits = pUnits(DefUnit) + "/" + taUnits(DefUnit)
  If UCase(InpCode) = "DDDT" Then PropertyUnits = "(" + dUnits(DefUnit) + ")/" + taUnits(DefUnit)
  If UCase(InpCode) = "DHDT" Then PropertyUnits = "(" + hUnits(DefUnit) + ")/" + taUnits(DefUnit)
  If UCase(InpCode) = "DHDP" Then PropertyUnits = "(" + hUnits(DefUnit) + ")/" + pUnits(DefUnit)
  If UCase(InpCode) = "DHDD" Then PropertyUnits = "(" + hUnits(DefUnit) + ")/(" + dUnits(DefUnit) + ")"
  If UCase(InpCode) = "COMPRESSIBILITY FACTOR" Then PropertyUnits = "dimLess"
  If UCase(InpCode) = "COEFFICIENT" Then PropertyUnits = "dimLess"
  If UCase(InpCode) = "PRANDTL" Then PropertyUnits = "dimLess"
  If UCase(InpCode) = "DIELECTRIC CONSTANT" Then PropertyUnits = "dimLess"
  If UCase(InpCode) = "MOLE FRACTION" Then PropertyUnits = molunits + "/" + molunits
  If UCase(InpCode) = "MASS FRACTION" Then PropertyUnits = massunits + "/" + massunits
  If UCase(InpCode) = "MOLAR MASS" Then PropertyUnits = massunits + "/" + molunits
  If Left(UCase(InpCode), 11) = "ENERGY FLOW" Then PropertyUnits = eunits + "/" + tmUnits(DefUnit)
  If Left(UCase(InpCode), 9) = "MASS FLOW" Then PropertyUnits = massunits + "/" + tmUnits(DefUnit)
  If Left(UCase(InpCode), 11) = "VOLUME FLOW" Then PropertyUnits = volunits + "/" + tmUnits(DefUnit)
  If Left(UCase(InpCode), 8) = "DISTANCE" Then PropertyUnits = distanceunits
  If Left(UCase(InpCode), 4) = "TIME" Then PropertyUnits = tmUnits(DefUnit)
  
  'If UCase(InpCode) = "MOLAR QUALITY" Then PropertyUnits = molunits + "/" + molunits
  'If UCase(InpCode) = "MASS QUALITY" Then PropertyUnits = massunits + "/" + massunits
  If UCase(InpCode) = "MOLAR QUALITY" Then PropertyUnits = "MoleFr.Vaporized"
  If UCase(InpCode) = "MASS QUALITY" Then PropertyUnits = "MassFr.Vaporized"
  If UCase(InpCode) = "QUALITY" Then
    PropertyUnits = "MoleFr.Vaporized"
    If molmass = 2 Then PropertyUnits = "MassFr.Vaporized"
  End If
  If UCase(InpCode) = "FLUID STRING" Then PropertyUnits = "name(i),molefract(i)"
End Function

Function SelectedDefaultUnits()
  Dim d1 As Double, d2 As Double, Units As String
  ierr = 0
  d1 = 1
  d2 = 1
  Call ConvertUnits("PD", Units, d1, d2)
  If DefaultUnits = 0 Then SelectedDefaultUnits = "SI"
  If DefaultUnits = 1 Then SelectedDefaultUnits = "MKS"
  If DefaultUnits = 2 Then SelectedDefaultUnits = "CGS"
  If DefaultUnits = 3 Then SelectedDefaultUnits = "E"
  If DefaultUnits = 4 Then SelectedDefaultUnits = "MIXED"
  If DefaultUnits = 5 Then SelectedDefaultUnits = "SI WITH C"
  If DefaultUnits = 6 Then SelectedDefaultUnits = "MOLAR SI"
  If DefaultUnits = 7 Then SelectedDefaultUnits = "MOLAR E"
  If DefaultUnits = 9 Then SelectedDefaultUnits = "MEUNITS"
End Function

Function RefpropXLSVersionNumber()
  RefpropXLSVersionNumber = "9.1103"
End Function

Function RefpropDLLVersionNumber()
  Dim nc2 As Long
  nc2 = -1
  Call SETUPdll(nc2, hfld, hfmix, hrf, ierr, herr, 10000&, 255&, 3&, 255&)
  RefpropDLLVersionNumber = ierr / 10000
  If ierr = 900 Then RefpropDLLVersionNumber = 9
End Function


-------------------------------------------------------------------------------
VBA MACRO WaterBathFunctions.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/WaterBathFunctions'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

Function Hco_film_boiling(Kvapor As Single, VaporDensity As Single, LiquidDensity As Single, h As Single, CpVapor As Single, Tsurf As Single, Tsat As Single, OD As Single, VaporViscosity As Single)
Dim Hc, hr, s, e As Single
s = 1.73E-09      'Stefan-Boltzmann constant
e = 0.3 'emissivity

Hc = (0.62) * (Kvapor ^ 3 * VaporDensity * (LiquidDensity - VaporDensity) * g * 3600 ^ 2 * (h + 0.4 * CpVapor * (Tsurf - Tsat) / (OD / 12 * VaporViscosity * 2.42 * (Tsurf - Tsat)))) ^ 0.25
hr = s * e * (Tsurf ^ 4 - Tsat ^ 4) / (Tsurf - Tsat)

Hco_film_boiling = Hc + hr * (3 / 4 + 0.25 * (hr / Hc) * (1 / (2.62 + (hr / Hc))))

End Function

Public Function BeamLengthCheck(length, L1, L2, L3, L4, L5, S1, S2, S3, S4, S5) As Single

If length <= L1 Then
    BeamLengthCheck = S1 / 12
    ElseIf length > L1 And length <= L2 Then
        BeamLengthCheck = S2 / 12
    ElseIf length > L2 And length <= L3 Then
        BeamLengthCheck = S3 / 12
    ElseIf length > L3 And length <= L4 Then
        BeamLengthCheck = S4 / 12
    ElseIf length > L4 And length <= L5 Then
        BeamLengthCheck = S5 / 12
End If


End Function

Public Function ExponentialDecayParameter(AverageFluxrate As Single, MaxFluxRate As Single, length As Single)

Dim x, b, step As Single
Dim iter, neg As Integer

b = -0.001
step = 0.1
iter = 1
neg = -1
x = MaxFluxRate / b * (Exp(b * length) - 1) / length

Do While Abs((x - AverageFluxrate) / AverageFluxrate) > 0.0001
    x = MaxFluxRate / b * (Exp(b * length) - 1) / length
    b = b + step
    
    If (x - AverageFluxrate) / AverageFluxrate * neg < 0 Then
        step = step / -10
        neg = neg * -1
    End If
    
    iter = iter + 1
    If iter > 100000 Then
        GoTo Err
    End If
    
Loop
ExponentialDecayParameter = b
Exit Function

Err:
ExponentialDecayParameter = "Exceeded Max Iterations"

End Function


Function hco_cond(numberTubes As Integer, LiquidDensity As Single, VaporDensity As Single, LiquidConductivity As Single, LiquidViscosity As Single, diameter As Single, T_Sat As Single, T_surf As Single, LatentHeat As Single)

hco_cond = 0.725 * ((LiquidDensity * g * (LiquidDensity - VaporDensity) * LatentHeat * LiquidConductivity ^ 3) / (LiquidViscosity * 0.000671969 * numberTubes * diameter / 12 * (T_Sat - T_surf) / 3600)) ^ 0.25

End Function

Function Grashoff(beta As Single, LMTD As Single, PipeOD As Single, Density As Single, Viscosity As Single)

'Viscosity in centipoise (cP)

'convert units from cP to lb/ft*s
Viscosity = Viscosity * 2.419 / 3600

Grashoff = g * beta * LMTD * (PipeOD / 12) ^ 3 / (Viscosity / Density) ^ 2

End Function

Function beta(DensityAvg As Single, DensityHighTemp As Single, DensityLowTemp As Single, TempBathHigh As Single, TempBathLow As Single)

'expansion coefficient for natural convection coefficient
'units are °F^-1

beta = DensityAvg * ((1 / DensityHighTemp) - (1 / DensityLowTemp)) / (TempBathHigh - TempBathLow)

End Function

Function hco_nat_conv(Ra As Single, k_bath As Single, OD As Single)

Dim Nu As Single ' Nusselt number


Select Case Ra
    ' laminar flow
    Case Is < 10 ^ 9
        Nu = 0.47 * Ra ^ (1 / 4)
    'turbulent flow
    Case Is > 10 ^ 9
        Nu = 0.13 * Ra ^ (1 / 3)
End Select

hco_nat_conv = k_bath * Nu / (OD / 12)

End Function

Function Hco_nucleate_boiling(visc_bath As Single, h As Single, DensAvg As Single, DensVapor As Single, Cp_Bath As Single, Tsurf As Single, Tsat As Single, Pr As Single, Csf As Single)
Dim SurfaceTension As Single

SurfaceTension = 0.00528 * (1 - 0.0013 * Tsat)

Hco_nucleate_boiling = visc_bath * 2.419 * h * ((DensAvg - DensVapor) / SurfaceTension) ^ 0.5 * ((Cp_Bath * (Tsurf - Tsat)) / (h * Pr * Csf)) ^ 3 / (Tsurf - Tsat)

End Function



Function ShellDia(coilSize As Single, numberTubes As Integer, FireTubeSize As Single, numberPasses As Integer) As Single

Dim coilArea, firetubeArea, EstimatedArea As Single
Dim diameter As Single
Dim StandardDia As Variant
Dim i As Integer

StandardDia = Array(24, 36, 42, 48, 54, 60, 66, 72, 78, 84, 90, 96, 102, 108, 114, 120, 126, 132, 138, 144, 0)

coilArea = 4 * coilSize ^ 2 * (numberTubes - 2)
firetubeArea = 4 * FireTubeSize ^ 2 * (numberPasses - 2)
If coilArea > firetubeArea Then
    EstimatedArea = coilArea
Else
    EstimatedArea = firetubeArea
End If

diameter = Sqr(EstimatedArea * 8 / pi)
i = 0

Do While StandardDia(i) <> 0
    If StandardDia(i) > diameter Then
        ShellDia = StandardDia(i)
        Exit Do
    End If
    i = i + 1
Loop

End Function

Function shellDia2(coilSize As Single, numberTubes As Integer, chamberSize As Single, returnSize As Single, numberPasses As Integer) As Single

' for determining shell size when firetube is u-tube or serpentine

Dim standardDiameters As Variant
standardDiameters = Array(24, 36, 42, 48, 54, 60, 66, 72, 78, 84, 90, 96, 102, 108, 114, 120, 126, 132, 138, 144, 0)

Dim coilArea, estDiameter1 As Single
coilArea = (numberTubes - 2) * Sqr(3) * coilSize ^ 2
estDiameter1 = Sqr(4 * coilArea / pi)

Dim firetubeArea, estDiameter2 As Single

If chamberSize = returnSize And numberPasses = 2 Then
    estDiameter2 = chamberSize * 3

ElseIf chamberSize = returnSize And numberPasses > 2 Then
    firetubeArea = (numberPasses - 2) * Sqr(3) * chamberSize ^ 2
    estDiameter2 = Sqr(4 * firetubeArea / pi)

ElseIf chamberSize <> returnSize And numberPasses = 2 Then
    estDiameter2 = (0.5 * chamberSize) + (2.5 * returnSize)

ElseIf chamberSize <> returnSize And numberPasses > 2 Then
    firetubeArea = (Sqr(3) * (chamberSize ^ 2)) + ((numberPasses - 3) * Sqr(3) * (returnSize ^ 2))
    estDiameter2 = Sqr(4 * firetubeArea / pi)

End If

Dim shellDiameter As Single
If estDiameter1 > estDiameter2 Then
    shellDiameter = estDiameter1
Else
    shellDiameter = estDiameter2
End If

Dim i As Integer
i = 0
Do While standardDiameters(i) <> 0
    If standardDiameters(i) > shellDiameter Then
        shellDia2 = standardDiameters(i)
        Exit Do
    End If
    i = i + 1
Loop

End Function

Function shellDia3(coilSize As Single, numberTubes As Integer, chamberSize As Single, returnSize As Single, numReturns As Integer) As Single

' for determining shell size when firetube is SFMR
Dim standardDiameters As Variant
standardDiameters = Array(24, 36, 42, 48, 54, 60, 66, 72, 78, 84, 90, 96, 102, 108, 114, 120, 126, 132, 138, 144, 0)

Dim coilArea, estDiameter1 As Single
coilArea = (numberTubes - 2) * Sqr(3) * coilSize ^ 2
estDiameter1 = Sqr(4 * coilArea / pi)

Dim firetubeArea, estDiameter2 As Single
firetubeArea = (3 * Sqr(3) * chamberSize ^ 2) + ((numReturns - 4) * Sqr(3) * returnSize ^ 2)
estDiameter2 = Sqr(4 * firetubeArea / pi)

Dim shellDiameter As Single
If estDiameter1 > estDiameter2 Then
    shellDiameter = estDiameter1
Else
    shellDiameter = estDiameter2
End If

Dim i As Integer
i = 0
Do While standardDiameters(i) <> 0
    If standardDiameters(i) > shellDiameter Then
        shellDia3 = standardDiameters(i)
        Exit Do
    End If
    i = i + 1
Loop

End Function




+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|Suspicious|Environ             |May read system environment variables        |
|Suspicious|environment         |May read system environment variables        |
|Suspicious|open                |May open a file                              |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
|Suspicious|Lib                 |May run code from a DLL                      |
|Suspicious|Chr                 |May attempt to obfuscate specific strings    |
|          |                    |(use option --deobf to deobfuscate)          |
|Suspicious|System              |May run an executable file or a system       |
|          |                    |command on a Mac (if combined with           |
|          |                    |libc.dylib)                                  |
|Suspicious|Hex Strings         |Hex-encoded strings were detected, may be    |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|Suspicious|Base64 Strings      |Base64-encoded strings were detected, may be |
|          |                    |used to obfuscate strings (option --decode to|
|          |                    |see all)                                     |
|IOC       |REFPRP64.DLL        |Executable file name                         |
|Hex String|Cv3&                |43763326                                     |
|Hex String|Uui)                |55756929                                     |
|Hex String|`UCq                |60554371                                     |
|Hex String|b@e!                |62406521                                     |
|Hex String|pUAt                |70554174                                     |
|Hex String|xp`%                |78706025                                     |
|Hex String|G#$p                |47232470                                     |
|Hex String|ta4"                |74613422                                     |
|Hex String|C rA                |43207241                                     |
|Hex String|i11                 |69313109                                     |
|Hex String|8@FD                |38404644                                     |
|Hex String|eAFR                |65414652                                     |
|Hex String|gf(e                |67662865                                     |
|Hex String|DA3S                |44413353                                     |
|Hex String|R 7)                |52203729                                     |
|Hex String|bCC)                |62434329                                     |
|Hex String|4       "           |34200922                                     |
|Hex String|#HPa                |23485061                                     |
|Hex String|QP4q                |51503471                                     |
|Hex String|hv))                |68762929                                     |
|Hex String|wV6r                |77563672                                     |
|Hex String|7sr!                |37737221                                     |
|Hex String|'S6P                |27533650                                     |
|Hex String|dtB1                |64744231                                     |
|Hex String|2DQy                |32445179                                     |
|Hex String|U')p                |55272970                                     |
|Hex String|iie&                |69696526                                     |
|Hex String|8%@7                |38254037                                     |
|Hex String|BVAH                |42564148                                     |
|Hex String|QDDDD               |5144444444                                   |
|Base64    |`TLCD               |C2BUTENE                                     |
|String    |                    |                                             |
|Base64    |O`TLCD              |T2BUTENE                                     |
|String    |                    |                                             |
|Base64    |(                   |KCAL                                         |
|String    |                    |                                             |
+----------+--------------------+---------------------------------------------+

