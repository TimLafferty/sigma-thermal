olevba 0.60.2 on Python 3.11.4 - http://decalage.info/python/oletools
===============================================================================
FILE: sources/HC2-Calculators.xlsm
Type: OpenXML
WARNING  For now, VBA stomping cannot be detected for files in memory
-------------------------------------------------------------------------------
VBA MACRO ThisWorkbook.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/ThisWorkbook'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub workbook_open()
Application.Iteration = True
Worksheets("Heater Calcs").Range("ReturnTemp") = 50
End Sub



-------------------------------------------------------------------------------
VBA MACRO Module7.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module7'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet26.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet26'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub ComboBox1_Change()

End Sub

Private Sub designComboBox_Change()
    Worksheets("Burner and Controls Equip Area").Cells(70, "AG") = Me.designComboBox.Value
    
End Sub

Private Sub operatingComboBox_Change()
    Worksheets("Burner and Controls Equip Area").Cells(70, "AH") = Me.OperatingComboBox.Value
    
End Sub

Private Sub turndownComboBox_Change()
    Worksheets("Burner and Controls Equip Area").Cells(70, "AI") = Me.turndownComboBox.Value
    
End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet28.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet28'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet21.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet21'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module4.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module4'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet5.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet5'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Declarations.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Declarations'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Public HighFlow As Boolean

-------------------------------------------------------------------------------
VBA MACRO Sheet16.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet16'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module9.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module9'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Private Sub HeaterBalance()
'Initialize fluid temp

With Worksheets("Heater Calcs")
   .Range("ReturnTemp") = "=C7-Capacity*10^6/(C38*D6)"
    
    HeaterModel = Worksheets("Heater Calcs").Range("Model")
    
    Select Case HeaterModel
    
    'Series Heaters
    Case "HC2-0.5-SF", "HC2-1.0-SF", "HC2-1.5-SF", "HC2-2.0-SF", "HC2-2.5-SF", "HC2-3.0-SF", "HC2-4.0-SF", "DH 01/30", "DH 02/30", "DH 03/50", "DH 03/20", "DH 04/40", "DH 05/40", "DH 05/30", "DH 06/40", "DH 08/50", "DH 10/40"
               
        .Range("FluidTempIn3") = .Range("HeaterTempIn")
        .Range("FluidTempIn4") = .Range("HeaterTempIn")
        
        .Range("FluidFlow1") = .Range("HeaterFlow")
        .Range("FluidFlow2") = .Range("HeaterFlow")
        .Range("FluidFlow3") = .Range("HeaterFlow")
        .Range("FluidFlow4") = .Range("HeaterFlow")
        
        .Range("FluidTempIn2") = .Range("FluidTempOut3")
        .Range("FluidTempIn1") = .Range("FluidTempOut4")
        .Range("DPDiff") = "=DPo+DPi"
        
        For i = 1 To 1
            Balance
        Next i
        
        With Worksheets("Heater Calcs")
            .Range("FluidTempOut1") = .Range("FluidTempCalc1")
            .Range("FluidTempOut4") = .Range("FluidTempCalc4")
            
            .Range("TubeWall1") = .Range("TubeWallCalc1")
            .Range("TubeWall2") = .Range("TubeWallCalc2")
            .Range("TubeWall3") = .Range("TubeWallCalc3")
            .Range("TubeWall4") = .Range("TubeWallCalc4")
        End With
        Worksheets("Burner and Controls Equip Area").Range("EfficiencyAssumed") = .Range("EfficiencyCalc")
        Balance
        .Range("G28") = "=FluidTempOut1"
    
    'Parallel Heaters
    Case Else
            .Range("FluidTempIn1") = .Range("HeaterTempIn")
            .Range("FluidTempIn2") = .Range("HeaterTempIn")
            .Range("FluidTempIn3") = .Range("HeaterTempIn")
            .Range("FluidTempIn4") = .Range("HeaterTempIn")
            
            .Range("FluidFlow3") = 10000
            .Range("FluidFlow4") = "=FluidFlow3"
            .Range("FluidFlow1") = "=FluidFlow2"
            .Range("FluidFlow2") = "=HeaterFlow - FluidFlow3"
            .Range("DPDiff") = "=DPo-DPi"
            
            SolverOptions Precision:=0.001
            SolverOK SetCell:=Range("DPDiff"), MaxMinVal:=3, valueof:=0, ByChange:=Range("FluidFlow3")
            SolverSolve UserFinish:=True
            
            For i = 1 To 1
            Balance
        Next i
        
        With Worksheets("Heater Calcs")
            .Range("FluidTempOut1") = .Range("FluidTempCalc1")
            .Range("FluidTempOut4") = .Range("FluidTempCalc4")
            
            .Range("TubeWall1") = .Range("TubeWallCalc1")
            .Range("TubeWall2") = .Range("TubeWallCalc2")
            .Range("TubeWall3") = .Range("TubeWallCalc3")
            .Range("TubeWall4") = .Range("TubeWallCalc4")
        End With
        Worksheets("Burner and Controls Equip Area").Range("EfficiencyAssumed") = .Range("EfficiencyCalc")
        Balance
        .Range("G28") = "=(FluidTempCalc1*FluidFlow1+FluidTempCalc3*FluidFlow3)/HeaterFlow"
    
    End Select
    
End With
Sheets("Heater Calcs").Select
End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet2.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet31.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet31'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit

-------------------------------------------------------------------------------
VBA MACRO Sheet32.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet32'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet25.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet25'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module3.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module3'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module6.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module6'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet10.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet10'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet11.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet11'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet12.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet12'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module1.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub Balance()
   With Worksheets("Heater Calcs")
          
            'Section 1 Balance
        Sheets("Heater Calcs").Select
        
            SolverOptions Precision:=0.01
            SolverOK SetCell:=.Range("Diff1g"), MaxMinVal:=3, valueof:=0, ByChange:=.Range("GasTempOut1")
            SolverSolve UserFinish:=True
            
            'Section 2i Balance
            SolverOptions Precision:=0.01
            SolverOK SetCell:=.Range("Diff2g"), MaxMinVal:=3, valueof:=0, ByChange:=.Range("GasTempOut2")
            SolverSolve UserFinish:=True
            
            'Section 2o Balance
            SolverOptions Precision:=0.01
            SolverOK SetCell:=.Range("Diff3g"), MaxMinVal:=3, valueof:=0, ByChange:=.Range("GasTempOut3")
            SolverSolve UserFinish:=True
            
            'Section 3 Balance
            SolverOptions Precision:=0.01
            SolverOK SetCell:=.Range("Diff4g"), MaxMinVal:=3, valueof:=0, ByChange:=.Range("GasTempOut4")
            SolverSolve UserFinish:=True
            
        Sheets("New Primary Inputs").Select
    End With
End Sub

Sub HCCalc()
With Worksheets("Heater Calcs")
    .Range("ReturnTemp") = "=C7-Capacity*10^6/(C38*D6)"
    
    
    HeaterModel = Worksheets("Heater Calcs").Range("Model")
    
    Select Case HeaterModel
    
    'Series Heaters
    Case "HC2-0.5-SF", "HC2-1.0-SF", "HC2-1.5-SF", "HC2-2.0-SF", "HC2-2.5-SF", "HC2-3.0-SF", "HC2-4.0-SF", "DH 01/30", "DH 02/30", "DH 03/50", "DH 03/20", "DH 04/40", "DH 05/40", "DH 05/30", "DH 06/40", "DH 08/50", "DH 10/40"
        .Range("FlowArrangement") = "Series"
        .Range("FluidTempIn3") = .Range("HeaterTempIn")
        .Range("FluidTempIn4") = .Range("HeaterTempIn")
        
        .Range("FluidFlow1") = .Range("HeaterFlow")
        .Range("FluidFlow2") = .Range("HeaterFlow")
        .Range("FluidFlow3") = .Range("HeaterFlow")
        .Range("FluidFlow4") = .Range("HeaterFlow")
        
        .Range("FluidTempIn2") = .Range("FluidTempOut3")
        .Range("FluidTempIn1") = .Range("FluidTempOut4")
        .Range("DPDiff") = "=DPo+DPi"
        
        For i = 1 To 1
            Balance
        Next i
    
        With Worksheets("Heater Calcs")
            .Range("FluidTempOut1") = .Range("FluidTempCalc1")
            .Range("FluidTempOut4") = .Range("FluidTempCalc4")
            
            .Range("TubeWall1") = .Range("TubeWallCalc1")
            .Range("TubeWall2") = .Range("TubeWallCalc2")
            .Range("TubeWall3") = .Range("TubeWallCalc3")
            .Range("TubeWall4") = .Range("TubeWallCalc4")
        End With
        Worksheets("Burner and Controls Equip Area").Range("EfficiencyAssumed") = .Range("EfficiencyCalc")
        Balance
        .Range("G28") = "=FluidTempOut1"
    
    'Parallel Heaters
    Case Else
            .Range("FlowArrangement") = "Parallel"
            .Range("FluidTempIn1") = .Range("HeaterTempIn")
            .Range("FluidTempIn2") = .Range("HeaterTempIn")
            .Range("FluidTempIn3") = .Range("HeaterTempIn")
            .Range("FluidTempIn4") = .Range("HeaterTempIn")
            
            .Range("FluidFlow3") = 30000
            .Range("FluidFlow4") = "=FluidFlow3"
            .Range("FluidFlow1") = "=FluidFlow2"
            .Range("FluidFlow2") = "=HeaterFlow - FluidFlow3"
            .Range("DPDiff") = "=DPo-DPi"
            
            Sheets("Heater Calcs").Select
                SolverOptions Precision:=0.01
                SolverOK SetCell:=.Range("DPDiff"), MaxMinVal:=3, valueof:=0, ByChange:=.Range("FluidFlow3")
                SolverSolve UserFinish:=True
            
            For i = 1 To 1
            Balance
        Next i
        
        With Worksheets("Heater Calcs")
            .Range("FluidTempOut1") = .Range("FluidTempCalc1")
            .Range("FluidTempOut4") = .Range("FluidTempCalc4")
            
            .Range("TubeWall1") = .Range("TubeWallCalc1")
            .Range("TubeWall2") = .Range("TubeWallCalc2")
            .Range("TubeWall3") = .Range("TubeWallCalc3")
            .Range("TubeWall4") = .Range("TubeWallCalc4")
        End With

        Worksheets("Burner and Controls Equip Area").Range("EfficiencyAssumed") = .Range("EfficiencyCalc")
        Balance
        .Range("G28") = "=(FluidTempCalc1*FluidFlow1+FluidTempCalc3*FluidFlow3)/HeaterFlow"
    
    End Select
End With
End Sub
-------------------------------------------------------------------------------
VBA MACRO Module2.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module2'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("Calcs").Select
    Range("C52").GoalSeek Goal:=0, ChangingCell:=Range("C22")
    Sheets("New Primary Inputs").Select
End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet6.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet6'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet27.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet27'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet20.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet20'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Option Explicit



Private Sub cmdBalance_Click()
Dim i, neg As Integer
Dim step As Single

Dim InnerCoilDrop As Single, OuterCoilDrop As Single, InnerCoilFlow As Single, OuterCoilFlow As Single, CoilType As String, Diff As Double

For i = 14 To 37
    CoilType = Worksheets("Heater Table").Cells(i, 8)
    
    If CoilType = "Parallel" Then
        Worksheets("Heater Table").Cells(i, 18) = 10000
    End If
        
    InnerCoilDrop = Worksheets("Heater Table").Cells(i, 24)
    OuterCoilDrop = Worksheets("Heater Table").Cells(i, 37)
    InnerCoilFlow = Worksheets("Heater Table").Cells(i, 18)
    OuterCoilFlow = Worksheets("Heater Table").Cells(i, 30)

step = 10000
neg = 1

    Diff = (OuterCoilDrop - InnerCoilDrop)
    Do While Abs(Diff) > 0.1 And CoilType = "Parallel"
        InnerCoilDrop = Worksheets("Heater Table").Cells(i, 24)
        OuterCoilDrop = Worksheets("Heater Table").Cells(i, 37)
        InnerCoilFlow = Worksheets("Heater Table").Cells(i, 18)
        OuterCoilFlow = Worksheets("Heater Table").Cells(i, 30)
        Diff = (OuterCoilDrop - InnerCoilDrop)
        
        InnerCoilFlow = InnerCoilFlow + step
        
        If Diff * neg < 0 Then
              step = step / -10
              neg = neg * -1
        End If
    
    Worksheets("Heater Table").Cells(i, 18) = InnerCoilFlow
       
    Loop

Next i

For i = 41 To 63
    CoilType = Worksheets("Heater Table").Cells(i, 8)
    
    If CoilType = "Parallel" Then
        Worksheets("Heater Table").Cells(i, 18) = 10000
    End If
        
    InnerCoilDrop = Worksheets("Heater Table").Cells(i, 24)
    OuterCoilDrop = Worksheets("Heater Table").Cells(i, 37)
    InnerCoilFlow = Worksheets("Heater Table").Cells(i, 18)
    OuterCoilFlow = Worksheets("Heater Table").Cells(i, 30)

step = 10000
neg = 1

    Diff = (OuterCoilDrop - InnerCoilDrop)
    Do While Abs(Diff) > 0.1 And CoilType = "Parallel"
        InnerCoilDrop = Worksheets("Heater Table").Cells(i, 24)
        OuterCoilDrop = Worksheets("Heater Table").Cells(i, 37)
        InnerCoilFlow = Worksheets("Heater Table").Cells(i, 18)
        OuterCoilFlow = Worksheets("Heater Table").Cells(i, 30)
        Diff = (OuterCoilDrop - InnerCoilDrop)
        
        InnerCoilFlow = InnerCoilFlow + step
        
        If Diff * neg < 0 Then
              step = step / -10
              neg = neg * -1
        End If
    
    Worksheets("Heater Table").Cells(i, 18) = InnerCoilFlow
       
    Loop

Next i
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet24.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet24'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet13.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet13'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module10.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module10'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet9.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet9'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 

Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_BeforeDelete()

End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)

End Sub

Private Sub Worksheet_Deactivate()

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
-------------------------------------------------------------------------------
VBA MACRO Sheet14.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet14'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module8.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module8'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module5.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module5'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet1.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet1'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet3.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet3'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Module11.bas 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Module11'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
Sub Create_Word()

    Dim applWord As Word.Application
    Dim docWord As Word.Document
    Dim objSelection As Word.Selection
    Dim ws, ws1, ws2, ws3 As Worksheet
    Set ws = Sheets("Customer Datasheet")
    Set ws1 = Sheets("Qtion")
    Set ws2 = Sheets("New Primary Inputs")
    Set ws3 = Sheets("Heater Calcs")
    
    ' Signature Vars
     Dim Signature As Word.InlineShape
     Dim SaveNameQtion, SignQtion As String
     
    'Other vars
     Dim HeaterRng As Word.Range
     Dim BurnerRng As Word.Range
     Dim ControlRng As Word.Range
     Dim LrowQ As Long
     Dim TemplateQ As String
    
    ' Scope Vars
        ' Heater
        Dim HeaterConfigQ, HeaterSkidQ, HeaterStackQ, HeaterTrainQ, HeaterPanelQ, HeaterQtion, HeaterCodeQ, HeaterFlowQ As String
        Dim HeaterInCoilQ, HeaterOutCoilQ As Integer
            Dim Heatershape As Word.InlineShape
            Dim Heatershapetype As Word.Shape
            HeaterInCoilQ = HeaterOutCoilQ = 0
        'Burner and train
        Dim BurnerFuelQ, BurnerMakeQ, BurnerModelQ, BurnerControlQ, BurnerQtion As String
            Dim Burnershape As Word.InlineShape
            Dim Burnershapetype As Word.Shape
        'Controls
        Dim BMSQ, SpecialBMSQ, CombControlQ, SpecialCombQ, OPBMSQ, OPCombQ, PanelAreaClassQ As String
            Dim Panelshape As Word.InlineShape
            Dim Panelshapetype As Word.Shape
            Dim PLCHMIshape As Word.InlineShape
        'Instrumentation
        Dim OverTQ, DifPQ As Integer
            OverTQ = DifPQ = 0
        'Pumps
        Dim PumpSkidQ, PumpValveQ, PumpValveTextQ, QtyPumpsQ, PumpOperQ As String
            Dim Pumpshape As Word.InlineShape
            Dim Pumpshapetype As Word.Shape
        'Expansion tank
        Dim DoubleLegQ, ExpTankGalQ, ExpTankCodeQ, ExpTankLevelQ, OPExpTankQ As String
        'Blanket
        Dim BlanketQ, BlankettypeQ As String
        'Tank Tower
        Dim TankTowerQ As String
            Dim TankTowershape As Word.InlineShape
            Dim TankTowershapetype As Word.Shape
        'Burner House
        Dim BurnerHouseQ As String
            Dim BurnerHouseshape As Word.InlineShape
            Dim BurnerHouseshapetype As Word.Shape
        ' Motor Starter
        Dim MotorStarterQ, VoltageQ, VoltageshortQ, NEMARatingQ, OPMotorStarterQ As String
            Dim MotorStartershape As Word.InlineShape
            Dim MotorStartershapetype As Word.Shape
        'Drain tank & Pump
        Dim DrainTankQ, DrainPumpQ, DrainTankCodeQ, OPDrainTankQ As String
        ' Exchangers
        Dim TCUQ, STHEXQ As String
           Dim TCUshape, STHEXQshape As Word.InlineShape
           Dim TCUshapetype, STHEXQshapetype As Word.Shape
        'Comb Air Pre-Heater
        Dim CombAirQ As String
           Dim CombAirshape As Word.InlineShape
           Dim CombAirshapetype As Word.Shape
        'Heat Tracing
        Dim FTHeatTQ, PSVHeatTQ As String
        'Stress Analysis
        Dim StressAnalysisQ As String
           Dim StressAnalysisshape As Word.InlineShape
           Dim StressAnalysisshapetype As Word.Shape
        'Remote I/O
        Dim RemoteIOQ As String
           Dim RemoteIOshape As Word.InlineShape
           Dim RemoteIOshapetype As Word.Shape
        'Oxygen Analyzer
        Dim O2AnalyzerQ As String
        'Ladder Platform
        Dim LaddersPlatQ As String
        'Low temp Fuel gas pipe
        Dim LowTempFGQ As String
        'Registrations, stamps, Canadian Codes
        Dim CanadianCodesQ, PEngStampQ As String
        'VFD
        Dim BlowerVFDQ, PumpVFDQ As String
        'Delivery
        Dim PIDMinQ, PIDMaxQ, GAMinQ, GAMaxQ, ReadyShMinQ, ReadyShMaxQ, PIDAppQ, GAAppQ As String
        

    ' Get Scope vars from Qtion Inputs
        ws1.Activate
        TemplateQ = Range("K3")
        SpecialBMSQ = Range("K8")
        SpecialCombQ = Range("K9")
        RemoteIOQ = Range("K15")
        DoubleLegQ = Range("K17")
        ExpTankLevelQ = Range("K18")
        TCUQ = Range("K19")
        STHEXQ = Range("K20")
        FTHeatTQ = Range("K21")
        PSVHeatTQ = Range("K22")
        StressAnalysisQ = Range("K23")
        O2AnalyzerQ = Range("K24")
        LaddersPlatQ = Range("K25")
        LowTempFGQ = Range("K26")
        CanadianCodesQ = Range("K27")
        PEngStampQ = Range("K28")
        OPBMSQ = Range("K12")
        OPCombQ = Range("K13")
        OPExpTankQ = Range("K16")
        OPMotorStarterQ = Range("K29")
        OPDrainTankQ = Range("K30")
        BlowerVFDQ = Range("K31")
        PumpVFDQ = Range("K32")
        PIDMinQ = Range("N17")
        PIDMaxQ = Range("O17")
        GAMinQ = Range("N18")
        GAMaxQ = Range("O18")
        ReadyShMinQ = Range("N19")
        ReadyShMaxQ = Range("O19")
        PIDAppQ = Range("P17")
        GAAppQ = Range("P18")
         
    ' Get Scope vars from Primary Inputs
        ws2.Activate
        HeaterConfigQ = Range("B18")
        HeaterSkidQ = Range("F65")
        HeaterStackQ = Range("F55")
        HeaterTrainQ = Range("B63")
        HeaterPanelQ = Range("B72")
        HeaterCodeQ = Range("F71")
        
        BurnerFuelQ = Range("F26")
        BurnerMakeQ = Range("F41")
        BurnerControlQ = Range("F44")
        BurnerModelQ = Range("F42")
        
        BMSQ = Range("B65")
        CombControlQ = Range("B66")
        
        PumpSkidQ = Range("B25")
        PumpValveQ = Range("B39")
        QtyPumpsQ = Range("B28")
        PumpOperQ = Range("C28")
        
        ExpTankGalQ = Range("B51")
        ExpTankCodeQ = Range("F73")
        
        BlanketQ = Range("F58")
        BlankettypeQ = Range("F59")
        
        TankTowerQ = Range("F61")
        
        BurnerHouseQ = Range("F64")
        
        MotorStarterQ = Range("B73")
        VoltageQ = Range("F12")
        VoltageshortQ = Left(Range("F12").Value, 4)
        NEMARatingQ = Range("B71")
        AmbientTempMinQ = Range("F7")
        AmbientTempMaxQ = Range("F6")
        
        DrainTankQ = Range("B53")
        DrainPumpQ = Range("B54")
        DrainTankCodeQ = Range("F74")
                
        SideStreamQ = Range("F68")

        CombAirQ = Range("F27")
        
        PanelAreaClassQ = Range("F9")

    ' Get Scope vars from Heater Calcs
        ws3.Activate
        HeaterFlowQ = Range("F4")
        HeaterInCoilQ = Range("C75")
        HeaterOutCoilQ = Range("E75")
    
    ' Copy from Customer Data to Qtion sheet
        ws.Activate
        LrowQ = ws.Cells(Rows.Count, 1).End(xlUp).Offset(1).Row - 1
        ws.Range("A1:G" & LrowQ).Copy
        ws1.Activate
        ws1.Range("A1").Select
        ws1.Paste
        ws1.Range("A1").Select
        ws1.PasteSpecial xlPasteColumnWidths
        Application.CutCopyMode = False
          
    ' Clear the pumps if not included and if they have not been cleared yet
    If PumpSkidQ = "No" Then
      ws1.Activate
        If Range("E37") = "Pump Model" Then
            Rows(37).EntireRow.Delete
            Rows(37).EntireRow.Delete
            Rows(37).EntireRow.Delete
            Rows(37).EntireRow.Delete
            Rows(37).EntireRow.Delete
            Range("A36") = "Piping, & Expansion Tank Data"
        End If
    Else
        ws1.Activate
    End If
    
    ' Open Word
    Set applWord = New Word.Application
    With applWord
        .Visible = True
        .Activate
        
    End With
    
    'Choose Americas or M.East template
    If TemplateQ = "Americas" Then
    Set docWord = applWord.Documents.Add(Template:="C:\Qtion\QuoteUS.docx", NewTemplate:=False, DocumentType:=0)
    Else
    Set docWord = applWord.Documents.Add(Template:="C:\Qtion\QuoteME.docx", NewTemplate:=False, DocumentType:=0)
    End If
    
    With docWord
         
      ' Styles being used in the quote
       If TemplateQ = "Americas" Then
        ' First Style
        .Styles.Add Name:="StyleQ1", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ1")
            .Font.Name = "Cambria"
            .Font.Size = 12
            .Font.Color = 0
            .Font.Bold = False
            .ParagraphFormat.Alignment = 0
            ' 0 left 1 center
        End With
        
        'Second style
        .Styles.Add Name:="StyleQ2", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ2")
            .Font.Name = "Cambria"
            .Font.Size = 12
            .AutomaticallyUpdate = False
            .BaseStyle = "List Bullet"
            .NextParagraphStyle = "StyleQ1"
            .ParagraphFormat.LeftIndent = 72
        End With
        
        'Third style
        .Styles.Add Name:="StyleQ3", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ3")
            .Font.Name = "Cambria"
            .Font.Size = 12
            .Font.Color = 0
            .Font.Bold = True
            .ParagraphFormat.LeftIndent = 42
        End With
        
        'Fourth Style
        .Styles.Add Name:="StyleQ4", Type:=wdStyleTypeCharacter
        With .Styles("StyleQ4")
            .Font.Name = "Cambria"
            .Font.Size = 12
            .Font.Color = 0
            .Font.Bold = True
        End With
        
        'Fifth Style
        .Styles.Add Name:="StyleQ5", Type:=wdStyleTypeCharacter
        With .Styles("StyleQ5")
            .Font.Name = "Cambria"
            .Font.Size = 12
            .Font.Color = 0
            .Font.Bold = False
        End With
        
        'Sixth style
        .Styles.Add Name:="StyleQ6", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ6")
            .Font.Name = "Cambria"
            .Font.Size = 12
            .Font.Color = 0
            .Font.Bold = False
            .ParagraphFormat.LeftIndent = 30
        End With

        'Seventh style
        .Styles.Add Name:="StyleQ7", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ7")
            .Font.Name = "Cambria"
            .Font.Size = 12
            .Font.Color = 0
            .Font.Bold = True
            .ParagraphFormat.LeftIndent = 30
        End With
        
        
        Else
         ' MIDDLE EAST STYLES
         ' First Style
        .Styles.Add Name:="StyleQ1", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ1")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = 0
            .Font.Bold = False
            .ParagraphFormat.Alignment = 0
            ' 0 left 1 center
        End With
        
        'Second style
        .Styles.Add Name:="StyleQ2", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ2")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .AutomaticallyUpdate = False
            .BaseStyle = "List Bullet"
            .NextParagraphStyle = "StyleQ1"
            .ParagraphFormat.LeftIndent = 72
            .ParagraphFormat.Alignment = 0
            .Font.Bold = False
        End With
        
        'Third style
        .Styles.Add Name:="StyleQ3", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ3")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = 0
            .Font.Bold = True
            .ParagraphFormat.LeftIndent = 42
            .ParagraphFormat.Alignment = 0
        End With
        
        'Fourth Style
        .Styles.Add Name:="StyleQ4", Type:=wdStyleTypeCharacter
        With .Styles("StyleQ4")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = 0
            .Font.Bold = True

        End With
        
        'Fifth Style
        .Styles.Add Name:="StyleQ5", Type:=wdStyleTypeCharacter
        With .Styles("StyleQ5")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = 0
            .Font.Bold = False
        End With
        
        'Sixth style
        .Styles.Add Name:="StyleQ6", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ6")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = 0
            .Font.Bold = False
            .ParagraphFormat.LeftIndent = 30
            .ParagraphFormat.Alignment = 0
        End With

        'Seventh style
        .Styles.Add Name:="StyleQ7", Type:=wdStyleTypeParagraph
        With .Styles("StyleQ7")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = 0
            .Font.Bold = False
            .ParagraphFormat.LeftIndent = 30
            .ParagraphFormat.Alignment = 0
        End With
        
        End If
        
        
        ' Insert Prospect Person Name, Location & OPP#
        If TemplateQ = "Americas" Then
                .Bookmarks("Person1Loc").Range.Text = Range("K5")
                .Bookmarks("CompanyLoc").Range.Text = Range("B2")
                .Bookmarks("PersonLoc").Range.Text = Range("K5")
                .Bookmarks("LocationLoc").Range.Text = Range("F3")
            If Range("B4") = 0 Then
                .Bookmarks("QuoteNum1Loc").Range.Text = Range("K4")
                Else
                .Bookmarks("QuoteNum1Loc").Range.Text = (Range("K4") & " R" & Range("B4"))
            End If
        
        Else
                .Bookmarks("Person1Loc").Range.Text = Range("K5")
                .Bookmarks("CompanyLoc").Range.Text = Range("B2")
                .Bookmarks("Company1Loc").Range.Text = Range("B2")
                .Bookmarks("Company2Loc").Range.Text = Range("B2")
                .Bookmarks("PersonLoc").Range.Text = Range("K5")
                .Bookmarks("LocationLoc").Range.Text = Range("F3")
                .Bookmarks("Location1Loc").Range.Text = Range("F3")
                .Bookmarks("ProjectLoc").Range.Text = Range("F2")
            If Range("B4") = 0 Then
                .Bookmarks("QuoteNum1Loc").Range.Text = Range("K4")
                Else
                .Bookmarks("QuoteNum1Loc").Range.Text = (Range("K4") & " R" & Range("B4"))
            End If
            If Range("B4") = 0 Then
                .Bookmarks("QuoteNum2Loc").Range.Text = Range("K4")
                Else
                .Bookmarks("QuoteNum2Loc").Range.Text = (Range("K4") & "-" & Range("B4"))
            End If
        
        End If
        
                
      ' Assign the right salesperson to the signature variables and fill text in
        With .ActiveWindow.Selection
        Select Case ws.Range("F4")
            Case Is = "Joaquin Grimaux"
                SignQtion = "C:\Qtion\Signatures\SignJG.png"
                .GoTo What:=-1, Name:="SignatureBLoc"
                .TypeText (Range("F4") & vbNewLine & "Sigma Thermal")
                .GoTo What:=-1, Name:="SignatureCLoc"
                .TypeText ("Phone:      (678) 905-3884" & vbNewLine & "Cell:           +54 (911) 6831-3298" & vbNewLine & "Fax:           (678) 254-1762")
                .GoTo What:=-1, Name:="SignatureDLoc"
                .TypeText ("jgrimaux@sigmathermal.com" & vbNewLine & "www.sigmathermal.com")
            Case Is = "William Gardiner"
                SignQtion = "C:\Qtion\Signatures\SignWG.png"
                .GoTo What:=-1, Name:="SignatureBLoc"
                .TypeText (Range("F4") & vbNewLine & "Sigma Thermal")
                .GoTo What:=-1, Name:="SignatureCLoc"
                .TypeText ("Phone:      (678) 324-5756" & vbNewLine & "Cell:           (678) 641-2900" & vbNewLine _
                & "Fax:           (678) 254-1762")
                .GoTo What:=-1, Name:="SignatureDLoc"
                .TypeText ("wgardiner@sigmathermal.com" & vbNewLine & "www.sigmathermal.com")
            Case Is = "Nick Krauss"
                SignQtion = "C:\Qtion\Signatures\SignNK.png"
                .GoTo What:=-1, Name:="SignatureBLoc"
                .TypeText (Range("F4") & vbNewLine & "Sigma Thermal")
                .GoTo What:=-1, Name:="SignatureCLoc"
                .TypeText ("Phone:      (678) 324-5767" & vbNewLine & "Cell:           (770) 317-0373" & vbNewLine _
                & "Fax:           (678) 254-1762")
                .GoTo What:=-1, Name:="SignatureDLoc"
                .TypeText ("nkrauss@sigmathermal.com" & vbNewLine & "www.sigmathermal.com")
            Case Is = "Vaheed Jaberi"
                SignQtion = "C:\Qtion\Signatures\SignVJ.png"
                .GoTo What:=-1, Name:="SignatureBLoc"
                .TypeText (Range("F4") & vbNewLine & "Sigma Thermal")
                .GoTo What:=-1, Name:="SignatureCLoc"
                .TypeText ("Phone:      (678) 324-5779" & vbNewLine & "Cell:           (770) 639-0229" & vbNewLine _
                & "Fax:           (678) 254-1762")
                .GoTo What:=-1, Name:="SignatureDLoc"
                .TypeText ("vjaberi@sigmathermal.com" & vbNewLine & "www.sigmathermal.com")
            Case Is = "Charlie Wadlington"
                SignQtion = "C:\Qtion\Signatures\SignCW.png"
                .GoTo What:=-1, Name:="SignatureBLoc"
                .TypeText (Range("F4") & vbNewLine & "Sigma Thermal")
                .GoTo What:=-1, Name:="SignatureCLoc"
                .TypeText ("Phone:      (678) 324-5726" & vbNewLine & "Cell:           (404) 428-1042" & vbNewLine _
                & "Fax:           (678) 254-1762")
                .GoTo What:=-1, Name:="SignatureDLoc"
                .TypeText ("cwadlington@sigmathermal.com" & vbNewLine & "www.sigmathermal.com")
            Case Is = "Caroline Blake"
                SignQtion = "C:\Qtion\Signatures\SignCB.png"
                .GoTo What:=-1, Name:="SignatureBLoc"
                .TypeText (Range("F4") & vbNewLine & "Sigma Thermal")
                .GoTo What:=-1, Name:="SignatureCLoc"
                .TypeText ("Phone:      (678) 324-5778" & vbNewLine & "Cell:           (205) 919-7115" & vbNewLine _
                & "Fax:           (678) 254-1762")
                .GoTo What:=-1, Name:="SignatureDLoc"
                .TypeText ("cblake@sigmathermal.com" & vbNewLine & "www.sigmathermal.com")
            Case Else
                MsgBox "Select a Salesperson and run it again. Be yourself."
        End Select
        End With
       ' Insert signature .png from file to the quote bookmark
        Set Signature = .Bookmarks("SignatureLoc").Range.InlineShapes.AddPicture(Filename:=SignQtion, LinkToFile:=False, SaveWithDocument:=True)
                            
       ' Insert OPP# into header
       If Range("B4") = 0 Then
            .Bookmarks("QuoteNum2Loc").Range.Text = Range("K4")
            Else
            .Bookmarks("QuoteNum2Loc").Range.Text = (Range("K4") & "-" & Range("B4"))
        End If
        
        ' Insert Data sheet as an image
        ws1.Activate
        LrowQ = ws1.Cells(Rows.Count, 1).End(xlUp).Offset(1).Row - 1
        Range("A1:G" & LrowQ).CopyPicture Appearance:=1, Format:=-4147
        DoEvents
        .Bookmarks("DSLoc").Range.Paste
        Application.CutCopyMode = False
                        
       'System Description for Middle East template only
        If TemplateQ = "Americas" Then
        Else
            .Bookmarks("BurnerFuelDescLoc").Range.Text = BurnerFuelQ
            .Bookmarks("BurnerFuelDesc1Loc").Range.Text = BurnerFuelQ
          With .ActiveWindow.Selection
            .GoTo What:=-1, Name:="StackDescLoc"
            If HeaterStackQ = "Free Standing" Or HeaterStackQ = "Heater Mounted" Then
                If HeaterStackQ = "Free Standing" Then
                .TypeText ("The heater will have its own self-supporting exhaust stack. ")
                Else
                .TypeText ("The heater will have its own Heater Mounted exhaust stack. ")
                End If
            Else
            End If
                .GoTo What:=-1, Name:="HeaterMountLoc"
                If HeaterSkidQ = Yes Then
                .TypeText ("The Heater will be skid mounted.")
                Else
                .TypeText ("The Heater will be saddle mounted.")
                End If
          End With
        End If
                        
       ' Define scope
       With .ActiveWindow.Selection
        .GoTo What:=-1, Name:="FirstScopeLoc"
        .Style = docWord.Styles("StyleQ6")
        .TypeText (vbNewLine & "Thermal fluid system equipment engineering, design, & project management" & vbNewLine)
        .TypeText ("Thermal fluid heater" & vbNewLine)
        .TypeText ("Burner and fuel train" & vbNewLine)
        If CombControlQ = "Standard" And BMSQ = "Standard" Then
         .TypeText ("Heater control panel and burner management system" & vbNewLine)
        Else
          If CombControlQ = "Standard" And BMSQ <> "Standard" Then
         .TypeText ("Heater control panel and PLC based burner management system" & vbNewLine)
          Else
           If CombControlQ <> "Standard" And BMSQ = "Standard" Then
           .TypeText ("Heater PLC based combustion control and burner management system panel" & vbNewLine)
           Else
           .TypeText ("Heater PLC based combustion control panel" & vbNewLine & "PLC based burner management system" & vbNewLine)
           End If
         End If
        End If
        
        If PumpSkidQ = "Yes" Then
         .TypeText (QtyPumpsQ & " " & PumpOperQ & " Primary Loop Pump Skid with motor/s and valve/s" & vbNewLine)
        Else
        End If
        
        If ExpTankGalQ = "By Others" Or OPExpTankQ = "Yes" Then
          Else
        .TypeText (ExpTankGalQ & " Gal Horizontal cylindrical Expansion Tank" & vbNewLine)
         End If
        
        If BlanketQ = "Yes" Then
        .TypeText (BlankettypeQ & " expansion tank blanket system" & vbNewLine)
        Else
        End If
        
        If MotorStarterQ = "By Sigma Thermal" Then
          .TypeText ("Motor Starters" & vbNewLine)
        Else
        End If
        
        If TankTowerQ = "Yes" Then
         .TypeText ("Expansion tank tower & interconnecting piping" & vbNewLine)
        Else
        End If

        If HeaterStackQ = "Free Standing" Or HeaterStackQ = "Heater Mounted" Then
          .TypeText ("Exhaust Stack" & vbNewLine)
        Else
        End If
        
        If DrainTankQ = "By Others" Or OPDrainTankQ = "Yes" Then
        Else
        .TypeText (DrainTankQ & " Gal System Drain Tank" & vbNewLine)
        End If
       
        If DrainPumpQ = "Sigma Thermal Standard" Or DrainPumpQ = "Custom" Then
        .TypeText ("Fill / Drain pump" & vbNewLine)
        Else
        End If
       
        If SideStreamQ = "Yes" Then
         .TypeText ("Side stream filter package" & vbNewLine)
        Else
        End If
       
        If TCUQ = "Yes" Then
        .TypeText ("Temperture Control Unit Skid" & vbNewLine)
        Else
        End If
       
        If STHEXQ = "Yes" Then
        .TypeText ("Shell & Tube Heat Exchanger Skid" & vbNewLine)
        Else
        End If
        
        If CombAirQ = "Yes" Then
        .TypeText ("Combustion air pre-heat package" & vbNewLine)
        Else
        End If
       
       If FTHeatTQ = "Yes" Then
         .TypeText ("Fuel train Heat Tracing" & vbNewLine)
        Else
        End If
       
       If PSVHeatTQ = "Yes" Then
         .TypeText ("PSV Heat Tracing" & vbNewLine)
        Else
       End If
       
       If StressAnalysisQ = "Yes" Then
        .TypeText ("Piping Design & Stress Analysis" & vbNewLine)
        Else
       End If
              
       If BurnerHouseQ = "Yes" Then
        .TypeText ("Burner House Enclosure" & vbNewLine)
       Else
       End If
       
       If RemoteIOQ = "Yes" Then
         .TypeText ("Remote I/O" & vbNewLine)
        Else
       End If
       
       If O2AnalyzerQ = "Yes" Then
        .TypeText ("Flue Gas Oxygen Analyzer" & vbNewLine)
       Else
       End If
       
       If LaddersPlatQ = "Yes" Then
        .TypeText ("Access Ladders & Platforms" & vbNewLine)
       Else
       End If
       
       If LowTempFGQ = "Yes" Then
        .TypeText ("Low temperature Fuel Gas Train" & vbNewLine)
       Else
       End If

       If CanadianCodesQ = "Yes" Then
        .TypeText ("Compliance with Canadian Codes" & vbNewLine)
       Else
       End If
       
       If PEngStampQ = "Yes" Then
        .TypeText ("Registered P. Eng. Stamp" & vbNewLine)
       Else
       End If
         
       If BlowerVFDQ = "Yes" Then
        .TypeText ("Blower VFD" & vbNewLine)
       Else
       End If
      
       If PumpVFDQ = "Yes" Then
        .TypeText ("Pump Variable Frequency Drive/s" & vbNewLine)
       Else
       End If
    
     'Optional Scope
      .GoTo What:=-1, Name:="FirstOptionScopeLoc"
      .Style = docWord.Styles("StyleQ6")
      .TypeText (vbNewLine)
      If OPBMSQ = "N/A" Then
       .TypeText ("")
       Else
       If OPBMSQ = "Standard" Then
        .TypeText ("Standard Burner Management System" & vbNewLine)
        Else
        .TypeText ("PLC Based Burner Management System" & vbNewLine)
       End If
      End If
      If OPCombQ = "N/A" Then
       .TypeText ("")
       Else
       If OPCombQ = "Standard" Then
        .TypeText ("Standard Combustion Control System" & vbNewLine)
        Else
        .TypeText ("PLC Based Combustion Control System" & vbNewLine)
       End If
      End If
      
       If PumpSkidQ = "Optional Scope" Then
        .TypeText (QtyPumpsQ & " " & PumpOperQ & " Primary Loop Pump Skid with motor/s and valve/s" & vbNewLine)
        Else
       End If
         
       If OPExpTankQ = "Yes" Then
       .TypeText (ExpTankGalQ & " Gal Horizontal Cylindrical Expansion Tank" & vbNewLine)
         Else
       End If
        
       If BlanketQ = "Optional Scope" Then
       .TypeText (BlankettypeQ & " for the Expansion Tank Blanketing System" & vbNewLine)
       Else
       End If
        
       If OPMotorStarterQ = "Yes" Then
         .TypeText ("Motor Starters" & vbNewLine)
       Else
       End If
        
       If TankTowerQ = "Optional Scope" Then
        .TypeText ("Expansion tank tower & interconnecting piping" & vbNewLine)
       Else
       End If

       If HeaterStackQ = "Optional Scope" Then
         .TypeText ("Exhaust Stack" & vbNewLine)
       Else
       End If
        
       If OPDrainTankQ = "Yes" Then
       .TypeText (DrainTankQ & " Gal System Drain Tank" & vbNewLine)
       Else
       End If
       
       If DrainPumpQ = "Optional Scope" Then
       .TypeText ("Fill / Drain pump" & vbNewLine)
       Else
       End If
       
       If SideStreamQ = "Optional Scope" Then
        .TypeText ("Side stream filter package" & vbNewLine)
       Else
       End If
       
       If TCUQ = "Optional" Then
       .TypeText ("Temperture Control Unit Skid" & vbNewLine)
       Else
       End If
       
       If STHEXQ = "Optional" Then
       .TypeText ("Shell & Tube Heat Exchanger Skid" & vbNewLine)
       Else
       End If
        
       If CombAirQ = "Optional Scope" Then
       .TypeText ("Combustion air pre-heat package" & vbNewLine)
       Else
       End If
       
       If FTHeatTQ = "Optional" Then
        .TypeText ("Fuel train Heat Tracing" & vbNewLine)
       Else
       End If
       
       If PSVHeatTQ = "Optional" Then
        .TypeText ("PSV Heat Tracing" & vbNewLine)
       Else
       End If
       
       If StressAnalysisQ = "Optional" Then
        .TypeText ("Piping Design & Stress Analysis" & vbNewLine)
       Else
       End If
              
       If BurnerHouseQ = "Optional Scope" Then
        .TypeText ("Burner House Enclosure" & vbNewLine)
       Else
       End If
       
       If RemoteIOQ = "Optional" Then
        .TypeText ("Remote I/O" & vbNewLine)
       Else
       End If
       
       If O2AnalyzerQ = "Optional" Then
        .TypeText ("Flue Gas Oxygen Analyzer" & vbNewLine)
       Else
       End If
       
       If LaddersPlatQ = "Optional" Then
        .TypeText ("Access Ladders & Platforms" & vbNewLine)
       Else
       End If
       
       If LowTempFGQ = "Optional" Then
        .TypeText ("Low temperature Fuel Gas Train" & vbNewLine)
       Else
       End If

       If CanadianCodesQ = "Optional" Then
        .TypeText ("Compliance with Canadian Codes" & vbNewLine)
       Else
       End If
       
       If PEngStampQ = "Optional" Then
        .TypeText ("Registered P. Eng. Stamp" & vbNewLine)
       Else
       End If
         
       If BlowerVFDQ = "Optional" Then
        .TypeText ("Blower VFD" & vbNewLine)
       Else
       End If
      
       If PumpVFDQ = "Optional" Then
        .TypeText ("Pump Variable Frequency Drive/s" & vbNewLine)
       Else
       End If
       
       .TypeText ("Start-up support")
       
      End With
     'Customer's scope
        Call CustomerscopeQ(docWord, MotorStarterQ, BlowerVFDQ, PumpVFDQ, HeaterStackQ)
        
    
      ' Select Heater image
       Select Case HeaterConfigQ
            'HORIZONTAL
            Case Is = "Horizontal"
                If HeaterSkidQ = "Yes" Then
                    Select Case HeaterStackQ
                        Case Is = "Heater Mounted"
                            If HeaterTrainQ = "Heater Mounted" And HeaterPanelQ = "Mounted Local to Heater" Then
                                HeaterQtion = "C:\Qtion\Heater\HmoStskid.png"
                            Else
                                HeaterQtion = "C:\Qtion\Heater\HnoStnoskid.png"
                            End If
                        Case Is = "Free Standing"
                            If HeaterTrainQ = "Heater Mounted" And HeaterPanelQ = "Mounted Local to Heater" Then
                                HeaterQtion = "C:\Qtion\Heater\Hfreestskid.png"
                            Else
                                HeaterQtion = "C:\Qtion\Heater\HnoStnoskid.png"
                            End If
                        
                        Case Else
                            If HeaterTrainQ = "Heater Mounted" And HeaterPanelQ = "Mounted Local to Heater" Then
                                HeaterQtion = "C:\Qtion\Heater\HnoStskid.png"
                            Else
                                HeaterQtion = "C:\Qtion\Heater\HnoStnoskid.png"
                            End If
                    End Select
                Else
                    Select Case HeaterStackQ
                        Case Is = "Heater Mounted"
                        If HeaterTrainQ = "Heater Mounted" And HeaterPanelQ = "Mounted Local to Heater" Then
                                HeaterQtion = "C:\Qtion\Heater\HmoStnoskid.png"
                            Else
                                HeaterQtion = "C:\Qtion\Heater\HnoStnoskid.png"
                            End If
                        Case Is = "Free Standing"
                        HeaterQtion = "C:\Qtion\Heater\HnoStnoskid.png"
                        Case Else
                        HeaterQtion = "C:\Qtion\Heater\HnoStnoskid.png"
                    End Select
                End If
            
            ' VERTICAL Downfired
            Case Is = "Vertical Downfired"
                If HeaterSkidQ = "Yes" Then
                    Select Case HeaterStackQ
                        Case Is = "Heater Mounted"
                        MsgBox ("Vertical Downfired with a Heater Mounted Stack ?")
                        HeaterQtion = "C:\Qtion\Heater\VdonoStnoskid.png"
                        Case Is = "Free Standing"
                        HeaterQtion = "C:\Qtion\Heater\VdonoStnoskid.png"
                        Case Else
                        HeaterQtion = "C:\Qtion\Heater\VdonoStnoskid.png"
                    End Select
                Else
                    Select Case HeaterStackQ
                        Case Is = "Heater Mounted"
                        MsgBox ("Vertical Downfired with a Heater Mounted Stack ?")
                        HeaterQtion = "C:\Qtion\Heater\VdonoStnoskid.png"
                        Case Is = "Free Standing"
                        HeaterQtion = "C:\Qtion\Heater\VdonoStnoskid.png"
                        Case Else
                        HeaterQtion = "C:\Qtion\Heater\VdonoStnoskid.png"
                    End Select
                End If
            
            ' VERTICAL Upfired
            Case Else
                If HeaterSkidQ = "Yes" Then
                    Select Case HeaterStackQ
                        Case Is = "Heater Mounted"
                        HeaterQtion = "C:\Qtion\Heater\VupnoStnoskid.png"
                        Case Is = "Free Standing"
                        HeaterQtion = "C:\Qtion\Heater\VupnoStnoskid.png"
                        Case Else
                        HeaterQtion = "C:\Qtion\Heater\VupnoStnoskid.png"
                    End Select
                Else
                    Select Case HeaterStackQ
                        Case Is = "Heater Mounted"
                        HeaterQtion = "C:\Qtion\Heater\VupnoStnoskid.png"
                        Case Is = "Free Standing"
                        HeaterQtion = "C:\Qtion\Heater\VupnoStnoskid.png"
                        Case Else
                        HeaterQtion = "C:\Qtion\Heater\VupnoStnoskid.png"
                    End Select
                End If
         End Select
             
        
        ' Insert Heater picture
         Set HeaterRng = .Bookmarks.Item("HeaterLoc").Range
         Set Heatershape = HeaterRng.InlineShapes.AddPicture(Filename:=HeaterQtion, LinkToFile:=False, SaveWithDocument:=True, Range:=HeaterRng)
           
         Set Heatershapetype = Heatershape.ConvertToShape
         With Heatershapetype
            .WrapFormat.Type = wdWrapSquare
            .WrapFormat.Side = wdWrapLeft
            .Left = wdShapeRight
         End With
         
         
         ' Insert Burner text
         .Bookmarks("BurnerFuelLoc").Range.Text = BurnerFuelQ
       
         ' Select Burner image
        Select Case BurnerMakeQ
            Case Is = "Maxon"
                Select Case BurnerModelQ
                        Case Is = "OVENPAK"
                        BurnerQtion = "C:\Qtion\Burner\OVENPAK.png"
                        Case Is = "KINEDIZER - 15% EA"
                        BurnerQtion = "C:\Qtion\Burner\LE.png"
                        Case Is = "EB-OVENPAK"
                        BurnerQtion = "C:\Qtion\Burner\OVENPAK.png"
                Case Else
                        BurnerQtion = "C:\Qtion\Burner\Other.png"
                End Select
            Case Is = "Webster"
                BurnerQtion = "C:\Qtion\Burner\Webster.png"
            Case Else
                BurnerQtion = "C:\Qtion\Burner\Other.png"
         End Select
         
       ' Insert Burner picture
        Set BurnerRng = .Bookmarks.Item("BurnerLoc").Range
        Set Burnershape = BurnerRng.InlineShapes.AddPicture(Filename:=BurnerQtion, LinkToFile:=False, SaveWithDocument:=True, Range:=BurnerRng)
    
        Set Burnershapetype = Burnershape.ConvertToShape
        With Burnershapetype
           .WrapFormat.Type = wdWrapSquare
           .WrapFormat.Side = wdWrapLeft
           .Left = wdShapeRight
        End With
        
        ' Insert panel picture
        Set ControlRng = .Bookmarks.Item("PanelLoc").Range
        Set Panelshape = ControlRng.InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\StdPanel.png", LinkToFile:=False, SaveWithDocument:=True, Range:=ControlRng)
    
        Set Panelshapetype = Panelshape.ConvertToShape
        With Panelshapetype
           .WrapFormat.Type = wdWrapSquare
           .WrapFormat.Side = wdWrapLeft
           .Left = wdShapeRight
        End With
        If TemplateQ = "Americas" Then
        Else
        .Bookmarks("PanelAreaClassLoc").Range.Text = PanelAreaClassQ
        .Bookmarks("NemaPanelLoc").Range.Text = NEMARatingQ
        .Bookmarks("MinAmbientLoc").Range.Text = AmbientTempMinQ
        .Bookmarks("MaxAmbientLoc").Range.Text = AmbientTempMaxQ
        .Bookmarks("VoltagePrimaryLoc").Range.Text = VoltageQ
        End If
        
       ' Insert controls
       Select Case CombControlQ
        Case Is = "Standard"
          If BMSQ = "Standard" Then
             With .ActiveWindow.Selection
             .GoTo What:=-1, Name:="ControlLoc"
             ' BMS standard
              Call BMSStandardQ(docWord)
             ' Comb control standard
              Call CombStandardQ(docWord)
             End With
          Else
            Select Case SpecialBMSQ
                Case Is = "HIMA"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS HIMA
                     Call BMSHIMAQ(docWord)
                    ' Comb control standard
                     Call CombStandardQ(docWord)
                    End With
                 Case Is = "GuardLogix"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS GuardLogix
                     Call BMSGuardlogixQ(docWord)
                    ' Comb control standard
                     Call CombStandardQ(docWord)
                    End With
                Case Is = "Other"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS OTHER
                    .TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
                    ' Comb control standard
                      Call CombStandardQ(docWord)
                    End With
                Case Else
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS GuardLogix
                     Call BMSGuardlogixQ(docWord)
                    ' Comb control standard
                     Call CombStandardQ(docWord)
                    End With
            End Select
          End If
        Case Else
          Select Case SpecialCombQ
           Case Is = "CompactLogix"
            If BMSQ = "Standard" Then
             With .ActiveWindow.Selection
             .GoTo What:=-1, Name:="ControlLoc"
             ' BMS standard
              Call BMSStandardQ(docWord)
             ' Comb control CompactLogix
              Call CombCompactLogixQ(docWord)
             Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
             End With
            Else
             Select Case SpecialBMSQ
                Case Is = "HIMA"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS HIMA
                      Call BMSHIMAQ(docWord)
                    ' Comb control CompactLogix
                     Call CombCompactLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
                Case Is = "GuardLogix"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS GuardLogix
                     Call BMSGuardlogixQ(docWord)
                    ' Comb control CompactLogix
                     Call CombCompactLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
                Case Is = "Other"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS OTHER
                    .TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
                    ' Comb control CompactLogix
                     Call CombCompactLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
                Case Else
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS GuardLogix
                     Call BMSGuardlogixQ(docWord)
                    ' Comb control CompactLogix
                     Call CombCompactLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
             End Select
            End If
           
           Case Is = "GuardLogix"
            If BMSQ = "Standard" Then
                 With .ActiveWindow.Selection
                 .GoTo What:=-1, Name:="ControlLoc"
                 ' BMS standard
                  Call BMSStandardQ(docWord)
                 ' Comb control GuardLogix
                 Call CombGuardLogixQ(docWord)
                Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                End With
            Else
             Select Case SpecialBMSQ
                Case Is = "HIMA"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS HIMA
                       Call BMSHIMAQ(docWord)
                    ' Comb control GuardLogix
                     Call CombGuardLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
                 Case Is = "GuardLogix"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS GuardLogix
                     Call BMSGuardlogixQ(docWord)
                    ' Comb control GuardLogix
                     Call CombGuardLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
                Case Is = "Other"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS OTHER
                    .TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
                    ' Comb control GuardLogix
                     Call CombGuardLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
                Case Else
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS OTHER
                    .TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
                    ' Comb control GuardLogix
                     Call CombGuardLogixQ(docWord)
                    Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                    End With
            End Select
          End If
    
           Case Else
            If BMSQ = "Standard" Then
             With .ActiveWindow.Selection
             .GoTo What:=-1, Name:="ControlLoc"
             ' BMS standard
              Call BMSStandardQ(docWord)
             ' Comb control PLC Other
             .TypeText (vbNewLine & "FILL IN WITH CUSTOM PLC INFO" & vbNewLine)
             End With
            Else
             Select Case SpecialBMSQ
                Case Is = "HIMA"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS HIMA
                        Call BMSHIMAQ(docWord)
                    ' Comb control PLC Other
                    .TypeText (vbNewLine & "FILL IN WITH CUSTOM PLC INFO" & vbNewLine)
                    End With
                 Case Is = "GuardLogix"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS GuardLogix
                     Call BMSGuardlogixQ(docWord)
                    ' Comb control PLC Other
                    .TypeText (vbNewLine & "FILL IN WITH CUSTOM PLC INFO" & vbNewLine)
                    End With
                Case Is = "Other"
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS OTHER
                    .TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
                    ' Comb control PLC Other
                    .TypeText (vbNewLine & "FILL IN WITH CUSTOM PLC INFO" & vbNewLine)
                    End With
                Case Else
                    With .ActiveWindow.Selection
                    .GoTo What:=-1, Name:="ControlLoc"
                    ' BMS OTHER
                    .TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
                    ' Comb control PLC Other
                    .TypeText (vbNewLine & "FILL IN WITH CUSTOM PLC INFO" & vbNewLine)
                    End With
             End Select
           End If
         End Select
        End Select
       
       ' Insert Fuel train
       
       .Bookmarks("BurnerFuelTLoc").Range.Text = BurnerFuelQ
       If BurnerFuelQ = "Gas" Then
        With .ActiveWindow.Selection
        .GoTo What:=-1, Name:="TrainLoc"
        .Style = docWord.Styles("StyleQ3")
        .TypeText ("Gas Train" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. 1 - Primary Pressure regulator" & vbNewLine & "Qty. 2 - Safety shut-off valves for primary line shut-off" & vbNewLine & "Qty. 1 - FO vent valve for primary line vent" & vbNewLine & "Qty. 2 - Manual isolation ball valves for primary line isolation" & vbNewLine & "Qty. 1 - Strainer for inlet fuel gas filtration" & vbNewLine & "Qty. 2 - High & Low gas pressure switch" & vbNewLine & "Qty. 2 – Pressure gauges with gauge valves for primary line indication" & vbNewLine & "Qty. 1 - Pilot pressure regulator" & vbNewLine _
        & "Qty. 2 - FC shut-off valve for pilot line shut-off" & vbNewLine & "Qty. 1 - FO vent valve for pilot line vent" & vbNewLine & "Qty. 2 - Manual isolation ball valves for pilot line isolation" & vbNewLine & "Qty. 1 - Pressure gauge with gauge valves for pilot line indication" & vbNewLine & "Pre-Piping & Pre-Wiring – The fuel train will be pre-piped from inlet isolation valve to outlet isolation valve.  All of the fuel train components will be pre-wired to the heater mounted control panel or junction box." & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
       Else
        With .ActiveWindow.Selection
        .GoTo What:=-1, Name:="TrainLoc"
        .Style = docWord.Styles("StyleQ3")
        .TypeText ("Oil Train" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. 2 - Safety shut-off valve for primary line shut-off" & vbNewLine & "Qty. 1 - Normally open solenoid vent valve for primary line vent" & vbNewLine & "Qty. 1 - Manual isolation ball valves for primary line isolation" & vbNewLine & "Qty. 1 - Duplex strainer for fuel oil filtration" & vbNewLine & "Qty. 1 - High pressure switch" & vbNewLine & "Qty. 1 - Low pressure switch" & vbNewLine & "Qty. 2 - Pressure gauges with pigtail and gauge" & vbNewLine & "Qty. 1 - Oil pump" & vbNewLine & "Qty. 1 - Pilot pressure regulator" & vbNewLine _
        & "Qty. 2 - FC shut-off valve for pilot line shut-off" & vbNewLine & "Qty. 1 - FO vent valve for pilot line vent" & vbNewLine & "Qty. 2 - Manual isolation ball valves for pilot line isolation" & vbNewLine & "Qty. 1 - Pressure gauge with gauge valves for pilot line indication" & vbNewLine & "Pre-Piping & Pre-Wiring – The fuel train will be pre-piped from inlet isolation valve to outlet isolation valve.  All of the fuel train components will be pre-wired to the heater mounted control panel or junction box." & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
       End If
       
       
       ' Define number of OverTemp Thermocouples and Orifice plates
       If HeaterFlowQ = "Series" Then
            DifPQ = HeaterInCoilQ
            OverTQ = DifPQ + 2
        Else
            DifPQ = HeaterInCoilQ + HeaterOutCoilQ
            OverTQ = DifPQ + 2
       End If
       
       ' Insert Heater instrumentation
        With .ActiveWindow.Selection
        .GoTo What:=-1, Name:="InstruHeaterLoc"
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. " & OverTQ & " - Thermocouples with thermowells for fluid temperature measurement" & vbNewLine & "Qty. 1 - Thermocouple for flue gas temperature measurement" & vbNewLine & "Qty. " & DifPQ & " - Flow orifice for differential pressure measurement" & vbNewLine & "Qty. " & DifPQ & " - DP switch with low and low-low flow switches for low fluid flow interlock")
        End With
        If HeaterCodeQ = "ASME Section I" Or HeaterCodeQ = "ASME Section 1 w/ABSA" Or HeaterCodeQ = "ASME Section 1 w/BCBB" Then
            With .ActiveWindow.Selection
            .GoTo What:=-1, Name:="InstruLooseLoc"
            .Style = docWord.Styles("StyleQ2")
            .TypeText ("Qty. 1 - Liquid filled pressure gauge for heat outlet pressure indication" & vbNewLine & "Qty. 2 - PSV for heater coil overpressure protection (per ASME Section I liquid relief)" & vbNewLine)
            .Style = docWord.Styles("StyleQ1")
            End With
        Else
            With .ActiveWindow.Selection
            .GoTo What:=-1, Name:="InstruLooseLoc"
            .Style = docWord.Styles("StyleQ2")
            .TypeText ("Qty. 1 - Liquid filled pressure gauge for heat outlet pressure indication" & vbNewLine & "Qty. 1 - PSV for heater coil overpressure protection (per ASME Section VIII liquid relief)" & vbNewLine)
            .Style = docWord.Styles("StyleQ1")
            End With
        End If

       ' Insert Pump Skids
       If PumpValveQ = "Bellows Seal Globe" Then
        PumpValveTextQ = " - Globe valve with butt weld connections and stainless steel bellows seal for pump"
        Else
        PumpValveTextQ = " - High temperature gate valve with butt weld connections for pump"
       End If
       
       If PumpSkidQ = "Yes" Then
        With .ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & QtyPumpsQ & " " & PumpOperQ & " Primary Loop Pump Skid & Valves - ")
        Select Case QtyPumpsQ
         Case Is = "1"
          Set Pumpshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Pumps\simplex.png", LinkToFile:=False, SaveWithDocument:=True)
          Set Pumpshapetype = Pumpshape.ConvertToShape
            With Pumpshapetype
               .WrapFormat.Type = wdWrapSquare
               .WrapFormat.Side = wdWrapLeft
               .Left = wdShapeRight
            End With
         Case Is = "2"
          Set Pumpshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Pumps\duplex.png", LinkToFile:=False, SaveWithDocument:=True)
          Set Pumpshapetype = Pumpshape.ConvertToShape
            With Pumpshapetype
               .WrapFormat.Type = wdWrapSquare
               .WrapFormat.Side = wdWrapLeft
               .Left = wdShapeRight
            End With
         Case Else
        End Select
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("A pump skid with a " & QtyPumpsQ & " " & PumpOperQ & " capacity primary system pump will be provided. The pump skid will be completely assembled on a structural steel skid frame. Isolation valves, strainers, check valves, expansion joints, and pressure gauges will all be provided as part of the assembled skid package. Butt weld connections will be used for all components (when possible) to minimize potential leak points. A summary of the equipment supplied with the primary loop pump skid is as follows:" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. " & QtyPumpsQ & " " & PumpOperQ & " Centrifugal thermal fluid pump & motor" & vbNewLine & "Qty. " & QtyPumpsQ & " - Air cooled mechanical seal" & vbNewLine & "Qty. " & QtyPumpsQ & PumpValveTextQ & " inlet isolation" & vbNewLine & "Qty. " & QtyPumpsQ & PumpValveTextQ & " outlet isolation and throttling" & vbNewLine & "Qty. " & QtyPumpsQ & " - Y-pattern strainer with butt weld connections and drain valve" & vbNewLine & "Qty. " & QtyPumpsQ & " - Drain valve" & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
        Else
       End If
             
       ' Insert Double Leg
       If DoubleLegQ = "By Others" Then
       Else
        Select Case DoubleLegQ
         Case Is = "Globe Valves"
             Call DoubleLegGlobeQ(docWord)
         Case Is = "Gate Valves"
             Call DoubleLegGateQ(docWord)
         Case Else
         End Select
       End If
       
       ' Insert Expansion tank
       If ExpTankGalQ = "By Others" Or OPExpTankQ = "Yes" Then
        Else
            Call ExpansionTankQ(docWord, ExpTankGalQ, ExpTankCodeQ, ExpTankLevelQ)
       End If
       
       'Insert Blanket
       If BlanketQ = "Yes" Then
        Call InertblanketQ(docWord, BlankettypeQ)
       Else
       End If
             
       ' Insert Motor Starters
       If MotorStarterQ = "By Sigma Thermal" Then
         Call MotorStarterinsertQ(docWord, VoltageQ, VoltageshortQ, NEMARatingQ, MotorStartershape, MotorStartershapetype)
       Else
       End If
       
       ' Insert Tank Tower
       If TankTowerQ = "Yes" Then
         Call TankTowerinsertQ(docWord, TankTowershape, TankTowershapetype)
        Else
       End If
       
       ' Insert Exhaust Stack
       If HeaterStackQ = "Free Standing" Or HeaterStackQ = "Heater Mounted" Then
            Call ExhaustStackinsertQ(docWord)
       Else
       End If
             
       ' Insert Drain Tank
       If DrainTankQ = "By Others" Or OPDrainTankQ = "Yes" Then
       Else
        Call DrainTankinsertQ(docWord, DrainTankQ, DrainTankCodeQ)
       End If
       
       'Insert drain pump
       If DrainPumpQ = "Sigma Thermal Standard" Or DrainPumpQ = "Custom" Then
        Call DrainPumpinsertQ(docWord)
       Else
       End If
       
       'Insert SideStream Filter
       If SideStreamQ = "Yes" Then
         Call SidestreaminsertQ(docWord)
       Else
       End If
       
       'Insert TCU
       If TCUQ = "Yes" Then
        Call TCUinsertQ(docWord, TCUshapetype, TCUshape)
        Else
       End If
       
       'Insert Shell/Tube HEX
       If STHEXQ = "Yes" Then
        Call STHEXinsertQ(docWord, STHEXshapetype, STHEXshape)
        Else
       End If
       
       'Insert Comb Air Pre-Heat
       If CombAirQ = "Yes" Then
        Call CombAirPreQ(docWord, CombAirshape, CombAirshapetype)
        Else
       End If
       
       'Insert Heat Tracings
        If FTHeatTQ = "Yes" Then
         Call FTHeatTinsertQ(docWord)
        Else
        End If
       
       'insert PSV Heat tracing
       If PSVHeatTQ = "Yes" Then
         Call PSVHeatTinsertQ(docWord)
        Else
       End If
       
       'Insert Piping Design/Stress Analysis
       If StressAnalysisQ = "Yes" Then
        Call StressAnalysisinsertQ(docWord, StressAnalysisshape)
        Else
       End If
              
       'Insert Burner House
       If BurnerHouseQ = "Yes" Then
        Call BurnerHouseinsertQ(docWord, BurnerHouseshape)
       Else
       End If
       
       'Insert Remote I/O
        If RemoteIOQ = "Yes" Then
         Call RemoteIOinsertQ(docWord, RemoteIOshape, RemoteIOshapetype)
        Else
       End If
       
       
       'Insert Exhaust Stack Oxygen Analyzer
       If O2AnalyzerQ = "Yes" Then
        Call O2AnalyzerinsertQ(docWord)
       Else
       End If
       
       'Insert Ladders & Platforms
       If LaddersPlatQ = "Yes" Then
        Call LadderPlatinsertQ(docWord)
       Else
       End If
       
       'Insert Low temperature Fuel Gas piping to an MDMT of -49 F
       If LowTempFGQ = "Yes" Then
        Call LowTempFGinsertQ(docWord)
       Else
       End If
       
       'Insert Canadian Codes
       If CanadianCodesQ = "Yes" Then
        Call CanadianCodesinsertQ(docWord)
       Else
       End If
       
      'Insert P.Eng. Stamp
       If PEngStampQ = "Yes" Then
        Call PEngStampinsertQ(docWord)
       Else
       End If
      
      'insert BlowerVFD
      If BlowerVFDQ = "Yes" Then
       Call BlowerVFDinsertQ(docWord)
      Else
      End If
      
      'insert PumpVFD
      If PumpVFDQ = "Yes" Then
       Call PumpVFDinsertQ(docWord, QtyPumpsQ)
       Else
      End If
       
    
      'Go to OPTIONS
      With .ActiveWindow.Selection
        .GoTo What:=-1, Name:="FirstOptionLoc"
      End With
      
      ' Optional Tank Tower
       If TankTowerQ = "Optional Scope" Then
         Call TankTowerinsertQ(docWord, TankTowershape, TankTowershapetype)
        Else
       End If
      
      'OPTIONAL Expansion tank
        If OPExpTankQ = "Yes" Then
         Call ExpansionTankQ(docWord, ExpTankGalQ, ExpTankCodeQ, ExpTankLevelQ)
        Else
        End If
        
      'OPTIONAL Double Leg Option
       If DoubleLegQ = "By Others" Then
       Else
        Select Case DoubleLegQ
        Case Is = "Optional - Globe"
         Call DoubleLegGlobeQ(docWord)
        Case Is = "Optional - Gate"
         Call DoubleLegGateQ(docWord)
        Case Else
        End Select
       End If
       
       'Optional blanket
       If BlanketQ = "Optional Scope" Then
        Call InertblanketQ(docWord, BlankettypeQ)
       Else
       End If
      
      'OPTIONAL controls
      If OPBMSQ <> "N/A" Then
        With .ActiveWindow.Selection
          .Style = docWord.Styles("StyleQ4")
          .TypeText (vbNewLine & "OPTIONAL Burner Management System" & vbNewLine)
          .Style = docWord.Styles("StyleQ5")
        End With
            Select Case OPBMSQ
              Case Is = "Standard"
                Call BMSStandardQ(docWord)
              Case Is = "HIMA"
                Call BMSHIMAQ(docWord)
              Case Is = "GuardLogix"
                Call BMSGuardlogixQ(docWord)
              Case Is = "Other"
                .ActiveWindow.Selection.TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
              Case Else
                .ActiveWindow.Selection.TypeText ("FILL IN WITH CUSTOM BMS INFO" & vbNewLine)
            End Select
      Else
      End If
      
      If OPCombQ <> "N/A" Then
        With .ActiveWindow.Selection
          .Style = docWord.Styles("StyleQ4")
          .TypeText (vbNewLine & "OPTIONAL Combustion Control System")
          .Style = docWord.Styles("StyleQ5")
        End With
            Select Case OPCombQ
              Case Is = "Standard"
                Call CombStandardQ(docWord)
              Case Is = "CompactLogix"
                Call CombCompactLogixQ(docWord)
                Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                With .ActiveWindow.Selection
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .MoveRight Unit:=wdCharacter, Count:=1
                .TypeText (vbNewLine)
                .Style = docWord.Styles("StyleQ1")
                End With
              Case Is = "GuardLogix"
                Call CombGuardLogixQ(docWord)
                Set PLCHMIshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\PLCHMI.png", LinkToFile:=False, SaveWithDocument:=True)
                With .ActiveWindow.Selection
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
                .MoveRight Unit:=wdCharacter, Count:=1
                .TypeText (vbNewLine)
                .Style = docWord.Styles("StyleQ1")
                End With
              Case Is = "Other"
                .ActiveWindow.Selection.TypeText ("FILL IN WITH CUSTOM Controller INFO" & vbNewLine)
              Case Else
                .ActiveWindow.Selection.TypeText ("FILL IN WITH CUSTOM Controller INFO" & vbNewLine)
            End Select
      Else
      End If
        
      ' OPTIONAL Pump Skid
       If PumpSkidQ = "Optional Scope" Then
        With .ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & QtyPumpsQ & " " & PumpOperQ & " OPTIONAL Primary Loop Pump Skid & Valves - ")
        Select Case QtyPumpsQ
         Case Is = "1"
          Set Pumpshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Pumps\simplex.png", LinkToFile:=False, SaveWithDocument:=True)
          Set Pumpshapetype = Pumpshape.ConvertToShape
            With Pumpshapetype
               .WrapFormat.Type = wdWrapSquare
               .WrapFormat.Side = wdWrapLeft
               .Left = wdShapeRight
            End With
         Case Is = "2"
          Set Pumpshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Pumps\duplex.png", LinkToFile:=False, SaveWithDocument:=True)
          Set Pumpshapetype = Pumpshape.ConvertToShape
            With Pumpshapetype
               .WrapFormat.Type = wdWrapSquare
               .WrapFormat.Side = wdWrapLeft
               .Left = wdShapeRight
            End With
         Case Else
        End Select
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("A pump skid with a " & QtyPumpsQ & " " & PumpOperQ & " capacity primary system pump will be provided. The pump skid will be completely assembled on a structural steel skid frame. Isolation valves, strainers, check valves, expansion joints, and pressure gauges will all be provided as part of the assembled skid package. Butt weld connections will be used for all components (when possible) to minimize potential leak points. A summary of the equipment supplied with the primary loop pump skid is as follows:" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. " & QtyPumpsQ & " " & PumpOperQ & " Centrifugal thermal fluid pump & motor" & vbNewLine & "Qty. " & QtyPumpsQ & " - Air cooled mechanical seal" & vbNewLine & "Qty. " & QtyPumpsQ & PumpValveTextQ & " inlet isolation" & vbNewLine & "Qty. " & QtyPumpsQ & PumpValveTextQ & " outlet isolation and throttling" & vbNewLine & "Qty. " & QtyPumpsQ & " - Y-pattern strainer with butt weld connections and drain valve" & vbNewLine & "Qty. " & QtyPumpsQ & " - Drain valve" & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
        Else
       End If
        
               
       'Optional Motor Starters
       If OPMotorStarterQ = "Yes" Then
         Call MotorStarterinsertQ(docWord, VoltageQ, VoltageshortQ, NEMARatingQ, MotorStartershape, MotorStartershapetype)
       Else
       End If
       
      ' Optional Exhaust Stack
       If HeaterStackQ = "Optional Scope" Then
            Call ExhaustStackinsertQ(docWord)
       Else
       End If
       
       'Optional drain tank
        If OPDrainTankQ = "Yes" Then
         Call DrainTankinsertQ(docWord, DrainTankQ, DrainTankCodeQ)
        Else
        End If
              
       'Optional drain pump
       If DrainPumpQ = "Optional Scope" Then
        Call DrainPumpinsertQ(docWord)
       Else
       End If
       
       'Optional SideStream Filter
       If SideStreamQ = "Optional Scope" Then
        Call SidestreaminsertQ(docWord)
       Else
       End If
       
       'Optional TCU
       If TCUQ = "Optional" Then
        Call TCUinsertQ(docWord, TCUshapetype, TCUshape)
        Else
       End If
       
       'Optional Shell/Tube HEX
       If STHEXQ = "Optional" Then
        Call STHEXinsertQ(docWord, STHEXshapetype, STHEXshape)
        Else
       End If
       
       'Optional Comb Air Pre-Heat
       If CombAirQ = "Optional Scope" Then
        Call CombAirPreQ(docWord, CombAirshape, CombAirshapetype)
        Else
       End If
       
        'Optional Heat Tracings
        If FTHeatTQ = "Optional" Then
         Call FTHeatTinsertQ(docWord)
        Else
        End If
       
       'Optional PSV Heat tracing
       If PSVHeatTQ = "Optional" Then
         Call PSVHeatTinsertQ(docWord)
        Else
       End If
       
      'Optional Piping Design/Stress Analysis
       If StressAnalysisQ = "Optional" Then
        Call StressAnalysisinsertQ(docWord, StressAnalysisshape)
        Else
       End If
       
      'Optional Burner House
       If BurnerHouseQ = "Optional Scope" Then
        Call BurnerHouseinsertQ(docWord, BurnerHouseshape)
       Else
       End If
        
      'Optional Remote I/O
        If RemoteIOQ = "Optional" Then
         Call RemoteIOinsertQ(docWord, RemoteIOshape, RemoteIOshapetype)
        Else
       End If
       
       'Optional Exhaust Stack Oxygen Analyzer
       If O2AnalyzerQ = "Optional" Then
        Call O2AnalyzerinsertQ(docWord)
       Else
       End If
    
       'Optional Ladders & Platforms
       If LaddersPlatQ = "Optional" Then
        Call LadderPlatinsertQ(docWord)
       Else
       End If
        
       'Optional Low temperature Fuel Gas piping to an MDMT of -49 F
       If LowTempFGQ = "Optional" Then
        Call LowTempFGinsertQ(docWord)
       Else
       End If
       
      'Optional Canadian Codes
       If CanadianCodesQ = "Optional" Then
        Call CanadianCodesinsertQ(docWord)
       Else
       End If
        
      'Optional P.Eng. Stamp
       If PEngStampQ = "Optional" Then
        Call PEngStampinsertQ(docWord)
       Else
       End If
    
      'Optional BlowerVFD
      If BlowerVFDQ = "Optional" Then
       Call BlowerVFDinsertQ(docWord)
      Else
      End If
      
      'Optional PumpVFD
      If PumpVFDQ = "Optional" Then
       Call PumpVFDinsertQ(docWord, QtyPumpsQ)
       Else
      End If
    
    
     ' Pricing Table
      With .ActiveWindow.Selection
       .GoTo What:=-1, Name:="PriceOptionLoc"
       
      If OPBMSQ = "N/A" Then
       .TypeText ("")
       Else
       If OPBMSQ = "Standard" Then
       .InsertRowsBelow 1
       .TypeText ("Standard Burner Management System")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
        Else
       .InsertRowsBelow 1
       .TypeText ("PLC Based Burner Management System")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       End If
      End If
      If OPCombQ = "N/A" Then
       .TypeText ("")
       Else
       If OPCombQ = "Standard" Then
       .InsertRowsBelow 1
       .TypeText ("Standard Combustion Control System")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
        Else
       .InsertRowsBelow 1
       .TypeText ("PLC Based Combustion Control System")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       End If
      End If
      
      If PumpSkidQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText (QtyPumpsQ & " " & PumpOperQ & " Primary Loop Pump Skid with motor/s and valve/s")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
      End If

      If OPExpTankQ = "Yes" Then
       .InsertRowsBelow 1
       .TypeText (ExpTankGalQ & " Gal Horizontal Cylindrical Expansion Tank")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
      End If

      If BlanketQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText (BlankettypeQ & " for the Expansion Tank Blanketing System")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
      Else
      End If
      
      If OPMotorStarterQ = "Yes" Then
       .InsertRowsBelow 1
       .TypeText ("Motor Starters")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
      Else
      End If
      
      If TankTowerQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText ("Expansion tank tower & interconnecting piping")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
      
       If HeaterStackQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText ("Exhaust Stack")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
        
       If OPDrainTankQ = "Yes" Then
       .InsertRowsBelow 1
       .TypeText (DrainTankQ & " Gal System Drain Tank")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If DrainPumpQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText ("Fill / Drain pump")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If SideStreamQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText ("Side stream filter package")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If TCUQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Temperture Control Unit Skid")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If STHEXQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Shell & Tube Heat Exchanger Skid")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
        
       If CombAirQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText ("Combustion air pre-heat package")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If FTHeatTQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Fuel train Heat Tracing")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If PSVHeatTQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("PSV Heat Tracing")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If StressAnalysisQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Piping Design & Stress Analysis")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
              
       If BurnerHouseQ = "Optional Scope" Then
       .InsertRowsBelow 1
       .TypeText ("Burner House Enclosure")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If RemoteIOQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Remote I/O")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If O2AnalyzerQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Flue Gas Oxygen Analyzer")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If LaddersPlatQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Access Ladders & Platforms")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If LowTempFGQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Low temperature Fuel Gas Train")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If

       If CanadianCodesQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Compliance with Canadian Codes")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
       
       If PEngStampQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Registered P. Eng. Stamp")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
         
       If BlowerVFDQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Blower VFD")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
      
       If PumpVFDQ = "Optional" Then
       .InsertRowsBelow 1
       .TypeText ("Pump Variable Frequency Drive/s")
       .MoveRight Unit:=wdCharacter, Count:=1
       .TypeText ("$ xxx.xx")
       Else
       End If
      
      End With
    
       'Insert Delivery dates
       Call InsertDeliveryQ(docWord, PIDMinQ, PIDMaxQ, GAMinQ, GAMaxQ, ReadyShMinQ, ReadyShMaxQ, PIDAppQ, GAAppQ)
        
       'SaveAs Quote # - Rev - Company name in Desktop
        On Error Resume Next
            If Range("B4") = 0 Then
                SaveNameQtion = Environ("UserProfile") & "\Desktop\" & Range("K4") & " - " & Range("B2") & ".docx"
                .SaveAs2 SaveNameQtion
                If Err.Number = 5356 Then
                MsgBox ("Can't overwrite an open doc." & vbNewLine & "Close any open word docs with the same name and rerun.")
                End If
                On Error GoTo 0
            Else
                SaveNameQtion = Environ("UserProfile") & "\Desktop\" & Range("K4") & "-" & Range("B4") & " - " & Range("B2") & ".docx"
                .SaveAs2 SaveNameQtion
                If Err.Number = 5356 Then
                MsgBox ("Can't overwrite an open doc." & vbNewLine & "Close any open word docs with the same name and rerun.")
                End If
                On Error GoTo 0
            End If
     
    End With
    
    Set applWord = Nothing
    
End Sub
Sub InsertDeliveryQ(docWord, PIDMinQ, PIDMaxQ, GAMinQ, GAMaxQ, ReadyShMinQ, ReadyShMaxQ, PIDAppQ, GAAppQ)
      With docWord.ActiveWindow.Selection
       .GoTo What:=-1, Name:="FirstDeliveryLoc"
       If PIDAppQ = "Yes" Then
        If GAAppQ = "Yes" Then
         .TypeText ("Submittal of P&IDs")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (PIDMinQ & " - " & PIDMaxQ & " weeks ARO")
         .InsertRowsBelow 1
         .TypeText ("Submittal of General Arrangements")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (GAMinQ & " - " & GAMaxQ & " weeks After Receipt of Approved P&IDs")
         .InsertRowsBelow 1
         .TypeText ("Equipment Ready to Ship")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (ReadyShMinQ & " - " & ReadyShMaxQ & " weeks After Receipt of Approved GAs")
        Else
         .TypeText ("Submittal of P&IDs")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (PIDMinQ & " - " & PIDMaxQ & " weeks ARO")
         .InsertRowsBelow 1
         .TypeText ("Submittal of General Arrangements")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (GAMinQ & " - " & GAMaxQ & " weeks After Receipt of Approved P&IDs")
         .InsertRowsBelow 1
         .TypeText ("Equipment Ready to Ship")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (ReadyShMinQ & " - " & ReadyShMaxQ & " weeks After Receipt of Approved P&IDs")
        End If
       Else
        If GAAppQ = "Yes" Then
         .TypeText ("Submittal of P&IDs")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (PIDMinQ & " - " & PIDMaxQ & " weeks ARO")
         .InsertRowsBelow 1
         .TypeText ("Submittal of General Arrangements")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (GAMinQ & " - " & GAMaxQ & " weeks ARO")
         .InsertRowsBelow 1
         .TypeText ("Equipment Ready to Ship")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (ReadyShMinQ & " - " & ReadyShMaxQ & " weeks After Receipt of Approved GAs")
        Else
         .TypeText ("Submittal of P&IDs")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (PIDMinQ & " - " & PIDMaxQ & " weeks ARO")
         .InsertRowsBelow 1
         .TypeText ("Submittal of General Arrangements")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (GAMinQ & " - " & GAMaxQ & " weeks ARO")
         .InsertRowsBelow 1
         .TypeText ("Equipment Ready to Ship")
         .MoveRight Unit:=wdCharacter, Count:=1
         .TypeText (ReadyShMinQ & " - " & ReadyShMaxQ & " weeks ARO")
        End If
       End If
      End With
End Sub


Sub CustomerscopeQ(docWord, MotorStarterQ, BlowerVFDQ, PumpVFDQ, HeaterStackQ)
 With docWord.ActiveWindow.Selection
    .GoTo What:=-1, Name:="FirstCustScopeLoc"
    .Style = docWord.Styles("StyleQ6")
    .TypeText (vbNewLine & "Unloading and placement of all Sigma Thermal supplied equipment" & vbNewLine)
    .TypeText ("Installation of all Sigma Thermal supplied equipment" & vbNewLine)
    .TypeText ("Design and installation of all system piping" & vbNewLine)
    .TypeText ("Connection of pre-piped fuel train outlet to burner inlet" & vbNewLine)
    .TypeText ("All piping, ducting, and exhaust stack insulation" & vbNewLine)
    .TypeText ("Design and installation foundations" & vbNewLine)
    .TypeText ("Mounting and interconnecting wiring of all loose valves and instruments" & vbNewLine)
    If HeaterStackQ = "Free Standing" Or HeaterStackQ = "Heater Mounted" Then
      .TypeText ("Mounting and connection of the exhaust stack" & vbNewLine)
    Else
      .TypeText ("Supply, mounting and connection of the exhaust stack" & vbNewLine)
    End If
    .TypeText ("Thermal fluid solution" & vbNewLine)
    .TypeText ("Power supply to motors and control panel" & vbNewLine)
    If MotorStarterQ = "By Sigma Thermal" Then
      .TypeText ("")
      If BlowerVFDQ = "Yes" Or PumpVFDQ = "Yes" Then
        .TypeText ("")
        Else
        .TypeText ("VFDs (if applicable)" & vbNewLine)
        End If
    Else
    .TypeText ("Motor starters and VFDs (if applicable)" & vbNewLine)
    End If
    .TypeText ("Fuel supply to fuel train at required inlet pressure" & vbNewLine)
    .TypeText ("PSV vent and/or discharge line connections" & vbNewLine)
    .TypeText ("User isolation and control valves (if required)" & vbNewLine)
    .TypeText ("PID control loops for all user control valves" & vbNewLine)
    .TypeText ("Transportation of all heating system components" & vbNewLine)
    .TypeText ("Verification of compliance with all required local codes" & vbNewLine)
   End With
End Sub
Sub PumpVFDinsertQ(docWord, QtyPumpsQ)
        With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Pump Variable Frequency Drive/s - ")
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("Qty. " & QtyPumpsQ & " - VFD/s will be provided to control the pump motor/s. The installation, configuration, and starters are provided by others. The following equipment will be provided, wired documented and tested:" & vbNewLine)
         .Style = docWord.Styles("StyleQ2")
         .TypeText ("All operator devices considered remote and supplied by customer" & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
        End With
End Sub

Sub BlowerVFDinsertQ(docWord)
        With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Blower Variable Frequency Drive - ")
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("A VFD will be provided to control the Fan motor. The installation, configuration, and starters are provided by others. The following equipment will be provided, wired documented and tested:" & vbNewLine)
         .Style = docWord.Styles("StyleQ2")
         .TypeText ("All operator devices considered remote and supplied by customer" & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
        End With
End Sub
Sub PEngStampinsertQ(docWord)
        With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Registered P. Eng. Stamp - ")
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("The following as-built drawings will be reviewed and stamped by a Province registered P. Eng.: GAs, P&IDs, and Building Drawings." & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub
Sub CanadianCodesinsertQ(docWord)
       With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Compliance with Canadian Codes - ")
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("Compliance with the following Canadian codes:" & vbNewLine)
         .Style = docWord.Styles("StyleQ2")
         .TypeText ("The following items are registered with ABSA; the heater coil, pipe fittings and valves" & vbNewLine & "The BMS and burner fuel piping trains are in accordance with CSA B149.3" & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
        End With
End Sub
Sub LowTempFGinsertQ(docWord)
        With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Upgrade - Fuel Gas Piping to an MDMT of -45 C (-49 F) - ")
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("The fuel gas piping with an actual MDMT rating of -29 C (-20F) will be upgraded to and MDMT rating of -45 C (-49 F) to prevent cold temperature damage." & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub

Sub LadderPlatinsertQ(docWord)
       With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Heater Access Ladders & Platforms - ")
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("This option includes the below list of access ladders and platforms, shop pre-fitted and broken out for shipping:" & vbNewLine)
         .Style = docWord.Styles("StyleQ2")
         .TypeText ("Qty. 1 - Access ladder and platform for the top front of the heater" & vbNewLine & "Qty. 1 - Access ladder and platform for the exhaust stack" & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
        End With
End Sub

Sub O2AnalyzerinsertQ(docWord)
        With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Exhaust Stack Oxygen Analyzer - ")
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("An exhaust stack oxygen analyzer is offered to measure the oxygen level in the exhaust stack." & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub

Sub RemoteIOinsertQ(docWord, RemoteIOshape, RemoteIOshapetype)
      With docWord.ActiveWindow.Selection
      .Style = docWord.Styles("StyleQ4")
      .TypeText (vbNewLine & "Remote I/O - ")
      .Style = docWord.Styles("StyleQ5")
      .TypeText ("The Remote I/O option is offered by Sigma Thermal to help decrease engineering time and installation cost due to multiple long wire runs, and instrument terminations. The Allen Bradley remote I/O platform is modular, reliable and has networking flexibility that includes Ethernet/IP, ControlNet, and DeviceNet protocols. This option is capable of handling up to 32 digital inputs, 16 digital outputs, 8 analog inputs and 4 analog outputs. The following is a list of those components:" & vbNewLine)
      Set RemoteIOshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Controls\remoteio.png", LinkToFile:=False, SaveWithDocument:=True)
      Set RemoteIOshapetype = RemoteIOshape.ConvertToShape
      With RemoteIOshapetype
         .WrapFormat.Type = wdWrapSquare
         .WrapFormat.Side = wdWrapLeft
         .Left = wdShapeRight
      End With
      .Style = docWord.Styles("StyleQ2")
      .TypeText ("Qty. 2 - Allen Bradley 1794 Flex I/O-16 Channel Digital Input" & vbNewLine & "Qty. 2 - Allen Bradley 1794 Flex I/O-8 Channel Digital Output" & vbNewLine & "Qty. 2 - Allen Bradley 1794 Flex I/O-4 Channel Analog Input" & vbNewLine & "Qty. 1 - Allen Bradley 1794 Flex I/O-4 Channel Analog Output" & vbNewLine & "Qty. (Lot) - Relays & Bases" & vbNewLine & "Qty. 1 - NEMA 4 Carbon Steel Enclosure (up sized from the standard)" & vbNewLine)
      .Style = docWord.Styles("StyleQ1")
      End With
End Sub
          
Sub BurnerHouseinsertQ(docWord, BurnerHouseshape)
     With docWord.ActiveWindow.Selection
     .Style = docWord.Styles("StyleQ4")
     .TypeText (vbNewLine & "Burner House - ")
     .Style = docWord.Styles("StyleQ5")
     .TypeText ("14' wide x 14' tall x 15' long self framing building with 4:12 gable roof, 22 ga steel exterior, and 24 ga aluminum interior with 6 mil vapor barrier. Wall insulation will be R12, roof insulation R20. A summary of the included components is as follows:" & vbNewLine)
     .Style = docWord.Styles("StyleQ2")
     .TypeText ("Qty. 1 - Exhaust fan" & vbNewLine & "Qty. 1 - Louvered dampers" & vbNewLine & "Qty. 1 - 10 kW Wall mounted electric heater" & vbNewLine & "Qty. 1 - Double Door 6' x 7'" & vbNewLine & "Qty. 1 - Flame detector" & vbNewLine & "Qty. 1 - Combustible gas detector" & vbNewLine & "Qty. 3 - Fluorescent light fixtures" & vbNewLine)
     .Style = docWord.Styles("StyleQ1")
    Set BurnerHouseshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\BurnerHouse\BurnerHouse.png", LinkToFile:=False, SaveWithDocument:=True)
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
    .MoveRight Unit:=wdCharacter, Count:=1
    .TypeText (vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End With
End Sub
        
Sub StressAnalysisinsertQ(docWord, StressAnalysisshape)
        With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Piping Design and Stress Analysis - ")
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("Piping design for the thermal oil system can be provided including a stress analysis per ASME B31.1 or B31.3 as required. Dimensional layout drawings, isometrics, support locations and support details can be provided. Stress analysis code compliance report, support loading and flange loads can also be provided." & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
        Set StressAnalysisshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\PipingAnalysis\stressanalysis.png", LinkToFile:=False, SaveWithDocument:=True)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .MoveRight Unit:=wdCharacter, Count:=1
        .TypeText (vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub
Sub FTHeatTinsertQ(docWord)
         With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Fuel train Heat Tracing - ")
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("The fuel train line can be heat traced to avoid the fuel gas from freezing, water condensation or other low temperature consequences along the line. The heat tracing system will include the following:" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. 1 - Electric heat tracing of the Fuel train" & vbNewLine & "Qty. 1 - Junction Box" & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub
Sub PSVHeatTinsertQ(docWord)
        With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "PSV Heat Tracing - ")
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("The heater pressure safety valve can be heat traced to avoid freezing, water condensation or other low temperature consequences. The heat tracing system will include the following:" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. 1 - Electric heat tracing of the PSV" & vbNewLine & "Qty. 1 - Junction Box" & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub

Sub CombAirPreQ(docWord, CombAirshape, CombAirshapetype)
        With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Combustion Air Pre-Heat System - ")
        Set CombAirshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Exchangers\combair.png", LinkToFile:=False, SaveWithDocument:=True)
        Set CombAirshapetype = CombAirshape.ConvertToShape
        With CombAirshapetype
           .WrapFormat.Type = wdWrapSquare
           .WrapFormat.Side = wdWrapLeft
           .Left = wdShapeRight
        End With
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("A combustion air pre-heat system to increase overall system efficiency and minimize system operating costs.  This system will utilize the heater exhaust gasses to pre-heat the incoming combustion air. This is a more efficiency utilization of the energy consumed, which results in lower natural gas operating costs. The estimated overall efficiency when using this system will exceed 95% (LHV basis). A summary of combustion air pre-heat system components is as follows:" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. 1 - Air to air heat exchanger" & vbNewLine & "Qty. 1 - Modified burner to accommodate elevated combustion air temperatures" & vbNewLine & "Combustion air ductwork from combustion fan to heat exchanger and from heat exchanger to heater. Exhaust gas ductwork from heater to heat exchanger, and from heat exchanger to stack (if applicable)" & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub


Sub STHEXinsertQ(docWord, STHEXshapetype, STHEXshape)
    With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Shell & Tube Heat Exchanger Skid - ")
        Set STHEXshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Exchangers\STHEX.png", LinkToFile:=False, SaveWithDocument:=True)
        Set STHEXshapetype = STHEXshape.ConvertToShape
        With STHEXshapetype
           .WrapFormat.Type = wdWrapSquare
           .WrapFormat.Side = wdWrapLeft
           .Left = wdShapeRight
        End With
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("A 1 x 100% ASME shell and tube heat exchanger will be provided rated for the process conditions. The exchanger will be assembled on a structural steel skid frame with the thermal fluid bypass piping and a 3-way control valve to control the flow of thermal fluid through the exchanger in order to maintain the specified process temperature. The thermal fluid piping will also come with manual isolation valves and safety relief valves for both the thermal fluid as well as the process stream. The Shell & Tube Heat Exchanger skid will contain the following:" & vbNewLine)
        .Style = docWord.Styles("StyleQ2")
        .TypeText ("Qty. 1 x 100% - ASME Shell and tube heat exchanger" & vbNewLine & "Qty. 1 - 3-way Automated control valve (controlled by OTHERS)" & vbNewLine & "Qty. 2 - Pressure safety valves" & vbNewLine & "Qty. 2 - Thermal fluid inlet/outlet isolation valves" & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With

End Sub
Sub TCUinsertQ(docWord, TCUshapetype, TCUshape)
    With docWord.ActiveWindow.Selection
    .Style = docWord.Styles("StyleQ4")
    .TypeText (vbNewLine & "Temperature Control Unit Skid - ")
    Set TCUshape = .InlineShapes.AddPicture(Filename:="C:\Qtion\Exchangers\TCU.png", LinkToFile:=False, SaveWithDocument:=True)
    Set TCUshapetype = TCUshape.ConvertToShape
    With TCUshapetype
       .WrapFormat.Type = wdWrapSquare
       .WrapFormat.Side = wdWrapLeft
       .Left = wdShapeRight
    End With
    .Style = docWord.Styles("StyleQ5")
    .TypeText ("The Temperature Control Unit skid will contain the following:" & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Qty. 1 - Shell and tube heat exchanger" & vbNewLine & "Qty. 1 - Circulation pump" & vbNewLine & "Qty. 1 - Junction Box for instrument landing" & vbNewLine & "Qty. 1 - Thermal oil temperature control valve (demand signal by OTHERS)" & vbNewLine & "Qty. 1 - Process stream temperature control valve (control signal by OTHERS)" & vbNewLine & "Qty. 2 - Expansion bellows" & vbNewLine & "Qty. 2 - Pressure gauges upstream and downstream of the pump" & vbNewLine & "Qty. 2 - Pump inlet and outlet isolation valves" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End With
End Sub

Sub SidestreaminsertQ(docWord)
    With docWord.ActiveWindow.Selection
     .Style = docWord.Styles("StyleQ4")
     .TypeText (vbNewLine & "Side Stream Filter - ")
     .Style = docWord.Styles("StyleQ5")
     .TypeText ("A thermal oil filtration package that continually filters the system oil. The filter is arranged in a side stream fashion, such that it may be taken out of service to change filter cartridges without shutting down the system. The filter package will be mounted on the primary pump skid package. A summary of the side stream filter system components is as follows:" & vbNewLine)
     .Style = docWord.Styles("StyleQ2")
     .TypeText ("Qty. 1 - Carbon steel housing with hinged cover plate" & vbNewLine & "Qty. 2 - Liquid filled pressure gauges with pigtail and gauge valves for pressure indication at the inlet and outlet of the filter housing" & vbNewLine & "Qty. 2 - 1"" Globe valves with butt weld connection, stainless steel bellows seal, and throttling cone for filter housing isolation and throttling" & vbNewLine & "Qty. 3 - Full set of replacement filter cartridges" & vbNewLine)
     .Style = docWord.Styles("StyleQ1")
    End With
End Sub

Sub DrainPumpinsertQ(docWord)
    With docWord.ActiveWindow.Selection
    .Style = docWord.Styles("StyleQ4")
    .TypeText (vbNewLine & "Thermal Oil Fill & Drain Pump - ")
    .Style = docWord.Styles("StyleQ5")
    .TypeText ("A fill and drain pump suitable for filling and draining thermal oil from the system will be provided. The pump will be a baseplate mounted reversible positive-displacement pump with motor and coupling. Qty. 2 - 1"" loose gate valves will be provided for pump isolation." & vbNewLine)
    End With
End Sub
Sub DoubleLegGlobeQ(docWord)
    With docWord.ActiveWindow.Selection
    .Style = docWord.Styles("StyleQ4")
    .TypeText (vbNewLine & "Loose Primary System Valves - ")
    .Style = docWord.Styles("StyleQ5")
    .TypeText ("The following valves will be supplied loose for installation in the primary heating loop piping." & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Qty. 1 - Globe valve with butt weld connections and stainless steel bellows seal for double leg drop isolation" & vbNewLine & "Qty. 1 - Globe valve with butt weld connection, stainless steel bellows seal, and throttling cone for double leg drop throttling" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End With
End Sub
Sub DoubleLegGateQ(docWord)
     With .ActiveWindow.Selection
    .Style = docWord.Styles("StyleQ4")
    .TypeText (vbNewLine & "Loose Primary System Valves - ")
    .Style = docWord.Styles("StyleQ5")
    .TypeText ("The following valves will be supplied loose for installation in the primary heating loop piping." & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Qty. 1 - High temperature gate valve with butt weld connections for double leg drop isolation" & vbNewLine & "Qty. 1 - High temperature gate valve with butt weld connections for double leg drop throttling" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End With
End Sub

Sub DrainTankinsertQ(docWord, DrainTankQ, DrainTankCodeQ)
    With docWord.ActiveWindow.Selection
    .Style = docWord.Styles("StyleQ4")
    .TypeText (vbNewLine & "Drain Tank - ")
    .Style = docWord.Styles("StyleQ5")
    .TypeText ("A " & DrainTankQ & "Gal system drain tank will be provided. ")
       If DrainTankCodeQ = "No" Then
       .TypeText ("Butt weld nozzle connections will be used where possible to minimize potential leak points." & vbNewLine)
       Else
       .TypeText ("This atmospheric tank will be designed and per " & DrainTankCodeQ & ", but not stamped. Butt weld nozzle connections will be used where possible to minimize potential leak points." & vbNewLine)
       End If
     End With
End Sub
Sub ExhaustStackinsertQ(docWord)
    With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Exhaust Stack - ")
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("An exhaust stack to direct the heater exhaust gasses outdoors. The stack will be designed with a false bottom that slopes to a drain if free standing." & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub

Sub TankTowerinsertQ(docWord, TankTowershape, TankTowershapetype)
    With docWord.ActiveWindow.Selection
        .Style = docWord.Styles("StyleQ4")
        .TypeText (vbNewLine & "Expansion Tank Tower & Pre-piping - ")
        Set TankTowershape = .InlineShapes.AddPicture(Filename:="C:\Qtion\TankTower\TankTower.png", LinkToFile:=False, SaveWithDocument:=True)
        Set TankTowershapetype = TankTowershape.ConvertToShape
        With TankTowershapetype
           .WrapFormat.Type = wdWrapSquare
           .WrapFormat.Side = wdWrapLeft
           .Left = wdShapeRight
        End With
        .Style = docWord.Styles("StyleQ5")
        .TypeText ("A tank tower can be provided to elevate the expansion tank above the top of either a simplex or duplex pump skid. The pump suction will be pre-piped to the expansion tank double leg drop, and the pump discharge will be extended to the skid edge as a single connection." & vbNewLine)
        .Style = docWord.Styles("StyleQ1")
        End With
End Sub
Sub MotorStarterinsertQ(docWord, VoltageQ, VoltageshortQ, NEMARatingQ, MotorStartershape, MotorStartershapetype)
    With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Single " & VoltageshortQ & " VAC Power Connection And Motor Starter Panels - ")
         Set MotorStartershape = .InlineShapes.AddPicture(Filename:="C:\Qtion\MotorStarters\MotorStarter.png", LinkToFile:=False, SaveWithDocument:=True)
         Set MotorStartershapetype = MotorStartershape.ConvertToShape
            With MotorStartershapetype
               .WrapFormat.Type = wdWrapSquare
               .WrapFormat.Side = wdWrapLeft
               .Left = wdShapeRight
            End With
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("This option includes the supply of a panel to accept a single " & VoltageQ & " power supply for the complete heater system, and individual motor starters. This option includes the following scope of supply." & vbNewLine)
         .Style = docWord.Styles("StyleQ6")
         .TypeText ("A motor starter panel will be supplied for the system motor/s. A summary of the components included in this panel are as follows:" & vbNewLine)
         .Style = docWord.Styles("StyleQ2")
         .TypeText ("Qty. 1 - NEMA 4 carbon steel enclosure" & vbNewLine & "Qty. 1 - Magnetic motor starter contactor" & vbNewLine & "Qty. 1 - Three-phase circuit breaker with thermal overload for primary motor protection" & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
         End With
End Sub

Sub InertblanketQ(docWord, BlankettypeQ)
Select Case BlankettypeQ
        Case Is = "Inert Gas"
         With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Nitrogen Blanket - ")
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("A pre-piped nitrogen blanket manifold will be provided to prevent oxygen from interacting with the fluid in the expansion tank. A summary of the nitrogen blanket equipment supplied is as follows:" & vbNewLine)
         .Style = docWord.Styles("StyleQ2")
         .TypeText ("Qty. 1 - Inlet nitrogen pressure regulator" & vbNewLine & "Qty. 1 - Low nitrogen pressure switch" & vbNewLine & "Qty. 1 - Vacuum breaker" & vbNewLine & "Qty. 1 - Liquid filled pressure gauge with gauge valve for expansion tank pressure indication" & vbNewLine & "Qty. 1 - PSV for tank overpressure protection" & vbNewLine & "Qty. 1 - Back pressure regulator for relief of excess gasses during system expansion" & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
         End With
        Case Is = "Fuel Gas"
         With docWord.ActiveWindow.Selection
         .Style = docWord.Styles("StyleQ4")
         .TypeText (vbNewLine & "Fuel Gas Blanket - ")
         .Style = docWord.Styles("StyleQ5")
         .TypeText ("A pre-piped Fuel Gas blanket manifold will be provided to prevent oxygen from interacting with the fluid in the expansion tank. A summary of the fuel gas blanket equipment supplied is as follows:" & vbNewLine)
         .Style = docWord.Styles("StyleQ2")
         .TypeText ("Qty. 1 - Inlet fuel gas pressure regulator" & vbNewLine & "Qty. 1 - Low fuel gas pressure switch" & vbNewLine & "Qty. 1 - Vacuum breaker" & vbNewLine & "Qty. 1 - Liquid filled pressure gauge with gauge valve for expansion tank pressure indication" & vbNewLine & "Qty. 1 - PSV for tank overpressure protection" & vbNewLine & "Qty. 1 - Back pressure regulator for relief of excess gasses during system expansion" & vbNewLine)
         .Style = docWord.Styles("StyleQ1")
         End With
         Case Else
        End Select
End Sub


Sub ExpansionTankQ(docWord, ExpTankGalQ, ExpTankCodeQ, ExpTankLevelQ)
    With docWord.ActiveWindow.Selection
    .Style = docWord.Styles("StyleQ4")
    .TypeText (vbNewLine & "Expansion Tank - ")
    .Style = docWord.Styles("StyleQ5")
    .TypeText ("A " & ExpTankGalQ & "Gal horizontal cylindrical expansion tank will be provided. The tank will be designed and stamped per " & ExpTankCodeQ & ". Butt weld nozzle connections will be used where possible to minimize potential leak points. A summary of the expansion tank equipment supplied is as follows:" & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Qty. 1 - Horizontal expansion tank with saddles" & vbNewLine & "Qty. 1 - Float type low level switches for expansion tank low level detection" & vbNewLine)
    If ExpTankLevelQ = "Optional" Or ExpTankLevelQ = "No" Then
    .Style = docWord.Styles("StyleQ1")
     Else
    .TypeText ("Qty. 1 - Level gauge for expansion tank level indication" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End If
    End With
End Sub

Sub BMSStandardQ(docWord)
    With docWord.ActiveWindow.Selection
    .TypeText ("A micro-processor based Burner Management System (BMS) will be supplied as part of the heater control system. This BMS system will provide proper burner sequencing, ignition and flame monitoring protection for the automatically ignited gas fired burner. A self-checking flame scanner will be provided as part of the new BMS system. The BMS system will be supplied with expanded text capabilities which will allow for detailed alarm messages to be displayed. In addition to local indication, the BMS comes equipped with a Modbus connection for data transmission back to a DCS or remote PLC. Yokogawa high temperature limit controllers will be used to monitor all temperature safety interlocks. Each temperature limit controller will have the capability of retransmitting the process variable back to the customer via a 4-20mA connection." & vbNewLine)
    End With
End Sub

Sub BMSHIMAQ(docWord)
    With docWord.ActiveWindow.Selection
    .TypeText ("The burner management system logic will be accomplished by the use of a programmable logic controller (PLC) mounted to the back panel of the main control panel. A HIMA HIMatrix processor and related hardware will be specified and used for the burner management system.  This processor will carry a SIL3 minimum certification to meet the requirements of NFPA & IEC standards as they relate to the use of PLC’s in burner management applications. This BMS system will provide proper burner sequencing, ignition and flame monitoring protection for the gas fired burner. A single Durag self-checking flame scanner will be provided as part of the BMS system. The BMS will be designed for a single fuel application with one flame scanner to monitor the pilot and main flame." & vbNewLine _
    & "The PLC will be programmed in HIMA’s ELOP II software using either ladder logic or function block formats. The program will be completely tested and provided to the customer in a PDF format at the completion of the project. The program logic will include, but not be limited to, the following:" & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Startup and shutdown sequence" & vbNewLine & "Purging sequence" & vbNewLine & "Safety interlock monitoring" & vbNewLine & "Pilot flame light off" & vbNewLine & "Main flame light off" & vbNewLine & "Flame monitoring" & vbNewLine & "First out alarm annunciation" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End With
End Sub

Sub BMSGuardlogixQ(docWord)
    With docWord.ActiveWindow.Selection
    .TypeText ("The burner management system logic will be accomplished by the use of a safety processor (PLC) mounted to the back panel of the unit control panel. An Allen-Bradley GuardLogix safety processor and related hardware will be specified and used for all of the burner management system logic. The system will be SIL2 capable. Design will be per NFPA, IEC and CSA standards as they relate to the use of PLC’s in burner management and safety applications. This BMS system will provide proper burner sequencing, ignition and flame monitoring protection for the gas fired burner. The BMS will be designed for a single fuel application with a single flame scanner to monitor the pilot and main flame. The PLC will be programmed with Rockwell Automation RS Logix 5000 software using function block format. The program will be completely tested and provided to the customer in a PDF format at the completion of the project." _
    & "The program logic will include, but not be limited to, the following:" & vbNewLine & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Startup and shutdown sequence" & vbNewLine & "Purging sequence" & vbNewLine & "Safety interlock monitoring" & vbNewLine & "Pilot flame light off" & vbNewLine & "Main flame light off" & vbNewLine & "Flame monitoring" & vbNewLine & "First out alarm annunciation" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End With
End Sub

Sub CombStandardQ(docWord)
    With docWord.ActiveWindow.Selection
    .TypeText (vbNewLine & "A temperature loop controller will be provided and configured to read a process variable input and adjust the firing rate of the system burner to achieve the desired outlet temperature or set point. The set point on this loop controller can be adjusted locally at the control panel or remotely via a 4-20mA signal.  High temperature limit controllers will be supplied and used to monitor system high temperature safeties as required to properly protect the process heater. Each high temperature limit will provide a 4-20mA signal for retransmission of its process variable. The high temperature limits will require local reset in the event of a high temperature trip. The following hardwired control signals/connections will be available to the customer." & vbNewLine & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Remote Temperature SP (4-20mA)" & vbNewLine & "High Temperature Controller PV Retransmission (4-20mA)" & vbNewLine & "Common BMS Alarm (Dry Contact)" & vbNewLine & "Remote Emergency Stop (Dry Contact)" & vbNewLine & "Remote Start/Stop (Dry Contact)" & vbNewLine & "Burner Interlocks as Required (Dry Contact)" & vbNewLine & "Burner Running (Dry Contact)" & vbNewLine & "BMS Operation Data-Modbus (RS485) Hardware and Programming by Customer" & vbNewLine & "Combustion Blower Start/Stop (Dry Contact)" & vbNewLine & "Combustion Blower Run Feedback (Dry Contact)" & vbNewLine & "Fluid Pump Run Feedback as Required (Dry Contacts)" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    End With
End Sub

Sub CombCompactLogixQ(docWord)
    With docWord.ActiveWindow.Selection
    .TypeText (vbNewLine & "The combustion control logic will be accomplished with the use of a programmable logic controller (PLC) mounted to the control panel back pan. The PLC will be configured to handle all of the system I/O as determined by the P&ID. All available process data will be gathered by the PLC and available for retransmission to the DCS via a network or hardwired connection. The PLC will be programmed in Allen Bradley RS5000 software using either ladder logic or function block formats. The program logic will include, but not be limited to, the following." & vbNewLine & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Process temperature control loop" & vbNewLine & "Burner Start/Stop and auto recycle logic" & vbNewLine & "Combustion air damper positioning for purge, light off, and normal modulation" & vbNewLine & "Parallel positioning combustion control logic" & vbNewLine & "Process control status and alarm monitoring" & vbNewLine & "Remote temperature set point" & vbNewLine & "First in alarm annunciation" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    .TypeText (vbNewLine & "A 12" & Chr(34) & " Maple operator interface (HMI) will be supplied and installed on the new control panel door. The HMI will provide local operator interface with the system data and controls. The HMI will display process information, system alarm status, loop controller information, and various control functions that can be accessed by plant personnel. The HMI will communicate with the PLC via Ethernet IP protocol. The HMI will have display screens developed for the project which will include but not be limited to, the following:" & vbNewLine & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Process Heater Overview Screen" & vbNewLine & "System Data Screen" & vbNewLine & "Alarm History Screen" & vbNewLine & "Temperature Loop Screen" & vbNewLine & "Maintenance Screen" & vbNewLine & "Fuel and Air Curves" & vbNewLine & "Fuel Train Overview Screen" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    .TypeText (vbNewLine)
    End With
End Sub

Sub CombGuardLogixQ(docWord)
    With docWord.ActiveWindow.Selection
    .TypeText (vbNewLine & "The combustion control logic will be accomplished by an AB GuardLogix programmable logic controller (PLC) mounted to the back panel of the control system. Will be used to handle all of the non-safety I/O. Air and fuel valve curves will be developed in the combustion control PLC program and setup by Sigma Thermal's service technician at the time of commissioning. The fuel and air curves may be accessed from the local operator interface on a password protected maintenance screen. All available process data will be gathered by the PLC and available for retransmission to the customer's DCS via the locally mounted Ethernet switch. The PLC will be programmed in Allen Bradley RS5000 software using ladder logic format. The program will be completely tested and provided to the customer in a PDF format at the completion of the project. The program logic will include, but not be limited to, the following:" & vbNewLine & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("System Overview" & vbNewLine & "Process temperature control loop-TIC" & vbNewLine & "Burner Start/Stop and auto recycle logic" & vbNewLine & "Combustion air damper positioning for purge, light off, and normal modulation" & vbNewLine & "Fuel Air Ratio Control logic" & vbNewLine & "Process control status and alarm monitoring" & vbNewLine & "Remote temperature set point" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    .TypeText (vbNewLine & "A 12" & Chr(34) & " Maple operator interface (HMI) will be supplied and installed on the new control panel door. The HMI will provide local operator interface with the system data and controls. The HMI will display process information, system alarm status, loop controller information, and various control functions that can be accessed by plant personnel. The HMI will communicate with the PLC via Ethernet IP protocol. The HMI will have display screens developed for the project which will include but not be limited to, the following:" & vbNewLine & vbNewLine)
    .Style = docWord.Styles("StyleQ2")
    .TypeText ("Process Heater Overview Screen" & vbNewLine & "System Data Screen" & vbNewLine & "Alarm History Screen" & vbNewLine & "Temperature Loop Screen" & vbNewLine & "Maintenance Screen" & vbNewLine & "Fuel and Air Curves" & vbNewLine & "Fuel Train Overview Screen" & vbNewLine)
    .Style = docWord.Styles("StyleQ1")
    .TypeText (vbNewLine)
    End With
End Sub

Function convertedPressure(SIunit As String, pressureValue As Single, EnglishUnit As String) As Single

If SIunit = "kPa (a)" And EnglishUnit = "psig" Then
    convertedPressure = pressureValue * 0.14504 - 14.7
ElseIf SIunit = "kPa (g)" And EnglishUnit = "psig" Then
    convertedPressure = (pressureValue + 101.325) * 0.14504 - 14.7
ElseIf SIunit = "in W.C." And EnglishUnit = "psig" Then
    convertedPressure = pressureValue * 0.03609 - 14.7
ElseIf SIunit = "mm W.C." And EnglishUnit = "psig" Then
    convertedPressure = pressureValue * 0.0014223 - 14.7
ElseIf SIunit = "kPa (a)" And EnglishUnit = "psid" Then
    convertedPressure = pressureValue * 0.14504
ElseIf SIunit = "kPa (g)" And EnglishUnit = "psid" Then
    convertedPressure = (pressureValue + 101.325) * 0.14504
ElseIf SIunit = "in W.C." And EnglishUnit = "psid" Then
    convertedPressure = pressureValue * 0.03609
ElseIf SIunit = "mm W.C." And EnglishUnit = "psid" Then
    convertedPressure = pressureValue * 0.0014223
ElseIf SIunit = "kPa (d)" And EnglishUnit = "psid" Then
    convertedPressure = pressureValue * 0.14504
ElseIf SIunit = "mm W.C." And EnglishUnit = "in W.C." Then
    convertedPressure = pressureValue * 0.039370079
Else
    convertedPressure = "Unknown conversion"
End If


End Function

Sub Button2_Click()

End Sub
Sub ConvertUnits()
Attribute ConvertUnits.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ConvertUnits Macro
    
'Convert power units
Worksheets("New Primary Inputs").Cells(7, "B").Select
    
If Worksheets("New Primary Inputs").Cells(3, "O") = "kW" Then
    Worksheets("New Primary Inputs").Cells(7, "B") = (Worksheets("New Primary Inputs").Cells(3, "N").Value) * 3412.14 / 1000000
ElseIf Worksheets("New Primary Inputs").Cells(3, "O") = "GJ/h" Then
    Worksheets("New Primary Inputs").Cells(7, "B") = (Worksheets("New Primary Inputs").Cells(3, "N").Value) * 947817.12 / 1000000
ElseIf Worksheets("New Primary Inputs").Cells(3, "O") = "kcal/h" Then
    Worksheets("New Primary Inputs").Cells(7, "B") = (Worksheets("New Primary Inputs").Cells(3, "N").Value) * 3.96567 / 1000000
Else
    Worksheets("New Primary Inputs").Cells(7, "B") = "Unknown Units"
End If
    

' Convert flowrate units
Worksheets("New Primary Inputs").Cells(8, "B").Select
    
If Worksheets("New Primary Inputs").Cells(4, "O") = "m3/h" Then
    Worksheets("New Primary Inputs").Cells(8, "B") = (Worksheets("New Primary Inputs").Cells(4, "N").Value) * 4.402868
ElseIf Worksheets("New Primary Inputs").Cells(4, "O") = "LPM" Then
    Worksheets("New Primary Inputs").Cells(8, "B") = (Worksheets("New Primary Inputs").Cells(4, "N").Value) / 3.7854118
ElseIf Worksheets("New Primary Inputs").Cells(4, "O") = "LPH" Then
    Worksheets("New Primary Inputs").Cells(8, "B") = (Worksheets("New Primary Inputs").Cells(4, "N").Value) / 227.124707
End If
    
    
' Convert temperature units
Worksheets("New Primary Inputs").Cells(14, "B").Select
Worksheets("New Primary Inputs").Cells(14, "B") = Worksheets("New Primary Inputs").Cells(9, "N") * 9 / 5 + 32

Worksheets("New Primary Inputs").Cells(22, "B").Select
Worksheets("New Primary Inputs").Cells(22, "B") = Worksheets("New Primary Inputs").Cells(12, "N") * 9 / 5 + 32

Worksheets("New Primary Inputs").Cells(40, "B").Select
Worksheets("New Primary Inputs").Cells(40, "B") = Worksheets("New Primary Inputs").Cells(21, "N") * 9 / 5 + 32

Worksheets("New Primary Inputs").Cells(58, "B").Select
Worksheets("New Primary Inputs").Cells(58, "B") = Worksheets("New Primary Inputs").Cells(38, "N") * 9 / 5 + 32

Worksheets("New Primary Inputs").Cells(6, "F").Select
Worksheets("New Primary Inputs").Cells(6, "F") = Worksheets("New Primary Inputs").Cells(15, "N") * 9 / 5 + 32

Worksheets("New Primary Inputs").Cells(7, "F").Select
Worksheets("New Primary Inputs").Cells(7, "F") = Worksheets("New Primary Inputs").Cells(16, "N") * 9 / 5 + 32

Worksheets("New Primary Inputs").Cells(35, "F").Select
Worksheets("New Primary Inputs").Cells(35, "F") = Worksheets("New Primary Inputs").Cells(24, "N") * 9 / 5 + 32


'Convert pressure units
Worksheets("New Primary Inputs").Cells(23, "B").Select
Worksheets("New Primary Inputs").Cells(23, "B") = Module11.convertedPressure(Worksheets("New Primary Inputs").Cells(13, "O"), Worksheets("New Primary Inputs").Cells(13, "N"), Worksheets("New Primary Inputs").Cells(23, "C"))

Worksheets("New Primary Inputs").Cells(26, "B").Select
Worksheets("New Primary Inputs").Cells(26, "B") = Module11.convertedPressure(Worksheets("New Primary Inputs").Cells(19, "O"), Worksheets("New Primary Inputs").Cells(19, "N"), Worksheets("New Primary Inputs").Cells(26, "C"))

Worksheets("New Primary Inputs").Cells(27, "B").Select
Worksheets("New Primary Inputs").Cells(27, "B") = Module11.convertedPressure(Worksheets("New Primary Inputs").Cells(20, "O"), Worksheets("New Primary Inputs").Cells(20, "N"), Worksheets("New Primary Inputs").Cells(27, "C"))

Worksheets("New Primary Inputs").Cells(41, "B").Select
Worksheets("New Primary Inputs").Cells(41, "B") = Module11.convertedPressure(Worksheets("New Primary Inputs").Cells(22, "O"), Worksheets("New Primary Inputs").Cells(22, "N"), Worksheets("New Primary Inputs").Cells(41, "C"))

Worksheets("New Primary Inputs").Cells(57, "B").Select
Worksheets("New Primary Inputs").Cells(57, "B") = Module11.convertedPressure(Worksheets("New Primary Inputs").Cells(37, "O"), Worksheets("New Primary Inputs").Cells(37, "N"), Worksheets("New Primary Inputs").Cells(57, "C"))

Worksheets("New Primary Inputs").Cells(60, "F").Select
Worksheets("New Primary Inputs").Cells(60, "F") = Module11.convertedPressure(Worksheets("New Primary Inputs").Cells(40, "O"), Worksheets("New Primary Inputs").Cells(40, "N"), Worksheets("New Primary Inputs").Cells(60, "G"))


'Convert length units
Worksheets("New Primary Inputs").Cells(43, "B").Select
Worksheets("New Primary Inputs").Cells(43, "B") = Application.WorksheetFunction.Convert(Worksheets("New Primary Inputs").Cells(26, "N"), Worksheets("New Primary Inputs").Cells(26, "O"), Worksheets("New Primary Inputs").Cells(43, "C"))

Worksheets("New Primary Inputs").Cells(8, "F").Select
If Worksheets("New Primary Inputs").Cells(17, "O") = "km" Then
    Worksheets("New Primary Inputs").Cells(8, "F") = Worksheets("New Primary Inputs").Cells(17, "N") * 3280.84
ElseIf Worksheets("New Primary Inputs").Cells(17, "O") = "m" Then
    Worksheets("New Primary Inputs").Cells(8, "F") = Worksheets("New Primary Inputs").Cells(17, "N") * 3.28084
Else
    Worksheets("New Primary Inputs").Cells(8, "F") = Application.WorksheetFunction.Convert(Worksheets("New Primary Inputs").Cells(17, "N"), Worksheets("New Primary Inputs").Cells(17, "O"), Worksheets("New Primary Inputs").Cells(8, "G"))
End If

Worksheets("New Primary Inputs").Cells(52, "F").Select
Worksheets("New Primary Inputs").Cells(52, "F") = Application.WorksheetFunction.Convert(Worksheets("New Primary Inputs").Cells(34, "N"), Worksheets("New Primary Inputs").Cells(34, "O"), Worksheets("New Primary Inputs").Cells(52, "G"))

Worksheets("New Primary Inputs").Cells(56, "F").Select
Worksheets("New Primary Inputs").Cells(56, "F") = Application.WorksheetFunction.Convert(Worksheets("New Primary Inputs").Cells(35, "N"), Worksheets("New Primary Inputs").Cells(35, "O"), Worksheets("New Primary Inputs").Cells(56, "G"))


'Convert volume units
Worksheets("New Primary Inputs").Cells(45, "B").Select
Worksheets("New Primary Inputs").Cells(45, "B") = Application.WorksheetFunction.Convert(Worksheets("New Primary Inputs").Cells(27, "N"), Worksheets("New Primary Inputs").Cells(27, "O"), Worksheets("New Primary Inputs").Cells(45, "C"))

Worksheets("New Primary Inputs").Cells(51, "B").Select
Worksheets("New Primary Inputs").Cells(51, "B") = Application.WorksheetFunction.Convert(Worksheets("New Primary Inputs").Cells(29, "N"), Worksheets("New Primary Inputs").Cells(29, "O"), Worksheets("New Primary Inputs").Cells(51, "C"))

Worksheets("New Primary Inputs").Cells(53, "B").Select
Worksheets("New Primary Inputs").Cells(53, "B") = Application.WorksheetFunction.Convert(Worksheets("New Primary Inputs").Cells(31, "N"), Worksheets("New Primary Inputs").Cells(31, "O"), Worksheets("New Primary Inputs").Cells(51, "C"))


End Sub

Sub LoadValves()

Dim design(20), operating(20), turndown(20) As Variant
Dim valve1, valve2, valve3 As String
Dim i, j, k, l, designCv, operatingCv, turndownCv As Integer
j = 0
k = 0
l = 0
designCv = Worksheets("Burner and Controls Equip Area").Cells(65, "AG").Value
operatingCv = Worksheets("Burner and Controls Equip Area").Cells(65, "AH").Value
turndownCv = Worksheets("Burner and Controls Equip Area").Cells(65, "AI").Value

For i = 7 To 58
    If designCv > Worksheets("Cv Tables").Cells(i, "H").Value And designCv < Worksheets("Cv Tables").Cells(i, "M").Value Then
        valve1 = Worksheets("Cv Tables").Cells(i, 3)
        design(j) = valve1
        j = j + 1
    End If
    
    If operatingCv > Worksheets("Cv Tables").Cells(i, "H").Value And operatingCv < Worksheets("Cv Tables").Cells(i, "M").Value Then
        valve2 = Worksheets("Cv Tables").Cells(i, 3)
        operating(k) = valve2
        k = k + 1
    End If
    
    If turndownCv > Worksheets("Cv Tables").Cells(i, "H").Value And turndownCv < Worksheets("Cv Tables").Cells(i, "M").Value Then
        valve3 = Worksheets("Cv Tables").Cells(i, 3)
        turndown(l) = valve3
        l = l + 1
    End If
Next i

Worksheets("Burner and Controls Equip Area").designComboBox.List = design
'Worksheets("Burner and Controls Equip Area").OperatingComboBox.List = operating
'Worksheets("Burner and Controls Equip Area").turndownComboBox.List = turndown

End Sub


-------------------------------------------------------------------------------
VBA MACRO Sheet4.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet4'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet7.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet7'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet8.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet8'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet15.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet15'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet17.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet17'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
-------------------------------------------------------------------------------
VBA MACRO Sheet18.cls 
in file: xl/vbaProject.bin - OLE stream: 'VBA/Sheet18'
- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
(empty macro)
+----------+--------------------+---------------------------------------------+
|Type      |Keyword             |Description                                  |
+----------+--------------------+---------------------------------------------+
|AutoExec  |workbook_open       |Runs when the Excel Workbook is opened       |
|AutoExec  |cmdBalance_Click    |Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|AutoExec  |ComboBox1_Change    |Runs when the file is opened and ActiveX     |
|          |                    |objects trigger events                       |
|Suspicious|Environ             |May read system environment variables        |
|Suspicious|Open                |May open a file                              |
|Suspicious|Output              |May write to a file (if combined with Open)  |
|Suspicious|Shell               |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|run                 |May run an executable file or a system       |
|          |                    |command                                      |
|Suspicious|Call                |May call a DLL using Excel 4 Macros (XLM/XLF)|
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
|Suspicious|VBA Stomping        |VBA Stomping was detected: the VBA source    |
|          |                    |code and P-code are different, this may have |
|          |                    |been used to hide malicious code             |
+----------+--------------------+---------------------------------------------+
VBA Stomping detection is experimental: please report any false positive/negative at https://github.com/decalage2/oletools/issues

