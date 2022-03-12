VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_Signal_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class module to wrap around a Visio Shape used to convert line shapes into signals

' handle to the shape being modified
Private shp as Shape
' the type of signal
Private sType as SignalType
' index of the current label
Private LabelIdx as Integer

Private Sub Class_Initialize()
  vw_cfg.Configure
End Sub

Private Sub Class_Terminate()
  Set shp = Nothing
End Sub

' Create a new signal of the given type on the input page
Public Sub NewSignal(p as Page, t as SignalType, x1 as double, y1 as double)
  set shp = p.DrawLine(x1, y1, x1 + VW_WIDTH, y1)
  Init
End Sub

' Convert a shape to the given type on the input page
Public Sub ConvertSignal(s as Shape, t as SignalType)
  set shp = s
  RemoveSections
  Init
End Sub

' Initialize the shape
Private Sub Init()
  sType = t

  AddUserCells
  SetUserCells

  AddDataCells
  ResetData
  SetDataPrompts
  SetDataLabels
  SetDataTypes
  SetDataFormats
  SetDataHidden

  ResetVisioData
  ResetGeometry
End Sub

' Remove sections from the shapesheet
Private Sub RemoveSections()
  shp.DeleteSection visSectionProp
  shp.DeleteSection visSectionUser
  shp.DeleteSection visSectionControls
  shp.DeleteSection visSectionActions
  shp.DeleteSection visSectionConnectionPts
End Sub

' Add the User-defined Cells
Private Sub AddUserCells()
  shp.AddNamedRow visSectionUser, "Type", visTagDefault
  shp.AddNamedRow visSectionUser, "ChildOffset", visTagDefault
  if sType = SignalType.Bus or sType = SignalType.Bit then
    shp.AddNamedRow visSectionUser, "BusWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "SkewWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "Edges", visTagDefault
  else
    shp.AddNamedRow visSectionUser, "ActiveWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "SkewWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "Pulses", visTagDefault
  end if
  shp.AddNamedRow visSectionUser, "Test", visTagDefault
End Sub

' Set the User-defined Cells
Private Sub SetUserCells()
  shp.Cells("User.Type").Formula = VW_TYPE_STR(SignalType.Clock)
  shp.Cells("User.ChildOffset").Formula = "BlockSizeY"
  if sType = SignalType.Bit then
    shp.Cells("User.BusWidth").Formula = 1
    shp.Cells("User.Edges").Formula = False
  elseif sType = SignalType.Bus then
    shp.Cells("User.BusWidth").Formula = 8
    shp.Cells("User.Edges").Formula = False
  end if
  shp.Cells("User.Test.Prompt").Formula = """Spare cell to test formulae and values"""
End Sub

' Add the Shape Data Cells
Private Sub AddDataCells()
  shp.AddNamedRow visSectionProp, "Name", visTagDefault
  shp.Cells("Prop.Name.Verify").Formula = True
  shp.AddNamedRow visSectionProp, "Clock", visTagDefault
  shp.AddNamedRow visSectionProp, "Signal", visTagDefault
  if sType = SignalType.Clock or sType = SignalType.Bit then
    shp.AddNamedRow visSectionProp, "ActiveLow", visTagDefault
  end if
  if sType = SignalType.Clock then
    shp.AddNamedRow visSectionProp, "Period", visTagDefault
    shp.AddNamedRow visSectionProp, "Skew", visTagDefault
  end if
  shp.AddNamedRow visSectionProp, "Delay", visTagDefault
  if sType = SignalType.Clock then
    shp.AddNamedRow visSectionProp, "DutyCycle", visTagDefault
    shp.AddNamedRow visSectionProp, "SignalSkew", visTagDefault
  end if
  shp.AddNamedRow visSectionProp, "Trigger", visTagDefault
  shp.AddNamedRow visSectionProp, "Events", visTagDefault
  shp.AddNamedRow visSectionProp, "LabelEdges", visTagDefault
  shp.AddNamedRow visSectionProp, "LabelSize", visTagDefault
  shp.AddNamedRow visSectionProp, "LabelFont", visTagDefault
End Sub

' Set the Shape Data Cell Defaults
Public Sub ResetData(Optional s as Shape = Nothing)
  if Not s Is Nothing then set shp = s

  if sType = SignalType.Clock or sType = SignalType.Bit then
    shp.Cells("Prop.ActiveLow").Formula = False
  end if
  if sType = SignalType.Clock then
    shp.Cells("Prop.Period").Formula = "BlockSizeX*2"
    shp.Cells("Prop.Skew").Formula = "10 %"
    shp.Cells("Prop.Delay").Formula = "BlockSizeX*0.5"
    shp.Cells("Prop.DutyCycle").Formula = "50 %"
    shp.Cells("Prop.SignalSkew").Formula = "BlockSizeX*0.1"
  else
    shp.Cells("Prop.Delay").Formula = 0
  end if
  shp.Cells("Prop.LabelEdges").Formula = "Index(0,Prop.LabelEdges.Format)"
  shp.Cells("Prop.LabelSize").Formula = "BlockSizeX*0.667"
  shp.Cells("Prop.LabelFont").Formula = "8 pt"
End Sub

' Set the Shape Data Cell Prompts
Private Sub SetDataPrompts()
  shp.Cells("Prop.Name.Prompt").Formula = """Signal Name"""
  shp.Cells("Prop.Clock.Prompt").Formula = """Reference Clock"""
  shp.Cells("Prop.Signal.Prompt").Formula = """Reference Signal"""
  if sType = SignalType.Clock then
    shp.Cells("Prop.Skew.Prompt").Formula = """Percentage of period/2 to delay transitions"""
    shp.Cells("Prop.Delay.Prompt").Formula = """Initial clock delay"""
    shp.Cells("Prop.SignalSkew.Prompt").Formula = """Amount of skew to apply to dependent signals"""
  else
    shp.Cells("Prop.Delay.Prompt").Formula = """Amount of delay before each transition"""
  end if
  shp.Cells("Prop.LabelEdges.Prompt").Formula = """Select which labels to show on transitions"""
End Sub

' Set the Shape Data Cell Labels
Private Sub SetDataLabels()
  shp.Cells("Prop.ActiveLow.Label").Formula = """Active Low"""
  if sType = SignalType.Clock then
    shp.Cells("Prop.Skew.Label").Formula = """Skew %"""
    shp.Cells("Prop.DutyCycle.Label").Formula = """Duty Cycle %"""
  end if
  shp.Cells("Prop.LabelEdges.Label").Formula = """Label Edges"""
  shp.Cells("Prop.LabelSize.Label").Formula = """Label Size"""
  shp.Cells("Prop.LabelFont.Label").Formula = """Label Font Pt"""
End Sub

' Set the Shape Data Cell Types
Private Sub SetDataTypes()
  shp.Cells("Prop.Name.Type").Formula = visPropTypeString
  shp.Cells("Prop.Clock.Type").Formula = visPropTypeListFix
  shp.Cells("Prop.Signal.Type").Formula = visPropTypeListFix
  if sType = SignalType.Clock or sType = SignalType.Bit then
    shp.Cells("Prop.ActiveLow.Type").Formula = visPropTypeBool
  end if
  if sType = SignalType.Clock then
    shp.Cells("Prop.Period.Type").Formula = visPropTypeNumber
    shp.Cells("Prop.Skew.Type").Formula = visPropTypeNumber
  end if
  shp.Cells("Prop.Delay.Type").Formula = visPropTypeNumber
  if sType = SignalType.Clock then
    shp.Cells("Prop.DutyCycle.Type").Formula = visPropTypeNumber
    shp.Cells("Prop.SignalSkew.Type").Formula = visPropTypeNumber
  end if
  shp.Cells("Prop.Trigger.Type").Formula = visPropTypeListFix
  shp.Cells("Prop.Events.Type").Formula = visPropTypeNumber
  shp.Cells("Prop.LabelEdges.Type").Formula = visPropTypeListVar
  shp.Cells("Prop.LabelSize.Type").Formula = visPropTypeNumber
  shp.Cells("Prop.LabelFont.Type").Formula = visPropTypeNumber
End Sub

' Set the Shape Data Cell Formats
Private Sub SetDataFormats()
  if sType = SignalType.Clock then
    shp.Cells("Prop.Period.Format").Formula = """0.0 u"""
    shp.Cells("Prop.SignalSkew.Format").Formula = """0.0 u"""
    shp.Cells("Prop.Skew.Format").Formula = """0.0 u"""
    shp.Cells("Prop.DutyCycle.Format").Formula = """0.0 u"""
    shp.Cells("Prop.LabelEdges.Format").Formula = """None;All"""
  end if
  shp.Cells("Prop.LabelFont.Format").Formula = """0 u"""
  shp.Cells("Prop.Delay.Format").Formula = """0.0 u"""
End Sub

' Set the Shape Data Cells Invisible field
Private Sub SetDataHidden()
  shp.Cells("Prop.Clock.Invisible").Formula = "STRSAME("""", Prop.Clock.Format)"
  shp.Cells("Prop.Signal.Invisible").Formula = "STRSAME("""", Prop.Clock.Format)"
  shp.Cells("Prop.LabelSize.Invisible").Formula = "OR(Prop.LabelEdges.Invisible, STRSAME(""None"", Prop.LabelEdges))"
  shp.Cells("Prop.LabelFont.Invisible").Formula = "OR(Prop.LabelEdges.Invisible, STRSAME(""None"", Prop.LabelEdges))"
End Sub

' Reset the Visio Shape Data
Public Sub ResetVisioData(Optional s as Shape = Nothing)
  if Not s is Nothing then set shp = s

  ' only available in 1D shapes
  if shp.CellExists("EndY", visExistsLocally) <> False then
    shp.Cells("EndY").FormulaForce = "Guard(BeginY)"
    shp.Cells("Width").FormulaForce = "Guard(EndX-BeginX)"
  end if

  shp.Cells("Height").Formula = "BlockSizeY"
  shp.Cells("PinX").FormulaForce = "Guard(BeginX)"
  shp.Cells("PinY").Formula = "Guard(BeginY)"
  shp.Cells("LocPinX").Formula = 0
  shp.Cells("LocPinY").Formula = "If(FlipY,Height*1,Height*0)"
  if sType = SignalType.Bus then
    select case VW_BUS_YANCHOR
     case visAlignMiddle
      shp.Cells("LocPinY").Formula = "Height*0.5"
     case visAlignTop
      shp.Cells("LocPinY").Formula = "If(FlipY,Height*0,Height*1)"
    end select
  end if
  if sType = SignalType.Bit or sType = SignalType.Clock then
    shp.Cells("FlipY").Formula = "Prop.ActiveLow"
  end if
  shp.Cells("Para.HorzAlign").Formula = visHorzRight
  shp.Cells("TxtPinX").Formula = "Width*0"
  shp.Cells("TxtWidth").Formula = "LEN(SHAPETEXT(TheText))*Char.Size"
  shp.Cells("FillForegndTrans").Formula = 1
  shp.Cells("Geometry1.NoFill").Formula = False
End Sub

Public Sub ResetGeometry(Optional s as Shape = Nothing)
  if Not s is Nothing then set shp = s

  shp.DeleteSection visSectionConnectionPts

  For i = 0 to shp.RowCount(visSectionScratch) - 1
    For Each s In shp.ContainingPage.Shapes
      If s.Name = shp.CellsSRC(visSectionScratch, i, VW_COL_LABEL_NAME).ResultStr("") Then s.Delete
    Next
  Next i

  shp.DeleteScratch visSectionScratch

  Do While shp.GeometryCount > 1
    shp.DeleteSection visSectionFirstComponent + 1
  Loop

  Do While shp.RowCount(visSectionFirstCompont) > 3
    ' delete row 2 to keep X1=0 and X2=Width
    shp.DeleteRow visSectionFirstComponent, visRowComponent + 2
  Loop
End Sub

Private Sub AddGeometry(Optional Section as Integer = -1, Optional NoShow as Variant = False, _
                        Optional NoFill, Optional NoLine, Optional NoSnap as Boolean = False)
  If Section < 0 then
    Section = shp.AddSection(visSectionFirstComponent + shp.GeometryCount)
    shp.AddRow Section, visRowComponent, visTagComponent
  End If
  shp.CellsSRC(Section, visRowComponent, visNoFill + 0).Formula = NoFill
  shp.CellsSRC(Section, visRowComponent, visNoFill + 1).Formula = NoLine
  shp.CellsSRC(Section, visRowComponent, visNoFill + 1).Formula = NoShow
  shp.CellsSRC(Section, visRowComponent, visNoFill + 1).Formula = NoSnap
End Sub

Private Sub StartGeometry()
  shp.CellsSRC(visSectionFirstComponent, visRowComponent, visNoFill + 0).Formula = False
  shp.CellsSRC(visSectionFirstComponent, visRowComponent, visNoFill + 1).Formula = False
  shp.CellsSRC(visSectionFirstComponent, visRowComponent, visNoFill + 1).Formula = False
  shp.CellsSRC(visSectionFirstComponent, visRowComponent, visNoFill + 1).Formula = False
  if sType = SignalType.Bus Then
    shp.AddSection visSectionFirstComponent + 1
    shp.AddRow visSectionFirstComponent + 1, visRowComponent, visTagComponent
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent, visNoFill + 0).Formula = True
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent, visNoFill + 1).Formula = False
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent, visNoFill + 1).Formula = False
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent, visNoFill + 1).Formula = False
    shp.AddRow visSectionFirstComponent + 1, visRowComponent + 1, visTagMoveTo
    shp.AddRow visSectionFirstComponent + 1, visRowComponent + 2, visTagLineTo
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent + 1, visX).Formula = "Geometry1.X1"
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent + 1, visY).Formula = "Height-Geometry1.Y1"
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent + 2, visX).Formula = "Geometry1.X2"
    shp.CellsSRC(visSectionFirstComponent + 1, visRowComponent + 2, visY).Formula = "Height-Geometry1.Y2"
  end if
End Sub
