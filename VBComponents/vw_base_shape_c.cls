VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_base_shape_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' class to apply shape user cells and data

' handle to the Visio Shape
Private vsoShp as Shape
' handle to the Visio Shape Wrapper
Private wShape as visio_shape_wrapper_c

Private Sub Class_Initialize()
  Set wShape = new visio_shape_wrapper_c
End Sub

Private Sub Class_Terminate()
  Set vsoShp = Nothing
  Set wShape = Nothing
End Sub

Public Property Set Shape(s as Shape)
  Set wShape.vsoShape = s
  Set vsoShp = s
End Property

Public Property Get Shape() as Shape
  Set Shape = vsoShp
End Property

Public Property Get Wrapper() as visio_shape_wrapper_c
  Set Wrapper = wShape
End Property

Public Sub Initialize()
  RemoveSections
  AddUserCells
  AddDataCells
  ResetVisioShape
  vsoShp.Text = vsoShp.Name
End Sub

Public Sub Delete()
  wShape.Delete
  Set vsoShp = Nothing
  Set wShape = Nothing
End Sub

Private Sub RemoveSections()
  vsoShp.DeleteSection visSectionProp
  vsoShp.DeleteSection visSectionUser
  vsoShp.DeleteSection visSectionControls
  vsoShp.DeleteSection visSectionActions
End Sub

' add every possible user cell and let the child shape delete the unused
Private Sub AddUserCells()
  With wShape
    .AddUserCell S_TYPE, """Type of shape this is"""
    .AddUserCell S_CHILDOFFSET, """TBD minimum distance between children"""
    .AddUserCell S_ACTIVEWIDTH
    .AddUserCell S_SKEWWIDTH
    .AddUserCell S_PULSES
    .AddUserCell S_BUSWIDTH
    .AddUserCell S_EDGES
    .AddUserCell S_TEST, """Spare cell to test formula/value combinations"""
  End With
End Sub

' add every possible data cell and let the child shape delete the unused
' each *Data* sub calls the next one
Private Sub AddDataCells()
  With wShape
    .AddDataCell S_NAME, """Enter a name for this object"""
    .SetCell CellName:=S_NAME, Position:="Verify", v:=True
    .SetCell S_NAME, Chr(34) & vsoShp.Name & Chr(34)

    .AddDataCell S_CLOCK, """Reference clock"""
    .AddDataCell S_SIGNAL, """Reference signal"""
    .AddDataCell S_ACTIVELOW
    .AddDataCell S_PERIOD
    .AddDataCell S_SKEW, """Percentage of period/2 to delay edges"""
    .AddDataCell S_DELAY
    .AddDataCell S_DUTYCYCLE
    .AddDataCell S_SIGNALSKEW, """Amount of skew to apply to dependent signals"""
    .AddDataCell S_EVENTTYPE, """Select an event type to add/modify"""
    .AddDataCell S_EVENTTRIGGER, """Select a trigger type for the event"""
    .AddDataCell S_EVENTPOSITION, """Input the positon based on the trigger"""
    .AddDataCell S_LABELEDGES, """Select which labels to show on transitions"""
    .AddDataCell S_LABELSIZE
    .AddDataCell S_LABELFONT
  End With
  SetDataLabels
End Sub

' called by AddDataCells
Private Sub SetDataLabels()
  With wShape
    .SetCell Position:="Label", CellName:=S_ACTIVELOW, v:=S_LABEL_ACTIVELOW
    .SetCell Position:="Label", CellName:=S_SKEW, v:=S_LABEL_SKEW
    .SetCell Position:="Label", CellName:=S_DUTYCYCLE, v:=S_LABEL_DUTYCYCLE
    .SetCell Position:="Label", CellName:=S_SIGNALSKEW, v:=S_LABEL_SIGNALSKEW
    .SetCell Position:="Label", CellName:=S_EVENTTYPE, v:=S_LABEL_EVENTTYPE
    .SetCell Position:="Label", CellName:=S_EVENTTRIGGER, v:=S_LABEL_EVENTTRIGGER
    .SetCell Position:="Label", CellName:=S_EVENTPOSITION, v:=S_LABEL_EVENTPOSITION
    .SetCell Position:="Label", CellName:=S_LABELEDGES, v:=S_LABEL_LABELEDGES
    .SetCell Position:="Label", CellName:=S_LABELSIZE, v:=S_LABEL_LABELSIZE
    .SetCell Position:="Label", CellName:=S_LABELFONT, v:=S_LABEL_LABELFONT
  End With
  SetDataTypes
End Sub

' called by SetDataLabels
Private Sub SetDataTypes()
  With wShape
    .SetCell Position:="Type", CellName:=S_NAME, v:=visPropTypeString
    .SetCell Position:="Type", CellName:=S_CLOCK, v:=visPropTypeListFix
    .SetCell Position:="Type", CellName:=S_SIGNAL, v:=visPropTypeListFix
    .SetCell Position:="Type", CellName:=S_ACTIVELOW, v:=visPropTypeBool
    .SetCell Position:="Type", CellName:=S_PERIOD, v:=visPropTypeNumber
    .SetCell Position:="Type", CellName:=S_SKEW, v:=visPropTypeNumber
    .SetCell Position:="Type", CellName:=S_DELAY, v:=visPropTypeNumber
    .SetCell Position:="Type", CellName:=S_DUTYCYCLE, v:=visPropTypeNumber
    .SetCell Position:="Type", CellName:=S_SIGNALSKEW, v:=visPropTypeNumber
    .SetCell Position:="Type", CellName:=S_EVENTTYPE, v:=visPropTypeListFix
    .SetCell Position:="Type", CellName:=S_EVENTTRIGGER, v:=visPropTypeListFix
    .SetCell Position:="Type", CellName:=S_EVENTPOSITION, v:=visPropTypeListVar
    .SetCell Position:="Type", CellName:=S_LABELEDGES, v:=visPropTypeListVar
    .SetCell Position:="Type", CellName:=S_LABELSIZE, v:=visPropTypeListNumber
    .SetCell Position:="Type", CellName:=S_LABELFONT, v:=visPropTypeListNumber
  End With
  SetDataFormats
End Sub

' called by SetDataTypes
Private Sub SetDataFormats()
  With wShape
    .SetCell Position:="Format", CellName:=S_PERIOD, v:="""0.00 u"""
    .SetCell Position:="Format", CellName:=S_SKEW, v:="""0.0 u"""
    .SetCell Position:="Format", CellName:=S_DELAY, v:="""0.000 u"""
    .SetCell Position:="Format", CellName:=S_DUTYCYCLE, v:="""0.0 u"""
    .SetCell Position:="Format", CellName:=S_SIGNALSKEW, v:="""0.000 u"""
    .SetCell Position:="Format", CellName:=S_LABELSIZE, v:="""0.000 u"""
    .SetCell Position:="Format", CellName:=S_LABELFONT, v:="""0 u"""
  End With
  SetDataHidden
End Sub

' called by SetDataFormat
Private Sub SetDataHidden()
  With wShape
    .SetCell Position:="Invisible", CellName:=S_CLOCK, v:="STRSAME("""", Prop." & S_CLOCK & ".Format)"
    .SetCell Position:="Invisible", CellName:=S_SIGNAL, v:="STRSAME("""", Prop." & S_SIGNAL & ".Format)"
    .SetCell Position:="Invisible", CellName:=S_EVENTTRIGGER, v:="STRSAME("""", Prop." & S_EVENTTYPE & ")"
    .SetCell Position:="Invisible", CellName:=S_EVENTPOSITION, v:="OR(Prop." & S_EVENTTRIGGER & ".Invisible,STRSAME("""", Prop." & S_EVENTTRIGGER & "))"
    .SetCell Position:="Invisible", CellName:=S_LABELEDGES, v:="STRSAME("""", Prop." & S_LABELEDGES & ".Format)"
    .SetCell Position:="Invisible", CellName:=S_LABELSIZE, v:="OR(Prop." & S_LABELEDGES & ".Invisible, STRSAME(""None"", Prop." & S_LABELEDGES & ", True))"
    .SetCell Position:="Invisible", CellName:=S_LABELFONT, v:="Prop." & S_LABELSIZE & ".Invisible"
  End With
End Sub

' Reset the Visio shape cells
Public Sub ResetVisioShape(Optional s as Shape = Nothing)
  If Not s is Nothing Then Set Me.Shape = s

  ' prevent one dimensional shapes from rotating or getting bent out of shape
  If vsoShp.OneD = True Then
    wShape.EndY = "Guard(BeginY)"
    wShape.Width = "Guard(EndX-BeginX)"
    wShape.PinX = "Guard(BeginX)"
    wShape.PinY = "Guard(BeginY)"
  End If

  wShape.Height = "BlockSizeY"
  wShape.LocPinX = 0
  wShape.LocPinY = "If(FlipY, Height*1, Height*0)"
  wShape.FlipY = "Prop.ActiveLow"

  vsoShp.Cells("Para.HorzAlign").Formula = visHorzRight
  vsoShp.Cells("TxtPinX").Formula = "Width*0"
  vsoShp.Cells("TxtLocPinX").Formula = "TxtWidth"
  vsoShp.Cells("TxtWidth").Formula = "LEN(SHAPETEXT(TheText))*Char.Size"
  vsoShp.Cells("FillForegndTrans").Formula = 1
  vsoShp.Cells("Geometry1.NoFill").Formula = False
  ResetGeometry
End Sub

' called by ResetVisioShape to reset the geometry
Public Sub ResetGeometry(Optional s as Shape = Nothing)
  Dim Height as Variant

  If Not s is Nothing Then Set Me.Shape = s

  Height = vsoShp.Cells("Height").Formula

  vsoShp.DeleteSection visSectionConnectionPts
  vsoShp.DeleteSection visSectionScratch

  Do While vsoShp.GeometryCount > 1
    vsoShp.DeleteSection visSectionFirstComponent + 1
  Loop

  Do While vsoShp.RowCount(visSectionFirstComponent) > 3
    vsoShp.DeleteRow visSectionFirstComponent, visRowComponent + 2
  Loop

  wShape.SetPoint visSectionFirstComponent, visRowComponent + 1, "Width*0", "Height*0"
  wShape.SetPoint visSectionFirstComponent, visRowComponent + 2, "Width*1", "Height*0"
  vsoShp.RowType(visSectionFirstComponent, visRowComponent + 1) = visTagMoveTo
  vsoShp.RowType(visSectionFirstComponent, visRowComponent + 2) = visTagLineTo
  vsoShp.Cells("Height").FormulaForce = Height
End Sub