VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_base_signal_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' handle to the Visio Shape
Private vsoShp as Shape
' handle to the Visio Shape Wrapper
Private wShape as visio_shape_wrapper_c
' handle to the base_shape
Private bShape as vw_base_shape_c
' type of signal
Private pSignalType as SignalType

Private Sub Class_Initialize()
  set bShape = new vw_base_shape_c
End Sub

Private Sub Class_Terminate()
  Set vsoShp = Nothing
  Set wShape = Nothing
  set bShape = Nothing
End Sub

Public Property Set Shape(s as Shape)
  Set bShape.Shape = s
  Set vsoShp = s
  Set wShape = bShape.Wrapper()
End Property

Public Property Get Wrapper() as visio_shape_wrapper_c
  Set Wrapper = wShape
End Property

Public Sub Initialize(pType as SignalType)
  bShape.Initialize
  pSignalType = pType
  SetUserCells
  SetDataCells
End Sub

Public Sub Delete()
  bShape.Delete()
  Set vsoShp = Nothing
  Set wShape = Nothing
  set bShape = Nothing
End Sub

Private Sub SetUserCells()
  With wShape
    .SetUserCell S_TYPE, Chr(34) & pSignalType & Chr(34)
    .SetUserCell S_CHILDOFFSET, "BlockSizeY"
    .SetUserCell S_ACTIVEWIDTH, VW_0 & "+(" & .GetCellName(S_PERIOD) & "*" & .GetCellName(S_DUTYCYCLE) & ")"
    .SetUserCell S_SKEWWIDTH, VW_0 & "+(" & .GetCellName(S_SKEW) & "*" & .GetCellName(S_PERIOD) & "*0.5)"
    .SetUserCell S_PULSES, "CEILING((Width-Geometry1.X1)/" & .GetCellName(S_PERIOD) & ")"
    .SetUserCell S_BUSWIDTH, 1
    .SetUserCell S_EDGES, False
  End With
End Sub

Private Sub SetDataCells()
  With wShape
    .SetDataCell S_ACTIVELOW, False
    .SetDataCell S_PERIOD, "BlockSizeX*2"
    .SetDataCell S_SKEW, "10 %"
    .SetDataCell S_DELAY, "BlockSizeX*0.5"
    .SetDataCell S_DUTYCYCLE, "50 %"
    .SetDataCell S_SIGNALSKEW, .GetCellName(S_SKEWWIDTH)
    .SetDataCell S_EVENTTYPE, "Index(0," & .GetCellName(S_EVENTTYPE) & ".Format)"
    .SetDataCell S_LABELEDGES, "Index(0," & .GetCellName(S_LABELEDGES) & ".Format)"
    .SetDataCell S_LABELSIZE, "BlockSizeX*0.667"
    .SetDataCell S_LABELFONT, "8 pt"

    .SetDataCell Position:="Format", Name:=S_LABELEDGES, v:=Chr(34) & S_LIST_NONE & Chr(34)
    .SetDataCell Position:="Format", Name:=S_EVENTTRIGGER, v:=Chr(34) & S_LIST_ABSOLUTE & Chr(34)
  End With
End Sub