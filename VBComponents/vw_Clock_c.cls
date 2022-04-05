VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_clock_c"
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
' handle to the base signals
Private bSignal as vw_base_signal_c

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
  Set vsoShp = Nothing
  Set wShape = Nothing
  set bShape = Nothing
  set bSignal = Nothing
End Sub

Public Property Set Shape(s as Shape)
  set bSignal = new vw_base_signal_c
  Set bSignal.Shape = s
  Set bShape = bSignal.Base
  Set vsoShp = s
  Set wShape = bSignal.Wrapper
End Property

Public Property Get Wrapper() as visio_shape_wrapper_c
  Set Wrapper = wShape
End Property

Public Property Get Base() as vw_base_signal_c
  Set Base = bSignal
End Property

Public Sub Initialize
  bSignal.Initialize vw_types.Clock
End Sub

Public Sub Delete()
  bSignal.Delete
End Sub

Public Sub Reset()
  bShape.ResetVisioShape
  bShape.ResetGeometry
  DoEdges
End Sub

Public Sub DoEdges()
  Dim Row as Integer
  Dim Remainder as Integer: Remainder = 0

  Row = 0

  wShape.SetCell S_TEST, wShape.GetCellName(S_DELAY)
  Do While wShape.Result(S_TEST) <= wShape.Width
    bSignal.AddEdge vsoShp.Cells( wShape.GetCellName(S_TEST) ).Formula
    Row = Row + 1
    If (Row Mod 2) = Remainder Then
      wShape.SetCell S_TEST, "Scratch.X" & CStr(Row) & "+" & wShape.GetCellName(S_ACTIVEWIDTH)
    Else
      wShape.SetCell S_TEST, "Scratch.X" & CStr(Row) & "+" & wShape.GetCellName(S_PERIOD) & "-" & wShape.GetCellName(S_ACTIVEWIDTH)
    End If
  Loop

  bSignal.UpdateEvents
End Sub
