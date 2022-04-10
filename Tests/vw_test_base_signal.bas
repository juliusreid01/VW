Attribute VB_Name = "vw_test_base_signal"

Private shp as Shape
Private bSignal as vw_base_signal_c
Private wShape as visio_shape_wrapper_c

' initial the objects in this module where s is passed from the other tests here
Private Sub Initialize(Optional s as Shape = Nothing)
  If s Is Nothing Then Set shp = ActivePage.DrawLine(1, 10, 4, 10)
  shp.OpenSheetWindow

  If bSignal Is Nothing Then
    Set bSignal = new vw_base_signal_c
    Set bSignal.Shape = shp
    Set wShape = bSignal.Wrapper
  End If

  bSignal.Initialize SignalType.Signal
End Sub

Public Sub Review()
  ' review
  If MsgBox(Title:="Base Signal Test", Buttons:=vbYesNo, Prompt:="Review Signal?") = vbYes Then Stop
  shp.ContainingPage.PageSheet.OpenSheetWindow
  bSignal.Delete
  Set bSignal = Nothing
End Sub

' checks base shape and initial cells
Public Sub Test_Initialization(Optional s as Shape = Nothing)
  Initialize s

  vw_test_base_shape.Test_BaseShape bSignal.Base

  ' test user cell values
  If bSignal.SignalType <> SignalType.Signal Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Type Incorrect", "Read " & CStr(bSignal.SignalType) & " instead of 'Signal'!"
  If wShape.Result(S_CHILDOFFSET) <> 0.25 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Child Offset Incorrect", "Read " & CStr(wShape.Result(S_CHILDOFFSET)) & " instead of '0.25'"
  If wShape.Result(S_ACTIVEWIDTH) <> 0.25 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Active Width Incorrect", "Read " & CStr(wShape.Result(S_ACTIVEWIDTH)) & " instead of '0.25'"
  If wShape.Result(S_SKEWWIDTH) <> 0.025 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Skew Width Incorrect", "Read " & CStr(wShape.Result(S_SKEWWIDTH)) & " instead of '0.025'"
  If wShape.Result(S_PULSES) <> 0 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Pulses Incorrect", "Read " & CStr(wShape.Result(S_PULSES)) & " instead of '6'"
  If wShape.Result(S_BUSWIDTH) <> 1 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Bus Width Incorrect", "Read " & CStr(wShape.Result(S_BUSWIDTH)) & " instead of '0.25'"
  If wShape.Result(S_EDGES) <> 0 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Edges Incorrect", "Read " & CStr(wShape.Result(S_EDGES)) & " instead of '0.25'"
End Sub

'//TODO move to a general location
Public Function GeometryError(s as Shape, ParamArray coords() as Variant) as String
  Dim str as String
  For i = LBound(coords) to UBound(coords) Step 2
    str = str & coords(i) & ": Expected = " & coords(i+1) & ", Actual = " & s.Cells(coords(i)).Result("") & vbNewLine
  Next
  GeometryError = str & "Continue?"
End Function
'//TODO move to a general location
Public Function GeometryFail(s as Shape, ParamArray coords() as Variant) as Boolean
  GeometryFail = False
  For i = LBound(coords) to UBound(coords) Step 2
    GeometryFail = GeometryFail Or (s.Cells(coords(i)).Result("") <> coords(i+1))
  Next i
End Function

' test adding an event
Public Sub Test_AddEvent(Optional s as Shape = Nothing)
  Initialize s
  bSignal.AddEvent 0.5, vw_types.Edge Or vw_types.Node
  bSignal.UpdateEvents
  bSignal.DrawEvents

  If GeometryFail(shp, "Geometry1.X2", 0.5, "Geometry1.X3", 0.525, "Geometry1.Y2", 0, "Geometry1.Y3", 0.25) Then
    If MsgBox(Title:="Base Signal Test: Geometry Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:=GeometryError(shp, "Geometry1.X2", 0.5, "Geometry1.X3", 0.625, "Geometry1.Y2", 0, "Geometry1.Y3", 0.25)) = vbNo Then Stop
  End If
  Review
End Sub

' run the full suite
Public Sub Test_BaseSignal()
  Initialize
  Test_Initialization
  ' event testing
  bSignal.AddEvent "Width/2"
  bSignal.AddEvent "Prop.Delay"
  bSignal.AddEvent 2.5
  bSignal.RemoveEvent 2.5
  bSignal.AddEvent 2.75
  bSignal.UpdateEvents
  bSignal.DrawEvents

  bSignal.AddEvent .5, vw_types.GateZ
  bSignal.UpdateEvents
  bSignal.DrawEvents
  ' test geometry 4 and 5
  If shp.Cells("Geometry1.X4").Result("") <> 0.5 Or shp.Cells("Geometry1.X5").Result("") <> 0.5 Then _
    If MsgBox(Title:="Base Signal Test: Geometry Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected 0.5, Actual X4 = " & shp.Cells("Geometry1.X4").Result("") & vbNewLine & _
                      "Expected 0.5, Actual X5 = " & shp.Cells("Geometry1.X5").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop
  bSignal.RemoveEvent .5, vw_types.GateZ
  bSignal.UpdateEvents
  bSignal.DrawEvents

  bSignal.AddEvent .6, vw_types.Gate0
  bSignal.UpdateEvents
  bSignal.DrawEvents
  ' test geometry 4 and 5
  If shp.Cells("Geometry1.X4").Result("") <> 0.6 Or shp.Cells("Geometry1.X5").Result("") <> 0.6 Then _
    If MsgBox(Title:="Base Signal Test: Geometry Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected 0.6, Actual X4 = " & shp.Cells("Geometry1.X4").Result("") & vbNewLine & _
                      "Expected 0.6, Actual X5 = " & shp.Cells("Geometry1.X5").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop
  bSignal.RemoveEvent .6, vw_types.Gate0
  bSignal.UpdateEvents
  bSignal.DrawEvents

  bSignal.AddEvent 2, vw_types.Gate1
  bSignal.UpdateEvents
  bSignal.DrawEvents
  ' test geometry 4 and 5
  If shp.Cells("Geometry1.X6").Result("") <> 2 Or shp.Cells("Geometry1.X7").Result("") <> 2 Then _
    If MsgBox(Title:="Base Signal Test: Geometry Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected 2, Actual X6 = " & shp.Cells("Geometry1.X6").Result("") & vbNewLine & _
                      "Expected 2, Actual X7 = " & shp.Cells("Geometry1.X7").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop
  bSignal.RemoveEvent 2, vw_types.Gate1
  bSignal.UpdateEvents
  bSignal.DrawEvents

  If shp.RowCount(visSectionScratch) <> 3 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Add Edge", "Add Edge Sub did not correctly add 3 rows to Scratch"
  If shp.Cells("Scratch.X1").Result("") <> 0.125 Or shp.Cells("Scratch.Y1").Result("") <> 0.25 Then _
    If MsgBox(Title:="Base Signal Test: Scratch Row1 Mismatch",  Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 0.125, Actual Scratch.X1 = " & shp.Cells("Scratch.X1").Result("") & vbNewLine & _
                      "Expected: 0.25, Actual Scratch.Y1 = " & shp.Cells("Scratch.Y1").Result("") _
                        & vbNewLine & "Continue?") = vbNo Then Stop
  If shp.Cells("Scratch.X2").Result("") <> 1.5 Or shp.Cells("Scratch.Y2").Result("") <> 0 Then _
    If MsgBox(Title:="Base Signal Test: Scratch Row2 Mismatch",  Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 1.5, Actual Scratch.X2 = " & shp.Cells("Scratch.X2").Result("") & vbNewLine & _
                      "Expected: 0, Actual Scratch.Y2 = " & shp.Cells("Scratch.Y2").Result("") _
                        & vbNewLine & "Continue?") = vbNo Then Stop
  If shp.Cells("Scratch.X3").Result("") <> 2.75 Or shp.Cells("Scratch.Y3").Result("") <> 0.25 Then _
    If MsgBox(Title:="Base Signal Test: Scratch Row3 Mismatch",  Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 2.75, Actual Scratch.X3 = " & shp.Cells("Scratch.X3").Result("") & vbNewLine & _
                      "Expected: 0.25, Actual Scratch.Y3 = " & shp.Cells("Scratch.Y3").Result("") _
                        & vbNewLine & "Continue?") = vbNo Then Stop

  If shp.Cells("Geometry1.X2").Result("") <> 0.125 Or shp.Cells("Geometry1.Y2").Result("") <> 0 Then _
    If MsgBox(Title:="Base Signal Test: Geometry1 Row2 Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 0.125, Actual Geometry1.X2 = " & shp.Cells("Geometry1.X2").Result("") & vbNewLine & _
                      "Expected: 0, Actual Geometry1.Y2 = " & shp.Cells("Geometry1.Y2").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop
  If shp.Cells("Geometry1.X3").Result("") <> 0.15 Or shp.Cells("Geometry1.Y3").Result("") <> 0.25 Then _
    If MsgBox(Title:="Base Signal Test: Geometry1 Row2 Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 0.15, Actual Geometry1.X3 = " & shp.Cells("Geometry1.X3").Result("") & vbNewLine & _
                      "Expected: 0.25, Actual Geometry1.Y3 = " & shp.Cells("Geometry1.Y3").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop

  If shp.Cells("Geometry1.X4").Result("") <> 1.5 Or shp.Cells("Geometry1.Y4").Result("") <> 0.25 Then _
    If MsgBox(Title:="Base Signal Test: Geometry1 Row2 Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 1.5, Actual Geometry1.X4 = " & shp.Cells("Geometry1.X4").Result("") & vbNewLine & _
                      "Expected: 0.25, Actual Geometry1.Y4 = " & shp.Cells("Geometry1.Y4").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop
  If shp.Cells("Geometry1.X5").Result("") <> 1.525 Or shp.Cells("Geometry1.Y5").Result("") <> 0 Then _
    If MsgBox(Title:="Base Signal Test: Geometry1 Row2 Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 1.525, Actual Geometry1.X5 = " & shp.Cells("Geometry1.X5").Result("") & vbNewLine & _
                      "Expected: 0, Actual Geometry1.Y5 = " & shp.Cells("Geometry1.Y5").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop
  If shp.Cells("Geometry1.X6").Result("") <> 2.75 Or shp.Cells("Geometry1.Y6").Result("") <> 0 Then _
    If MsgBox(Title:="Base Signal Test: Geometry1 Row2 Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 2.75, Actual Geometry1.X6 = " & shp.Cells("Geometry1.X6").Result("") & vbNewLine & _
                      "Expected: 0, Actual Geometry1.Y6 = " & shp.Cells("Geometry1.Y6").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop
  If shp.Cells("Geometry1.X7").Result("") <> 2.775 Or shp.Cells("Geometry1.Y7").Result("") <> 0.25 Then _
    If MsgBox(Title:="Base Signal Test: Geometry1 Row2 Mismatch", Buttons:=vbYesNo + vbQuestion, _
              Prompt:="Expected: 2.775, Actual Geometry1.X7 = " & shp.Cells("Geometry1.X7").Result("") & vbNewLine & _
                      "Expected: 0.25, Actual Geometry1.Y7 = " & shp.Cells("Geometry1.Y7").Result("") _
                      & vbNewLine & "Continue?") = vbNo Then Stop

  bSignal.AddEvent "Prop.Delay", vw_types.Node
  bSignal.AddEvent 2.75, vw_types.Node
  bSignal.UpdateEvents
  bSignal.DrawEvents
  Review()
End Sub

