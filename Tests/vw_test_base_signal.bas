Attribute VB_Name = "vw_test_base_signal"

Public Sub Test_BaseSignal()
  Dim shp as Shape
  Dim bSignal as vw_base_signal_c

  Set shp = ActivePage.DrawLine(1, 10, 4, 10)
  shp.OpenSheetWindow

  Set bSignal = new vw_base_signal_c
  Set bSignal.Shape = shp
  bSignal.Initialize SignalType.Signal

  vw_test_base_shape.Test_BaseShape bSignal.Base

  ' test user cell values
  If bSignal.SignalType <> SignalType.Signal Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Type Incorrect", "Read " & CStr(bSignal.SignalType) & " instead of 'Signal'!"
  If bSignal.ChildOffset <> 0.25 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Child Offset Incorrect", "Read " & CStr(bSignal.ChildOffset) & " instead of '0.25'"
  If bSignal.ActiveWidth <> 0.25 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Active Width Incorrect", "Read " & CStr(bSignal.ActiveWidth) & " instead of '0.25'"
  If bSignal.SkewWidth <> 0.025 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Skew Width Incorrect", "Read " & CStr(bSignal.SkewWidth) & " instead of '0.025'"
  If bSignal.Pulses <> 0 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Pulses Incorrect", "Read " & CStr(bSignal.Pulses) & " instead of '6'"
  If bSignal.BusWidth <> 1 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Bus Width Incorrect", "Read " & CStr(bSignal.BusWidth) & " instead of '0.25'"
  If bSignal.HasEdges <> 0 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Edges Incorrect", "Read " & CStr(bSignal.HasEdges) & " instead of '0.25'"

  ' event testing
  bSignal.AddEdge "Width/2"
  bSignal.AddEdge "Prop.Delay"
  bSignal.AddEdge 2.75
  bSignal.AddEdge 2.5
  bSignal.AddEdge "Width/2"
  bSignal.RemoveEvent 2.5
  bSignal.UpdateEvents
  bSignal.DrawEdges

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
  ' review
  If MsgBox(Title:="Base Signal Test", Buttons:=vbYesNo, Prompt:="Review Signal?") = vbYes Then Stop
  shp.ContainingPage.PageSheet.OpenSheetWindow
  bSignal.Delete
End Sub
