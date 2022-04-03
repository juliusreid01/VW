Attribute VB_Name = "vw_test_base_signal"

Public Sub Test_BaseSignal()
  Dim shp as Shape
  Dim bSignal as vw_base_signal_c

  Set shp = ActivePage.DrawLine(1, 10, 4, 10)

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
  If bSignal.Pulses <> 6 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Pulses Incorrect", "Read " & CStr(bSignal.Pulses) & " instead of '6'"
  If bSignal.BusWidth <> 1 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Bus Width Incorrect", "Read " & CStr(bSignal.BusWidth) & " instead of '0.25'"
  If bSignal.HasEdges <> 0 Then _
    Err.Raise vbObjectError + 2003, "Base Signal Test: Signal Edges Incorrect", "Read " & CStr(bSignal.HasEdges) & " instead of '0.25'"

  bSignal.Delete
End Sub
