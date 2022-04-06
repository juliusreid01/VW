Attribute VB_Name = "vw_test_clock_c"

Public Sub Test_Clock()
  Dim clk as vw_clock_c
  Set clk = new vw_clock_c
  clk.NewSignal 1, 9.5, 3
  clk.Shape.OpenSheetWindow

  If clk.Wrapper is Nothing Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Wrapper is null", "Missing Wrapper!"
  If clk.Base is Nothing Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Base is null", "Missing Base!"
  If clk.Base.SignalType <> SignalType.Clock Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Signal Type Incorrect", "Read " & CStr(bSignal.SignalType) & " instead of 'Signal'!"

  If clk.Base.Name <> clk.Shape.Text Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Signal Name Incorrect", "Shape Name does not match text!"
  If clk.Base.ActiveLow <> False Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Active Low Default: False, Read: " & CStr(clk.Base.ActiveLow)
  If clk.Base.Period <> 0.25 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Default Period: 0.25, Read: " & CStr(clk.Base.Period)
  If clk.Base.Skew <> 0.10 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Default Skew: 0.10, Read: " & CStr(clk.Base.Skew)

  ' review
  If MsgBox(Title:="Clock Test", Buttons:=vbYesNo, Prompt:="Review Signal?") = vbYes Then Stop
  clk.Delete
End Sub