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
    Err.Raise vbObjectError + 2004, "Clock Test: Signal Type Incorrect", "Read " & CStr(clk.Base.SignalType) & " instead of 'Signal'!"

  If clk.Base.Name <> clk.Shape.Text Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Signal Name Incorrect", "Shape Name does not match text!"
  If clk.Base.ActiveLow <> False Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Active Low Default: False, Read: " & CStr(clk.Base.ActiveLow)
  If clk.Base.Period <> 0.5 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Default Period: 0.5, Read: " & CStr(clk.Base.Period)
  If clk.Base.Skew <> 0.10 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Default Skew: 0.10, Read: " & CStr(clk.Base.Skew)
  If clk.Base.Delay <> 0.125 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Default Delay: 0.125, Read: " & CStr(clk.Base.Delay)
  If clk.Base.DutyCycle <> 0.5 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Default Duty Cycle: 0.5, Read: " & CStr(clk.Base.DutyCycle)
  If clk.Base.SignalSkew <> 0.025 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Default", "Default Signal Skew: 0.025, Read: " & CStr(clk.Base.SignalSkew)

  If clk.Shape.Cells("Prop.EventType.Format").Formula <> """;Node;Spacer;Drive0;Drive1;DriveX;DriveZ;Delete""" Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect List", "Expected List: ';Node;Spacer;Drive0;Drive1;DriveX;DriveZ;Delete', Read: " & clk.Shape.Cells("Prop.EventType.Format").Formula
  clk.Base.SelectEventType = 1
  If clk.Base.EventType <> "Node" Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Selection Failure", "Failed to select Node Event Type"
  If clk.Shape.Cells("Prop.EventTrigger.Invisible") <> 0 Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect Visibilty", "Event Trigger should be visible when set"
  If clk.Shape.Cells("Prop.EventTrigger.Format").Formula <> """;Posedge;Negedge""" Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect List", "Expected List: ';Posedge;Negedge'"

  If clk.Shape.Cells("Prop.LabelEdges.Format").Formula <> """None;All;Posedge;Negedge""" Then _
    Err.Raise vbObjectError + 2004, "Clock Test: Incorrect List", "Expected List: 'None;All;Posedge;Negedge', Read: " & clk.Shape.Cells("Prop.LabelEdges.Format").Formula

  ' review
  If MsgBox(Title:="Clock Test", Buttons:=vbYesNo, Prompt:="Review Signal?") = vbYes Then Stop
  clk.Shape.ContainingPage.PageSheet.OpenSheetWindow
  clk.Delete
End Sub