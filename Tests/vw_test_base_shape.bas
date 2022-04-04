Attribute VB_Name = "vw_test_base_shape"

Private shp as Shape
Private bShp as vw_base_shape_c
Private wShp as visio_shape_wrapper_c
Private CheckDefaults as Boolean

Private Sub Create()
  Set shp = ActivePage.DrawLine(1, 10, 4, 11)
  Set bShp = New vw_base_shape_c
  Set bShp.Shape = shp
  Set wShp = bShp.Wrapper
  If wShp.EndY = wShp.BeginY Then Err.Raise vbObjectError + 2001, "Base Shape Test", "Failed to draw the shape correctly"
  bShp.Initialize
  If wShp.EndY <> wShp.BeginY Then Err.Raise vbObjectError + 2001, "Base Shape Test", "Failed to initialize the shape correctly"
End Sub

Private Sub Delete()
  bShp.Delete
End Sub

Public Sub Test_BaseShape(Optional s as vw_base_shape_c = Nothing)
  If Not s is Nothing Then
    Set bShp = s
    Set shp = bShp.Shape
    Set wShp = bShp.Wrapper
    CheckDefaults = False
    RunTest
  Else
    Create
    CheckDefaults = True
    RunTest
    Delete
  End If
End Sub

Private Sub RunTest()
  Dim ExpInvisible as Integer

  If shp.Text <> shp.Name Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Text", "Shape text does not match the name!"
  If shp.Cells("Height").Result("") <> shp.ContainingPage.PageSheet.Cells("BlockSizeY").Result("") Then _
    Err.Raise vbObjectError + 2003, "Base Shape Test: Height Error", "Failed to initialize shape height"

  If shp.OneD = True Then
    If wShp.EndY <> wShp.BeginY Then
      If MsgBox(Title:="Base Shape Test",  Buttons:=vbYesNo + vbQuestion, _
                Prompt:="Shape is not level. BeginY = " & CStr(wShp.BeginY) & ", EndY = " & CStr(wShp.EndY) _
                        & vbNewLine & "Continue?") = vbNo Then Stop
    End If
    If wShp.PinX <> wShp.BeginX Or wShp.PinY <> wShp.BeginY Or _
       wShp.LocPinX <> 0 Or wShp.LocPinY <> 0 Then
      If MsgBox(Title:="Base Shape Test",  Buttons:=vbYesNo + vbQuestion, _
                Prompt:="Shape orientation is not Bottom-Left." & vbNewLine & _
                         "BeginX = " & CStr(wShp.BeginX) & ", PinX = " & CStr(wShp.PinX) & vbNewLine & _
                         "BeginY = " & CStr(wShp.BeginY) & ", PinY = " & CStr(wShp.PinY) & vbNewLine & _
                         "LocPinX = " & CStr(wShp.LocPinX) & ", LocPinY = " & CStr(wShp.LocPinY) _
                         & vbNewLine & "Continue?") = vbNo Then Stop
    End If
  ElseIf wShp.LocPinX <> 0 Or wShp.LocPinY <> 0 Then
      If MsgBox(Title:="Base Shape Test",  Buttons:=vbYesNo + vbQuestion, _
                Prompt:="Shape orientation is not Bottom-Left." & vbNewLine & _
                         "PinX = " & CStr(wShp.PinX) & ", PinY = " & CStr(wShp.PinY) & vbNewLine & _
                         "LocPinX = " & CStr(wShp.LocPinX) & ", LocPinY = " & CStr(wShp.LocPinY) _
                         & vbNewLine & "Continue?") = vbNo Then Stop
  End If

  ' check the rows we expect are indeed created
  If shp.CellExists("User.Type", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.Type Cell was not found!"
  If shp.CellExists("User.ChildOffset", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.ChildOffset Cell was not found!"
  If shp.CellExists("User.BusWidth", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.BusWidth Cell was not found!"
  If shp.CellExists("User.SkewWidth", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.SkewWidth Cell was not found!"
  If shp.CellExists("User.Edges", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.Edges Cell was not found!"
  If shp.CellExists("User.ActiveWidth", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.ActiveWidth Cell was not found!"
  If shp.CellExists("User.Pulses", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.Pulses Cell was not found!"
  If shp.CellExists("User.Test", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.Test Cell was not found!"
  'If shp.CellExists("User.Parent", visExistsLocally) = False Then _
  '  Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected User.Parent Cell was not found!"

  If CheckDefaults = True Then
    For i = 0 to shp.RowCount(visSectionUser) - 1
      If shp.CellsSRC(visSectionUser, i, visUserValue).Result("") <> 0 Then _
        Err.Raise vbObjectError + 2002, "Base Shape Test: Default Error", "Expected Default: 0, Read: " & CStr(shp.CellsSRC(visSectionUser, i, visUserValue).Result(""))
    Next i
  End If

  If shp.CellExists("Prop.Name", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.Name Cell was not found!"
  If shp.CellExists("Prop.Clock", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.Clock Cell was not found!"
  If shp.CellExists("Prop.Signal", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.Signal Cell was not found!"
  If shp.CellExists("Prop.ActiveLow", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.ActiveLow Cell was not found!"
  If shp.CellExists("Prop.Period", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.Period Cell was not found!"
  If shp.CellExists("Prop.Skew", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.Skew Cell was not found!"
  If shp.CellExists("Prop.Delay", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.Delay Cell was not found!"
  If shp.CellExists("Prop.DutyCycle", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.DutyCycle Cell was not found!"
  If shp.CellExists("Prop.SignalSkew", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.SignalSkew Cell was not found!"
  If shp.CellExists("Prop.EventType", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.EventType Cell was not found!"
  If shp.CellExists("Prop.EventTrigger", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.EventTrigger Cell was not found!"
  If shp.CellExists("Prop.EventPosition", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.EventPosition Cell was not found!"
  If shp.CellExists("Prop.LabelEdges", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.LabelEdges Cell was not found!"
  If shp.CellExists("Prop.LabelSize", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.LabelSize Cell was not found!"
  If shp.CellExists("Prop.LabelFont", visExistsLocally) = False Then _
    Err.Raise vbObjectError + 2002, "Base Shape Test: Missing Cell", "Expected Prop.LabelFont Cell was not found!"

  If CheckDefaults = True Then
    For i = 0 to shp.RowCount(visSectionProp) - 1
      If shp.CellsSRC(visSectionProp, i, visCustPropsValue).Result("") <> 0 Then _
        Err.Raise vbObjectError + 2002, "Base Shape Test: Default Error", "Expected Default: 0, Read: " & CStr(shp.CellsSRC(visSectionProp, i, visCustPropsValue).Result(""))
    Next i
  End If

  For i = 0 to shp.RowCount(visSectionProp) - 1
    Select Case shp.CellsSRC(visSectionProp, i, visCustPropsInvis).RowName
      Case S_CLOCK, S_SIGNAL, S_LABELSIZE, S_LABELFONT
        ExpInvisible = 1
      Case S_LABELEDGES
        ExpInvisible = IIf(shp.Cells("Prop.LabelEdges").ResultStr("") <> "None", 1, 0)
      Case S_EVENTTRIGGER, S_EVENTPOSITION
        ExpInvisible = IIf(shp.Cells("Prop.EventType").ResultStr("") = "", 1, 0)
      Case Else
        ExpInvisible = 0
    End Select
    If ExpInvisible <> shp.CellsSRC(visSectionProp, i, visCustPropsInvis).Result("") Then _
      Err.Raise vbObjectError + 2002, "Base Shape Test: Visibility Error", _
        shp.CellsSRC(visSectionProp, i, visCustPropsInvis).RowName & " is not being displayed correctly!"
  Next i

End Sub