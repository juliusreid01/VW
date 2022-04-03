Attribute VB_Name = "vw_test_base_shape"

Private shp as Shape
Private bShp as vw_base_shape_c
Private wShp as visio_shape_wrapper_c

Private Sub Create()
  Set shp = ActivePage.DrawLine(1, 10, 4, 11)
  Set bShp = New vw_base_shape_c
  Set bShp.Shape = shp
  Set wShp = bShp.Wrapper
  If wShp.EndY = wShp.BeginY Then Err.Raise vbObjectError + 2001, "Base Shape Test", "Failed to draw the shape correctly"
  bShp.Initialize
End Sub

Private Sub Delete()
  bShp.Delete
End Sub

Public Sub Test_BaseShape(Optional s as vw_base_shape_c = Nothing)
  If Not s is Nothing Then
    Set bShp = s
    Set shp = bShp.Shape
    Set wShp = bShp.Wrapper
    RunTest
  Else
    Create
    RunTest
    Delete
  End If
End Sub

Private Sub RunTest()
  If wShp.EndY <> wShp.BeginY Then Err.Raise vbObjectError + 2001, "Base Shape Test", "Failed to initialize the shape correctly"

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
End Sub