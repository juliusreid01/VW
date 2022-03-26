Attribute VB_Name = "Test"

Private Function Mismatch(Name as String, Exp as Double, Actual as Double) as String
  Mismatch = "Expected " & Name & " = " & CStr(Exp) & ", Actual =" & CStr(Actual)
End Function

Public Sub Test_1D()
  Dim shp as vw_shape_wrapper_c
  Set shp = new vw_shape_wrapper_c
  Set shp.vsoShape = ActivePage.DrawLine(1, 10, 4, 10)

  If shp.BeginX <> 1 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("BeginX", 1, shp.BeginX)
  If shp.EndX <> 4 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("EndX", 4, shp.EndX)
  If shp.BeginY <> 10 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("BeginY", 10, shp.BeginY)
  If shp.EndY <> 10 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("EndY", 10, shp.EndY)

  If shp.Width <> 3 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("Width", 3, shp.Width)
  If shp.Height <> 0 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("Height", 0, shp.Height)
  If shp.PinX <> (shp.BeginX+shp.EndX)/2 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("PinX", 2.5, shp.PinX)
  If shp.PinY <> (shp.BeginY+shp.EndY)/2 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("PinY", 0, shp.PinY)
  If shp.LocPinX <> shp.Width/2 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("LocPinX", 1.5, shp.LocPinX)
  If shp.LocPinY <> shp.Height/2 Then Err.Raise vbObjectError + 1001, "Test_1D", Mismatch("LocPinY", 0, shp.LocPinY)

  If shp.GetCellRefName("Width") <> shp.Name & "!Width" Then _
    Err.Raise vbObjectError + 1001, "Test_1D", "Incorrect Reference Returned!"

  shp.SetPoint visSectionFirstComponent, 2, 2, "Height/2"

  shp.Delete
End Sub

Public Sub Test_2D()
  Dim shp as vw_shape_wrapper_c
  Set shp = new vw_shape_wrapper_c
  Set shp.vsoShape = ActivePage.DrawRectangle(1, 10, 4, 10.5)

  If shp.BeginX <> 1 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("BeginX", 1, shp.BeginX)
  If shp.EndX <> 4 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("EndX", 4, shp.EndX)
  If shp.BeginY <> 10 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("BeginY", 10, shp.BeginY)
  If shp.EndY <> 10.5 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("EndY", 10.5, shp.EndY)

  If shp.Width <> 3 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("Width", 3, shp.Width)
  If shp.Height <> 0.5 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("Height", 0.5, shp.Height)
  If shp.PinX <> (shp.BeginX+shp.EndX)/2 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("PinX", 2.5, shp.PinX)
  If shp.PinY <> (shp.BeginY+shp.EndY)/2 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("PinY", 0, shp.PinY)
  If shp.LocPinX <> shp.Width/2 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("LocPinX", 1.5, shp.LocPinX)
  If shp.LocPinY <> shp.Height/2 Then Err.Raise vbObjectError + 2001, "Test_2D", Mismatch("LocPinY", 0, shp.LocPinY)

  If shp.GetCellRefName("Width") <> shp.Name & "!Width" Then _
    Err.Raise vbObjectError + 2001, "Test_2D", "Incorrect Reference Returned!"

  shp.SetPoint visSectionFirstComponent, 2, 2, "Height/2"

  shp.Delete
End Sub
