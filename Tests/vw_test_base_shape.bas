Attribute VB_Name = "vw_test_base_shape"

Public Sub Test_BaseShape()
  Dim shp as Shape
  Dim bShp as vw_base_shape_c
  Dim wShp as visio_shape_wrapper_c

  Set shp = ActivePage.DrawLine(1, 10, 4, 11)
  Set bShp = New vw_base_shape_c
  Set bShp.Shape = shp
  Set wShp = bShp.Wrapper

  If wShp.EndY = wShp.BeginY Then Err.Raise vbObjectError + 2001, "Base Shape Test", "Failed to draw the shape correctly"

  bShp.Initialize
  If wShp.EndY <> wShp.BeginY Then Err.Raise vbObjectError + 2001, "Base Shape Test", "Failed to initialize the shape correctly"

  bShp.Delete
End Sub