VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_child_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' class to handle nodes

Private vsoShp as Shape
Private wShape as visio_shape_wrapper_c
Private bShape as vw_base_shape_c
Private Row as Integer
Private pType as Integer
Private MyShape as visio_shape_wrapper_c
Private NodeID as Integer
Private LabelID as Integer
Private iExist as Boolean
Private TypeCell as String
Private ParentCell as String

Public Sub Class_Initialize()
  iExist = False
End Sub

Private Sub Class_Terminate()
  Set bShape = Nothing
  Set wShape = Nothing
  Set vsoShp = Nothing
  Set MyShape = Nothing
End Sub

Public Sub SetParent(s as vw_base_shape_c, ChildType as Integer, ScratchRow as Integer)
  Set bShape = s
  Set wShape = s.Wrapper
  Set vsoShp = s.Shape
  Row = ScratchRow
  pType = ChildType
  '//TODO. the initial node value should be configurable
  NodeID = 97
  '//TODO. the initial label value should be configurable
  LabelID = 1
  TypeCell = wShape.GetCellName(S_TYPE, False)
  ParentCell = wShape.GetCellName(S_PARENT, False)
End Sub

Private Property Get NodeXF() as String
  Dim GeometryRow as Integer
  GeometryRow = vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE) And vw_types.ROW_MASK

  NodeXF = vsoShp.Name & "!PinX + " & vsoShp.Name & "!Geometry1.X" & CStr(GeometryRow) & "+((" & _
           (vsoShp.Name & "!Geometry1.X" & CStr(GeometryRow + 1)) & "-" & (vsoShp.Name & "!Geometry1.X" & CStr(GeometryRow)) & _
           ")*0.5)"
End Property

Private Property Get NodeX() as Double
  wShape.SetCell S_TEST, NodeXF
  NodeX = wShape.Result(S_TEST)
End Property

Private Property Get NodeYF() as String
  NodeYF = vsoShp.Name & "!PinY + (" & vsoShp.Name & "!Height*0.5)"
End Property

Private Property Get NodeY() as Double
  wShape.SetCell S_TEST, NodeYF
  NodeY = wShape.Result(S_TEST)
End Property

Private Property Get LabelXF() as String
  LabelXF = NodeXF
End Property

Private Property Get LabelX() as Double
  wShape.SetCell S_TEST, LabelXF
  LabelX = wShape.Result(S_TEST)
End Property

Private Property Get LabelYF() as String
  LabelYF = vsoShp.Name & "!PinY + " & vsoShp.Name & "!Height + Height"
End Property

Private Property Get LabelY() as Double
  wShape.SetCell S_TEST, LabelYF
  LabelY = wShape.Result(S_TEST)
End Property


Private Property Get MyX() as Double
  Select Case pType
   Case vw_types.Node: MyX = NodeX
   Case vw_types.Edge: MyX = LabelX
  End Select
End Property

Private Property Get MyY() as Double
  Select Case pType
   Case vw_types.Node: MyY = NodeY
   Case vw_types.Edge: MyY = LabelY
  End Select
End Property

Private Function Exists(Container as Variant) as Boolean
  If iExist = True Then Exit Function

  For Each s in Container.Shapes
    If s.Shapes.Count > 0 Then
      Exists s
    ElseIf s.CellExists(TypeCell, visExistsLocally) <> False And _
           s.CellExists(ParentCell, visExistsLocally) <> False Then
      If CBool(((s.Cells(TypeCell).Result("") And pType) <> 0) And _
                 s.Cells(ParentCell).ResultStr("") = vsoShp.Name And _
                 s.Cells("PinX").Result("") = MyX And s.Cells("PinY").Result("") = MyY) Then
        iExist = True
        Exit For
      End If
    End If
  Next
  Exists = iExist
End Function

Private Sub UpdateNodeID(Container as Variant)
  For Each s in Container.Shapes
    If s.Shapes.Count > 0 Then
      UpdateNodeID s
    ElseIf s.CellExists(TypeCell, visExistsLocally) <> False And s.Text <> "" Then
      If s.Cells(TypeCell) = vw_types.Node Then NodeID = Asc(Left(s.Text, 1)) + 1
    End If
  Next
End Sub

Private Sub UpdateLabelID(Container as Variant)
  For Each s in Container.Shapes
    If s.Shapes.Count > 0 Then
      UpdateLabelID s
    ElseIf s.CellExists(TypeCell, visExistsLocally) <> False And IsNumeric(s.Text) Then
      If s.Cells(TypeCell) = vw_types.Edge Then LabelID = CInt(s.Text) + 1
    End If
  Next
End Sub

Public Sub Draw()
  If Exists(vsoShp.ContainingPage) Then Exit Sub
  Set MyShape = new visio_shape_wrapper_c
  Select Case pType
   Case vw_types.Node: DrawNode
   Case vw_types.Edge: DrawLabel
  End Select
  ' reselect the shape
  For Each w in Application.Windows
    If w.Page = vsoShp.ContainingPage And w.Type = visTypePage Then
      w.Select vsoShp, visDeselectAll Or visSelect
    End If
  Next
End Sub

Private Sub DrawCommon(ByRef shp as Shape)
  Set shp = vsoShp.ContainingPage.DrawRectangle(wShape.PinX, wShape.PinY, wShape.PinX, wShape.PinY)

  Set MyShape.vsoShape = shp
  MyShape.Width = "Guard(Height)"
  shp.Cells("LeftMargin").Formula = "0"
  shp.Cells("RightMargin").Formula = "0"
  shp.Cells("BottomMargin").Formula = "0"
  shp.Cells("TopMargin").Formula = "0"
  MyShape.AddUserCell "Type"
  MyShape.SetCell CellName:="Type", v:=pType
  MyShape.AddUserCell "Parent"
  MyShape.SetCell CellName:="Parent", v:=Chr(34) & vsoShp.Name & Chr(34)
  MyShape.AddUserCell "ScratchID"
  MyShape.SetCell CellName:="ScratchID", v:=Row
  shp.BringToFront
End Sub

Private Sub DrawNode()
  Dim shp as Shape
  DrawCommon shp
  '//TODO. this should be coming from vw_cfg as NODE_HEIGHT_MULTIPLIER
  MyShape.Height = "Guard(" & vsoShp.Name & ("!Height") & "*0.5)"
  MyShape.PinX = "Guard(" & NodeXF & ")"
  MyShape.PinY = "Guard(" & NodeYF & ")"
  shp.Cells("Rounding").Formula = "Width*0.5"
  '//TODO. should nodes have their own setting
  shp.Cells("Char.Size").Formula = "Guard(" & vsoShp.Name & "!" & wShape.GetCellName(S_LABELFONT) & ")"
  ' update the row data
  vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Formula = _
    vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).ResultInt("", 0) Or vw_types.Node
  UpdateNodeID vsoShp.ContainingPage
  shp.Text = Chr(NodeID)
End Sub

Private Sub DrawLabel()
  Dim shp as Shape
  DrawCommon shp
  MyShape.Height = "Guard(" & vsoShp.Name & "!" & wShape.GetCellName(S_LABELSIZE) & ")"
  MyShape.Width = "Guard(Height)"
  MyShape.PinX = "Guard(" & LabelXF & ")"
  MyShape.PinY = "Guard(" & LabelYF & ")"
  '//TODO this should be configurable
  shp.Cells("Rounding").Formula = "Width*0.5"
  shp.Cells("Char.Size").Formula = "Guard(" & vsoShp.Name & "!" & wShape.GetCellName(S_LABELFONT) & ")"
  UpdateLabelID vsoShp.ContainingPage
  shp.Text = CStr(LabelID)
End Sub
