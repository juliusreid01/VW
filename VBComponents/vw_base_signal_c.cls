VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_base_signal_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' handle to the Visio Shape
Private vsoShp as Shape
' handle to the Visio Shape Wrapper
Private wShape as visio_shape_wrapper_c
' handle to the base_shape
Private bShape as vw_base_shape_c
' type of signal
Private pSignalType as SignalType

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
  Set vsoShp = Nothing
  Set wShape = Nothing
  set bShape = Nothing
End Sub

Public Property Set Shape(s as Shape)
  Set bShape = new vw_base_shape_c
  Set bShape.Shape = s
  Set vsoShp = s
  Set wShape = bShape.Wrapper
  pSignalType = wShape.Result(CellName:=S_TYPE, RaiseError:=False)
End Property

Public Property Get Shape() as Shape
  Set Shape = vsoShp
End Property

Public Property Get Wrapper() as visio_shape_wrapper_c
  Set Wrapper = wShape
End Property

Public Property Get Base() as vw_base_shape_c
  Set Base = bShape
End Property

Public Sub Delete()
  bShape.Delete
End Sub

Public Sub Initialize(pType as SignalType)
  bShape.Initialize
  pSignalType = pType
  SetUserCells
  SetDataCells
End Sub

Private Sub SetUserCells()
  With wShape
    .SetCell S_TYPE, Chr(34) & pSignalType & Chr(34)
    .SetCell S_CHILDOFFSET, "BlockSizeY"
    .SetCell S_ACTIVEWIDTH, VW_0 & "+(" & .GetCellName(S_PERIOD) & "*" & .GetCellName(S_DUTYCYCLE) & ")"
    .SetCell S_SKEWWIDTH, VW_0 & "+(" & .GetCellName(S_SKEW) & "*" & .GetCellName(S_PERIOD) & "*0.5)"
    .SetCell S_PULSES, "(Width-Geometry1.X2)/" & .GetCellName(S_PERIOD)
    .SetCell S_BUSWIDTH, 1
    .SetCell S_EDGES, False
  End With
End Sub

Private Sub SetDataCells()
  With wShape
    .SetCell S_ACTIVELOW, False
    .SetCell S_PERIOD, "BlockSizeX*2"
    .SetCell S_SKEW, "10 %"
    .SetCell S_DELAY, "BlockSizeX*0.5"
    .SetCell S_DUTYCYCLE, "50 %"
    .SetCell S_SIGNALSKEW, .GetCellName(S_SKEWWIDTH)
    .SetCell S_EVENTTYPE, "Index(0," & .GetCellName(S_EVENTTYPE) & ".Format)"
    .SetCell S_LABELEDGES, "Index(0," & .GetCellName(S_LABELEDGES) & ".Format)"
    .SetCell S_LABELSIZE, "BlockSizeX*0.667"
    .SetCell S_LABELFONT, "8 pt"

    .SetCell Position:="Format", CellName:=S_LABELEDGES, v:=Chr(34) & S_LIST_NONE & Chr(34)
    .SetCell Position:="Format", CellName:=S_EVENTTRIGGER, v:=Chr(34) & S_LIST_ABSOLUTE & Chr(34)
  End With
End Sub

Public Property Get SignalType() as SignalType
  SignalType = pSignalType
End Property

'//TODO. do we REALLY need these property getters
Public Property Get Name() as String
  Name = wShape.ResultStr(S_NAME)
End Property

Public Property Get Clock() as String
  Clock = wShape.ResultStr(S_CLOCK)
End Property

Public Property Get Signal() as String
  Signal = wShape.ResultStr(S_SIGNAL)
End Property

Public Property Get EventType() as String
  EventType = wShape.ResultStr(S_EVENTTYPE)
End Property

Public Property Get EventTrigger() as String
  EventTrigger = wShape.ResultStr(S_EVENTTRIGGER)
End Property

Public Property Get EventPosition() as Variant
  EventPosition = wShape.Result(S_EVENTPOSITION)
End Property

Public Property Get LabelEdges() as String
  LabelEdges = wShape.ResultStr(S_LABELEDGES)
End Property

Public Property Get LabelSize() as Double
  LabelSize = wShape.Result(S_LABELSIZE)
End Property

Public Property Get LabelFont() as Integer
  LabelFont = wShape.Result(S_LABELFONT)
End Property

Public Sub AddEvent(xRelative as Variant, Optional eType as EventType = vw_types.Edge)
  Dim Row as Integer

  Row = visRowLast
  ' add the row if there are none
  If vsoShp.RowCount(visSectionScratch) = 0 Then
    Row = vsoShp.AddRow(visSectionScratch, visRowLast, visTagDefault)
  ' otherwise select where to insert or modify an existing row
  Else
    wShape.SetCell S_TEST, xRelative
    For i = vsoShp.RowCount(visSectionScratch) - 1 to 0 Step -1
      If wShape.Result(S_TEST) < vsoShp.CellsSRC(visSectionScratch, i, visScratchX).Result("") Then
        Row = i
      ElseIf wShape.Result(S_TEST) = vsoShp.CellsSRC(visSectionScratch, i, visScratchX).Result("") Then
        Row = i
        Exit For
      End If
    Next i
    If wShape.Result(S_TEST) <> vsoShp.CellsSRC(visSectionScratch, Row, visScratchX).Result("") Then _
      Row = vsoShp.AddRow(visSectionScratch, Row, visTagDefault)
  End If
  vsoShp.CellsSRC(visSectionScratch, Row, visScratchX).Formula = xRelative
  '//TODO if we decide to store additional information in the decimal portion this formula will fail
  vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Formula = eType Or _
    vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Result("")
End Sub

Public Sub RemoveEvent(xRelative as Variant, Optional eType as EventType = vw_types.Edge)
  Dim Row as Integer
  Row = -1
  wShape.SetCell S_TEST, xRelative

  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    If wShape.Result(S_TEST) = vsoShp.CellsSRC(visSectionScratch, i, visScratchX).Result("") Then
      Row = i
      Exit For
    End If
  Next

  If Row >= 0 Then
    '//TODO if we decide to store additional information in the decimal portion this formula will fail
    vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Formula = Not(eType) And _
      vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Result("")
    '//TODO some eTypes have items that must be removed
    If (vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Result("") And Not(ROW_MASK)) = 0 Then _
      vsoShp.DeleteRow visSectionScratch, Row
  End If
End Sub

Public Sub UpdateEvents()
  Dim PreviousEvent as Integer
  Dim y as Variant
  Dim nEdges as Integer
  Dim eType as Long
  ' update the y formula for each row
  ' Previous Type | Current Type | y
  ' Edge->Edge = Height-(PreviousRow)
  ' *->Gate1 = Height*1
  ' *->Gate0 = Height*0
  ' *->GateX = Height*0
  ' *->GateZ = Height*0.5
  nEdges = 0
  PreviousEvent = -1
  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    eType = vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultInt("", 0)
    '//TODO. comes up as edge and node
    If (eType And vw_types.Edge) <> 0 Then
      nEdges = nEdges + 1
      If (PreviousEvent < 0) Or ((PreviousEvent And vw_types.Gate0) <> 0) Then
        y = "Height*1"
      ElseIf (PreviousEvent And vw_types.Gate) <> 0 Then
        y = "Height*0"
      ElseIf (PreviousEvent And vw_types.Edge) <> 0 Then
        y = "Height-Scratch.Y" & CStr(i)
      End If
    ElseIf (eType And vw_types.Gate1) <> 0 Then
      y = "Height*1"
    ElseIf (eType And vw_types.GateZ) <> 0 Then
      y = "Height*0.5"
    Else 'vw_types.Gate0, vw_types.GateX
      y = "Height*0"
    End If
    vsoShp.CellsSRC(visSectionScratch, i, visScratchY).Formula = y
    PreviousEvent = vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultInt("", 0) And Not(vw_types.ROW_MASK)
  Next i

  ' remove any rows >= width
  Do While vsoShp.CellsSRC(visSectionScratch, visRowLast, visScratchX).Result("") >= wShape.Width
    If IsEdge(vsoShp.RowCount(visSectionScratch-1)) Then nEdges = nEdges - 1
    vsoShp.DeleteRow visSectionScratch, visRowLast
  Loop

  ' update which labels can be shown
  If nEdges >= 2 Then
    y = GenList(S_LIST_NONE, S_LIST_ALL, S_LIST_POSEDGE, S_LIST_NEGEDGE)
  ElseIf nEdges = 1 Then
    y = GenList(S_LIST_NONE, S_LIST_ALL)
  Else
    y = ""
  End If
  wShape.SetCell CellName:=S_LABELEDGES, v:=Chr(34) & y & Chr(34), Position:="Format"
  wShape.SetCell S_EDGES, nEdges
  'DrawEvents
End Sub

Public Sub DrawEvents()
  Dim eType as Integer
  Dim GeometryRow as Integer

  GeometryRow = 2
  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    eType = vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultInt("", 0) And Not(ROW_MASK)

    If ((eType And Edge) = Edge) Or ((eType And Gate) = Gate) Then DrawTransition CInt(i), GeometryRow
    If (eType And Node) = Node Then DrawNode CInt(i)
    'If (eType And Gap) = Gap Then
  Next

  ' remove additional geometry rows due to width changing
  Do While vsoShp.RowCount(visSectionFirstComponent) - 1 > GeometryRow
    vsoShp.DeleteRow visSectionFirstComponent, GeometryRow
  Loop

  ' correct last row
  wShape.SetPoint visSectionFirstComponent, GeometryRow, "Width", vsoShp.Cells("Geometry1.Y" & CStr(GeometryRow-1)).Formula

  Set Collection = Nothing
End Sub

Private Sub DrawTransition(i as Integer, ByRef GeometryRow as Integer)
  Dim GeometryIdx as Integer
  Dim xFormula as Variant
  Dim eType as Integer
  Dim EventY as Double

  eType = vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultInt("", 0) And Not(ROW_MASK)
  EventY = vsoShp.CellsSRC(visSectionScratch, i, visScratchY).Result("")

  ' add a row if needed
  If vsoShp.RowCount(visSectionFirstComponent) - 1 <= GeometryRow Then vsoShp.AddRow visSectionFirstComponent, GeometryRow, visTagLineTo
  ' makes the line to the event
  wShape.SetPoint visSectionFirstComponent, GeometryRow, "Scratch.X" & CStr(i+1), "Geometry1.Y" & CStr(GeometryRow - 1)
  ' Or with the GeometryRow
  vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).Formula = eType Or GeometryRow
  GeometryRow = GeometryRow + 1
  ' determine If we should apply skew
  If (eType And Edge) = Edge Then
    xFormula = wShape.GetCellName(S_SKEWWIDTH) & "+Scratch.X" & CStr(i+1)
  ElseIf (eType And Gate) = Gate Then
    xFormula = "Scratch.X" & CStr(i+1)
  End If
  ' make the event
  If vsoShp.RowCount(visSectionFirstComponent) - 1 <= GeometryRow Then vsoShp.AddRow visSectionFirstComponent, GeometryRow, visTagLineTo
  wShape.SetPoint visSectionFirstComponent, GeometryRow, xFormula, "Scratch.Y" & CStr(i+1)
  GeometryRow = GeometryRow + 1
  ' draw additional geometry
End Sub

Private Sub DrawNode(ScratchRow as Integer)
  Dim Node as vw_child_c
  Set Node = new vw_child_c
  Node.SetParent bShape, vw_types.Node, ScratchRow
  Node.Draw
End Sub

Private Sub Setup_Node()
  Dim Row as Integer
  Dim Posedge as Boolean
  Dim Negedge as Boolean

  Row = 0
  Posedge = False
  Negedge = False

  Do While Row < vsoShp.RowCount(visSectionScratch) And (Posedge = False Or Negedge = False)
    Posedge = Posedge Or IsPosedge(Row)
    Negedge = Negedge Or IsNegedge(Row)
    Row = Row + 1
  Loop

  wShape.SetCell Position:="Type", CellName:=S_EVENTPOSITION, v:=visPropTypeListFix
  EdgeList
  wShape.SetCell CellName:=S_EVENTPOSITION, v:="INDEX(0, " & wShape.GetCellName(S_EVENTPOSITION) & ".Format)"

  If Posedge And Negedge Then
    wShape.SetCellFormat S_EVENTTRIGGER, ";" & GenList(S_LIST_EDGE, S_LIST_POSEDGE, S_LIST_NEGEDGE)
  ElseIf Posedge Then
    wShape.SetCellFormat S_EVENTTRIGGER, ";" & GenList(S_LIST_POSEDGE)
  ElseIf Negedge Then
    wShape.SetCellFormat S_EVENTTRIGGER, ";" & GenList(S_LIST_NEGEDGE)
  Else
    wShape.SetCell CellName:=S_EVENTTYPE, _
                   v:="INDEX(0," & wShape.GetCellName(S_EVENTTYPE) & ".Format)"
  End If
End Sub

Private Sub EdgeList()
  Dim incr as Boolean
  Dim cnt as Integer
  Dim list as String

  cnt = 0
  list = ""

  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    incr = False
    Select Case Me.EventTrigger
     Case S_LIST_EDGE
      If IsPosedge(CInt(i)) Or IsNegedge(CInt(i)) Then incr = True
     Case S_LIST_POSEDGE
      If IsPosedge(CInt(i)) Then incr = True
     Case S_LIST_NEGEDGE
      If IsNegedge(CInt(i)) Then incr = True
    End Select
    If incr = True Then
      cnt = cnt + 1
      list = list & ";" & CStr(cnt)
    End If
  Next i

  wShape.SetCellFormat S_EVENTPOSITION, list
End Sub

Private Function IsPosedge(Row as Integer) as Boolean
  Dim EventY as Double
  EventY = vsoShp.CellsSRC(visSectionScratch, Row, visScratchY).Result("")
  If Not IsEdge(Row) Then
    IsPosedge = False
  ElseIf (wShape.FlipY = 0 And EventY = wShape.Height) Or (wShape.FlipY = 1 And EventY = 0) Then
    IsPosedge = True
  Else
    IsPosedge = False
  End If
End Function

Private Function IsNegedge(Row as Integer) as Boolean
  If Not IsEdge(Row) Then
    IsNegedge = False
  Else
    IsNegedge = Not IsPosedge(Row)
  End If
End Function

Public Function IsEdge(Row as Integer) as Boolean
  Dim eType as Long
  IsEdge = False
  eType = vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).ResultInt("", 0)
  If (eType and vw_types.Edge) = vw_types.Edge Then IsEdge = True
End Function

Public Function GetRowIdx(Row as Integer) as Integer
  Dim eType as Long
  eType = vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).ResultInt("", 0)
  GetRowIdx = eType And ROW_MASK
End Function