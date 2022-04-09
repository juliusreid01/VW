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

Public Sub NewSignal(x as Double, y as Double, w as Double, Optional ContainingPage as Page = Nothing)
  If ContainingPage Is Nothing Then Set ContainingPage = ActivePage
  Set Me.Shape = ContainingPage.DrawLine(x, y, x + w, y)
End Sub

Public Property Set Shape(s as Shape)
  Set bShape = new vw_base_shape_c
  Set bShape.Shape = s
  Set vsoShp = s
  Set wShape = bShape.Wrapper
End Property

Public Property Get Wrapper() as visio_shape_wrapper_c
  Set Wrapper = wShape
End Property

Public Property Get Base() as vw_base_shape_c
  Set Base = bShape
End Property

Public Sub Initialize(pType as SignalType)
  bShape.Initialize
  pSignalType = pType
  SetUserCells
  SetDataCells
End Sub

Public Sub Delete()
  bShape.Delete
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

Private Function GetNumeric(CellName as String, Optional AltReturn as Variant = 0) as Variant
  Dim Res as Variant
  Res = wShape.Result(CellName)
  If IsNumeric(Res) Then
    GetNumeric = Res
  Else
    GetNumeric = AltReturn
  End If
End Function

Public Property Get SignalType() as Integer
  SignalType = GetNumeric(S_TYPE, vw_types.SignalType.Void)
End Property

Public Property Get ChildOffset() as Double
  ChildOffset = GetNumeric(S_CHILDOFFSET)
End Property

Public Property Get ActiveWidth() as Double
  ActiveWidth = GetNumeric(S_ACTIVEWIDTH)
End Property

Public Property Get SkewWidth() as Double
  SkewWidth = GetNumeric(S_SKEWWIDTH)
End Property

Public Property Get Pulses() as Integer
  Pulses = GetNumeric(S_PULSES)
End Property

Public Property Get BusWidth() as Integer
  BusWidth = GetNumeric(S_BUSWIDTH, 1)
End Property

Public Property Get HasEdges() as Integer
  HasEdges = GetNumeric(S_EDGES)
End Property

Public Property Get Name() as String
  Name = wShape.ResultStr(S_NAME)
End Property

Public Property Get Clock() as String
  Clock = wShape.ResultStr(S_CLOCK)
End Property

Public Property Get Signal() as String
  Signal = wShape.ResultStr(S_SIGNAL)
End Property

Public Property Get ActiveLow() as Boolean
  ActiveLow = CBool(wShape.Result(S_ACTIVELOW))
End Property

Public Property Get Period() as Double
  Period = wShape.Result(S_PERIOD)
End Property

Public Property Get Skew() as Double
  Skew = wShape.Result(S_SKEW)
End Property

Public Property Get Delay() as Double
  Delay = wShape.Result(S_DELAY)
End Property

Public Property Get DutyCycle() as Double
  DutyCycle = wShape.Result(S_DUTYCYCLE)
End Property

Public Property Get SignalSkew() as Double
  SignalSkew = wShape.Result(S_SIGNALSKEW)
End Property

Public Property Get EventType() as String
  EventType = wShape.ResultStr(S_EVENTTYPE)
End Property

Public Property Let EventType(list as String)
  wShape.SetCell Position:="Format", CellName:=S_EVENTTYPE, v:=Chr(34) & list & Chr(34)
End Property

Public Property Let SelectEventType(idx as Integer)
  'Dim EventTriggerFormat as String
  'Dim EventPositionFormat as String
  wShape.SetCell CellName:=S_EVENTTYPE, _
                 v:="INDEX(" & idx & "," & wShape.GetCellName(S_EVENTTYPE) & ".Format)"
  Select Case Me.EventType
   Case S_LIST_EDGE
   Case S_LIST_NODE
    Setup_Node
   Case S_LIST_SPACER
   Case S_LIST_DRIVE_0
   Case S_LIST_DRIVE_1
   Case S_LIST_DRIVE_X
   Case S_LIST_DRIVE_Z
   Case ""
   Case Else
    Err.Raise vbObjectError + 512, "Base Signal: Unknown Event Type", "Actions not defined for Event Type: " & Me.EventType
  End Select
End Property

Public Property Get EventTrigger() as String
  EventTrigger = wShape.ResultStr(S_EVENTTRIGGER)
End Property

Public Property Let SelectEventTrigger(idx as Integer)
  wShape.SetCell CellName:=S_EVENTTRIGGER, _
                 v:="INDEX(" & idx & "," & wShape.GetCellName(S_EVENTTRIGGER) & ".Format)"
  Select Case Me.EventTrigger
   Case S_LIST_ABSOLUTE
   Case S_LIST_EDGE, S_LIST_POSEDGE, S_LIST_NEGEDGE
    EdgeList
    wShape.SetCell CellName:=S_EVENTPOSITION, v:="INDEX(0, " & wShape.GetCellName(S_EVENTPOSITION) & ".Format)"
   Case ""
   Case Else
    Err.Raise vbObjectError + 512, "Base Signal: Unknown Event Trigger", "Actions not defined for Event Trigger: " & Me.EventTrigger
  End Select
End Property

Public Property Let EventTrigger(list as String)
  wShape.SetCell Position:="Format", CellName:=S_EVENTTRIGGER, v:=Chr(34) & list & Chr(34)
End Property

Public Property Get EventPosition() as Variant
  EventPosition = wShape.Result(S_EVENTPOSITION)
End Property

Public Property Let EventPosition(list as Variant)
  wShape.SetCell Position:="Format", CellName:=S_EVENTPOSITION, v:=Chr(34) & list & Chr(34)
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

Public Sub AddEdge(xRelative as Variant)
  Dim Row as Integer

  Row = visRowLast

  ' add the row if there are none
  If vsoShp.RowCount(visSectionScratch) = 0 Then
    Row = vsoShp.AddRow(visSectionScratch, visRowLast, visTagDefault)
  ' otherwise select where to insert or modify an existing row
  Else
    wShape.SetCell S_TEST, xRelative
    For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
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
  vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Formula = vw_types.Edge
End Sub

Public Sub RemoveEvent(xRelative as Variant)
  Dim Row as Integer
  Row = -1
  wShape.SetCell S_TEST, xRelative

  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    If wShape.Result(S_TEST) = vsoShp.CellsSRC(visSectionScratch, i, visScratchX).Result("") Then Row = i
  Next

  If Row >= 0 Then vsoShp.DeleteRow visSectionScratch, Row
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
    Select Case eType And Not(vw_types.ROW_MASK)
     Case vw_types.Edge
      nEdges = nEdges + 1
      Select Case PreviousEvent
       Case Is < 0, vw_types.Gate0
        y = "Height*1"
       Case vw_types.GateX, vw_types.GateZ, vw_types.Gate1
        y = "Height*0"
       Case vw_types.Edge
        y = "Height-Scratch.Y" & CStr(i)
      End Select
     Case vw_types.Gate1
      y = "Height*1"
     Case vw_types.Gate0, vw_types.GateX
      y = "Height*0"
     Case vw_types.GateZ
      y = "Height*0.5"
    End Select
    vsoShp.CellsSRC(visSectionScratch, i, visScratchY).Formula = y
    PreviousEvent = vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultInt("", 0) And Not(vw_types.ROW_MASK)
  Next i

  Do While vsoShp.CellsSRC(visSectionScratch, visRowLast, visScratchX).Result("") >= wShape.Width
    If RowIsEdge(vsoShp.RowCount(visSectionScratch-1)) Then nEdges = nEdges - 1
    vsoShp.DeleteRow visSectionScratch, visRowLast
  Loop

  If nEdges >= 2 Then
    y = GenList(S_LIST_NONE, S_LIST_ALL, S_LIST_POSEDGE, S_LIST_NEGEDGE)
  ElseIf nEdges = 1 Then
    y = GenList(S_LIST_NONE, S_LIST_ALL)
  Else
    y = ""
  End If
  wShape.SetCell CellName:=S_LABELEDGES, v:=Chr(34) & y & Chr(34), Position:="Format"
  wShape.SetCell S_EDGES, nEdges
  'DrawEdges
End Sub

Public Sub DrawEdges()
  Dim GeometryRow as Integer
  Dim GeometryIdx as Integer

  GeometryRow = 2
  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    If vsoShp.RowCount(visSectionFirstComponent) - 1 <= GeometryRow Then _
      vsoShp.AddRow visSectionFirstComponent, GeometryRow, visTagLineTo
    wShape.SetPoint visSectionFirstComponent, GeometryRow, _
      "Scratch.X" & CStr(i+1), "Geometry1.Y" & CStr(GeometryRow - 1)
    '//TODO
    vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).Formula = _
      vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultInt("", 0) Or GeometryRow
    GeometryRow = GeometryRow + 1
    If vsoShp.RowCount(visSectionFirstComponent) - 1 <= GeometryRow Then _
      vsoShp.AddRow visSectionFirstComponent, GeometryRow, visTagLineTo
    wShape.SetPoint visSectionFirstComponent, GeometryRow, _
      wShape.GetCellName(S_SKEWWIDTH) & "+Scratch.X" & CStr(i+1), "Scratch.Y" & CStr(i+1)
    GeometryRow = GeometryRow + 1
  Next

  GeometryRow = vsoShp.RowCount(visSectionFirstComponent) - 1
  vsoShp.Cells("Geometry1.Y" & CStr(GeometryRow)).Formula = vsoShp.Cells("Geometry1.Y" & CStr(GeometryRow-1)).Formula
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
    Me.EventTrigger = ";" & GenList(S_LIST_EDGE, S_LIST_POSEDGE, S_LIST_NEGEDGE)
  ElseIf Posedge Then
    Me.EventTrigger = ";" & GenList(S_LIST_POSEDGE)
  ElseIf Negedge Then
    Me.EventTrigger = ";" & GenList(S_LIST_NEGEDGE)
  Else
    Me.SelectEventType = 0
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

  Me.EventPosition = list
End Sub

Private Function IsPosedge(Row as Integer) as Boolean
  Dim EventY as Double
  EventY = vsoShp.CellsSRC(visSectionScratch, Row, visScratchY).Result("")
  If Not RowIsEdge(Row) Then
    IsPosedge = False
  ElseIf (wShape.FlipY = 0 And EventY = wShape.Height) Or (wShape.FlipY = 1 And EventY = 0) Then
    IsPosedge = True
  Else
    IsPosedge = False
  End If
End Function

Private Function IsNegedge(Row as Integer) as Boolean
  If Not RowIsEdge(Row) Then
    IsNegedge = False
  Else
    IsNegedge = Not IsPosedge(Row)
  End If
End Function

Public Function RowIsEdge(Row as Integer) as Boolean
  Dim eType as Long
  RowIsEdge = False
  eType = vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).ResultInt("", 0)
  If (eType and vw_types.Edge) = vw_types.Edge Then RowIsEdge = True
End Function