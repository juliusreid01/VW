VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_signal_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' handle to the shape being modified
Private vsoShp as Shape
'shape wrapper
Private shp as vw_shape_wrapper_c
' the type of signal
Private sType as SignalType

Private Sub Class_Initialize()
  vw_cfg.Configure
  Set shp = new vw_shape_wrapper_c
End Sub

Private Sub Class_Terminate()
  Set vsoShp = Nothing
  Set shp = Nothing
End Sub

Public Property Get vsoShape as Shape
  set vsoShape = vsoShp
End Property

' Create a new signal of the given type on the input page
Public Sub NewSignal(p as Page, t as SignalType, x1 as double, y1 as double)
  Register p.DrawLine(x1, y1, x1 + VW_WIDTH, y1), t
  Init
End Sub

' Convert a shape to the given type on the input page
Public Sub ConvertSignal(s as Shape, t as SignalType)
  Register s, t
  RemoveSections
  Init
End Sub

' Take a shape but don't do anything
Public Sub Register(s as Shape, t as SignalType)
  Set vsoShp = s
  Set shp.vsoShape = s
  sType = t
End Sub

' Initialize the shape
Private Sub Init()
  AddDataCells
  ResetData
  SetDataPrompts
  SetDataLabels
  SetDataTypes
  SetDataFormats
  SetDataHidden

  AddUserCells
  SetUserCells

  ResetVisioData
  ResetGeometry
End Sub

' Remove sections from the shapesheet
Private Sub RemoveSections()
  vsoShp.DeleteSection visSectionProp
  vsoShp.DeleteSection visSectionUser
  vsoShp.DeleteSection visSectionControls
  vsoShp.DeleteSection visSectionActions
  vsoShp.DeleteSection visSectionConnectionPts
End Sub

' Add the User-defined Cells
Private Sub AddUserCells()
  shp.AddUserCell S_TYPE
  shp.AddUserCell S_CHILDOFFSET
  if sType = SignalType.Bus or sType = SignalType.Bit then
    shp.AddUserCell S_BUSWIDTH
    shp.AddUserCell S_SKEWWIDTH
    shp.AddUserCell S_EDGES
  elseif sType = SignalType.Clock Then
    shp.AddUserCell S_ACTIVEWIDTH
    shp.AddUserCell S_SKEWWIDTH
    shp.AddUserCell S_PULSES
  end if
  shp.AddUserCell S_TEST, """Spare cell to test formulae and values"""
End Sub

' Set the User-defined Cells
Private Sub SetUserCells()
  shp.SetUserCell(S_TYPE, chr(34) & VW_TYPE_STR(sType) & chr(34))
  shp.SetUserCell(S_CHILDOFFSET, "BlockSizeY")
  if sType = SignalType.Clock then
    shp.SetUserCell(S_ACTIVEWIDTH, VW_0 & "+ (Prop.Period * Prop.DutyCycle)")
    shp.SetUserCell(S_SKEWWIDTH, VW_0 & "+ (Prop.Skew * Prop.Period * 0.5)")
    shp.SetUserCell(S_PULSES, "CEILING((Width-Geometry1.X1)/Prop.Period)")
  elseif sType = SignalType.Bit then
    shp.SetUserCell(S_BUSWIDTH, 1)
    shp.SetUserCell(S_EDGES, False)
  elseif sType = SignalType.Bus then
    shp.SetUserCell(S_BUSWIDTH, 8)
    shp.SetUserCell(S_EDGES, False)
  end if
End Sub

' Add the Shape Data Cells
Private Sub AddDataCells()
  vsoShp.AddNamedRow visSectionProp, S_NAME, visTagDefault
  vsoShp.Cells(Prop(S_NAME) & ".Verify").Formula = True
  vsoShp.AddNamedRow visSectionProp, S_CLOCK, visTagDefault
  vsoShp.AddNamedRow visSectionProp, S_SIGNAL, visTagDefault
  if sType = SignalType.Clock or sType = SignalType.Bit then
    vsoShp.AddNamedRow visSectionProp, S_ACTIVELOW, visTagDefault
  end if
  if sType = SignalType.Clock then
    vsoShp.AddNamedRow visSectionProp, S_PERIOD, visTagDefault
    vsoShp.AddNamedRow visSectionProp, S_SKEW, visTagDefault
  end if
  vsoShp.AddNamedRow visSectionProp, S_DELAY, visTagDefault
  if sType = SignalType.Clock then
    vsoShp.AddNamedRow visSectionProp, S_DUTYCYCLE, visTagDefault
    vsoShp.AddNamedRow visSectionProp, S_SIGNALSKEW, visTagDefault
  end if
  vsoShp.AddNamedRow visSectionProp, S_EVENTTYPE, visTagDefault
  vsoShp.AddNamedRow visSectionProp, S_EVENTTRIGGER, visTagDefault
  vsoShp.AddNamedRow visSectionProp, S_EVENTPOSITION, visTagDefault
  vsoShp.AddNamedRow visSectionProp, S_LABELEDGES, visTagDefault
  vsoShp.AddNamedRow visSectionProp, S_LABELSIZE, visTagDefault
  vsoShp.AddNamedRow visSectionProp, S_LABELFONT, visTagDefault
End Sub

' Set the Shape Data Cell Defaults
Public Sub ResetData(Optional s as Shape = Nothing)
  if Not s Is Nothing then set vsoShp = s

  if sType = SignalType.Clock or sType = SignalType.Bit then
    vsoShp.Cells(Prop(S_ACTIVELOW)).Formula = False
  end if
  if sType = SignalType.Clock then
    vsoShp.Cells(Prop(S_PERIOD)).Formula = "BlockSizeX*2"
    vsoShp.Cells(Prop(S_SKEW)).Formula = "10 %"
    vsoShp.Cells(Prop(S_DELAY)).Formula = "BlockSizeX*0.5"
    vsoShp.Cells(Prop(S_DUTYCYCLE)).Formula = "50 %"
    vsoShp.Cells(Prop(S_SIGNALSKEW)).Formula = "BlockSizeX*0.1"
  else
    vsoShp.Cells(Prop(S_DELAY)).Formula = 0
  end if
  vsoShp.Cells(Prop(S_EVENTTYPE)).Formula = "Index(0, Prop.EventType.Format)"
  vsoShp.Cells(Prop(S_LABELEDGES)).Formula = "Index(0,Prop.LabelEdges.Format)"
  vsoShp.Cells(Prop(S_LABELSIZE)).Formula = "BlockSizeX*0.667"
  vsoShp.Cells(Prop(S_LABELFONT)).Formula = "8 pt"
End Sub

' Set the Shape Data Cell Prompts
Private Sub SetDataPrompts()
  vsoShp.Cells(Prop(S_NAME) & ".Prompt").Formula = """Signal Name"""
  vsoShp.Cells(Prop(S_CLOCK) & ".Prompt").Formula = """Reference Clock"""
  vsoShp.Cells(Prop(S_SIGNAL) & ".Prompt").Formula = """Reference Signal"""
  if sType = SignalType.Clock then
    vsoShp.Cells(Prop(S_SKEW) & ".Prompt").Formula = """Percentage of period/2 to delay transitions"""
    vsoShp.Cells(Prop(S_DELAY) & ".Prompt").Formula = """Initial clock delay"""
    vsoShp.Cells(Prop(S_SIGNALSKEW) & ".Prompt").Formula = """Amount of skew to apply to dependent signals"""
  else
    vsoShp.Cells(Prop(S_DELAY) & ".Prompt").Formula = """Amount of delay before each transition"""
  end if
  vsoShp.Cells(Prop(S_EVENTTYPE) & ".Prompt").Formula = """Select an event type to add/modify"""
  vsoShp.Cells(Prop(S_EVENTTRIGGER) & ".Prompt").Formula = """Input the trigger type for this event e.g. Posedge"""
  vsoShp.Cells(Prop(S_EVENTPOSITION) & ".Prompt").Formula = """Input the position based on the trigger"""
  vsoShp.Cells(Prop(S_LABELEDGES) & ".Prompt").Formula = """Select which labels to show on transitions"""
End Sub

' Set the Shape Data Cell Labels
Private Sub SetDataLabels()
  If sType = SignalType.Bit or sType = SignalType.Clock Then
    vsoShp.Cells(Prop(S_ACTIVELOW) & ".Label").Formula = S_LABEL_ACTIVELOW
  end if
  if sType = SignalType.Clock then
    vsoShp.Cells(Prop(S_SKEW) & ".Label").Formula = S_LABEL_SKEW
    vsoShp.Cells(Prop(S_SIGNALSKEW) & ".Label").Formula = S_LABEL_SIGNALSKEW
    vsoShp.Cells(Prop(S_DUTYCYCLE) & ".Label").Formula = S_LABEL_DUTYCYCLE
  end if
  vsoShp.Cells(Prop(S_LABELEDGES) & ".Label").Formula = S_LABEL_LABELEDGES
  vsoShp.Cells(Prop(S_LABELSIZE) & ".Label").Formula = S_LABEL_LABELSIZE
  vsoShp.Cells(Prop(S_LABELFONT) & ".Label").Formula = S_LABEL_LABELFONT
  vsoShp.Cells(Prop(S_EVENTTYPE) & ".Label").Formula = S_LABEL_EVENTTYPE
  vsoShp.Cells(Prop(S_EVENTTRIGGER) & ".Label").Formula = S_LABEL_EVENTTRIGGER
  vsoShp.Cells(Prop(S_EVENTPOSITION) & ".Label").Formula = S_LABEL_EVENTPOSITION
End Sub

' Set the Shape Data Cell Types
Private Sub SetDataTypes()
  vsoShp.Cells(Prop(S_NAME) & ".Type").Formula = visPropTypeString
  vsoShp.Cells(Prop(S_CLOCK) & ".Type").Formula = visPropTypeListFix
  vsoShp.Cells(Prop(S_SIGNAL) & ".Type").Formula = visPropTypeListFixG
  if sType = SignalType.Clock or sType = SignalType.Bit then
    vsoShp.Cells("Prop.ActiveLow.Type").Formula = visPropTypeBool
  end if
  if sType = SignalType.Clock then
    vsoShp.Cells(Prop(S_PERIOD) & ".Type").Formula = visPropTypeNumber
    vsoShp.Cells(Prop(S_SKEW) & ".Type").Formula = visPropTypeNumber
  end if
  vsoShp.Cells(Prop(S_DELAY) & ".Type").Formula = visPropTypeNumber
  if sType = SignalType.Clock then
    vsoShp.Cells(Prop(S_DUTYCYCLE) & ".Type").Formula = visPropTypeNumber
    vsoShp.Cells(Prop(S_SIGNALSKEW) & ".Type").Formula = visPropTypeNumber
  end if
  vsoShp.Cells(Prop(S_EVENTTYPE) & ".Type").Formula = visPropTypeListFix
  vsoShp.Cells(Prop(S_EVENTTRIGGER) & ".Type").Formula = visPropTypeListFix
  vsoShp.Cells(Prop(S_EVENTPOSITION) & ".Type").Formula = visPropTypeNumber
  vsoShp.Cells(Prop(S_LABELEDGES) & ".Type").Formula = visPropTypeListVar
  vsoShp.Cells(Prop(S_LABELSIZE) & ".Type").Formula = visPropTypeNumber
  vsoShp.Cells(Prop(S_LABELFONT) & ".Type").Formula = visPropTypeNumber
End Sub

'//TODO. should the strings here be in strings.bas???
' Set the Shape Data Cell Formats
Private Sub SetDataFormats()
  if sType = SignalType.Clock then
    vsoShp.Cells(Prop(S_PERIOD) & ".Format").Formula = """0.00 u"""
    vsoShp.Cells(Prop(S_SKEW) & ".Format").Formula = """0.0 u"""
    vsoShp.Cells(Prop(S_SIGNALSKEW) & ".Format").Formula = """0.000 u"""
    vsoShp.Cells(Prop(S_DUTYCYCLE) & ".Format").Formula = """0.0 u"""
    vsoShp.Cells(Prop(S_LABELEDGES) & ".Format").Formula = """None;Positive;Negative;All"""
    vsoShp.Cells(Prop(S_EVENTTYPE) & ".Format").Formula = """;Delay;Node;GateX;GateZ;Gate0;Gate1"""
  end if
  vsoShp.Cells(Prop(S_LABELFONT) & ".Format").Formula = """0 u"""
  vsoShp.Cells(Prop(S_DELAY) & ".Format").Formula = """0.000 u"""
  vsoShp.Cells(Prop(S_EVENTTRIGGER) & ".Format").Formula = """Absolute;Posedge;Negedge"""
End Sub

'//TODO. update with names from strings.bas
' Set the Shape Data Cells Invisible field
Private Sub SetDataHidden()
  vsoShp.Cells(Prop(S_CLOCK) & ".Invisible").Formula = "STRSAME("""", Prop.Clock.Format)"
  vsoShp.Cells(Prop(S_SIGNAL) & ".Invisible").Formula = "STRSAME("""", Prop.Signal.Format)"
  vsoShp.Cells(Prop(S_LABELSIZE) & ".Invisible").Formula = "OR(Prop.LabelEdges.Invisible, STRSAME(""None"", Prop.LabelEdges, TRUE))"
  vsoShp.Cells(Prop(S_LABELFONT) & ".Invisible").Formula = "OR(Prop.LabelEdges.Invisible, STRSAME(""None"", Prop.LabelEdges, TRUE))"
  vsoShp.Cells(Prop(S_EVENTTRIGGER) & ".Invisible").Formula = "STRSAME("""", Prop.EventType)"
  vsoShp.Cells(Prop(S_EVENTPOSITION) & ".Invisible").Formula = "STRSAME("""", Prop.EventType)"
End Sub

' Reset the Visio Shape Data
Public Sub ResetVisioData(Optional s as Shape = Nothing)
  if Not s is Nothing then
    Set vsoShp = s
    Set shp.vsoShape = s
  End If

  ' only available in 1D shapes
  if vsoShp.CellExists("EndY", visExistsLocally) <> False then
    shp.EndY = "Guard(BeginY)"
    shp.Width = "Guard(EndX-BeginX)"
  end if

  shp.Height = "BlockSizeY"
  shp.PinX = "Guard(BeginX)"
  shp.PinY = "Guard(BeginY)"
  shp.LocPinX = 0
  shp.LocPinY = "If(FlipY,Height*1,Height*0)"

  if sType = SignalType.Bus then
    select case VW_BUS_YANCHOR
     case visAlignMiddle
      shp.LocPinY = "Height*0.5"
     case visAlignTop
      shp.LocPinY = "If(FlipY,Height*0,Height*1)"
    end select
  end if
  if sType = SignalType.Bit or sType = SignalType.Clock then
    shp.FlipY = "Prop.ActiveLow"
  end if
  vsoShp.Cells("Para.HorzAlign").Formula = visHorzRight
  vsoShp.Cells("TxtPinX").Formula = "Width*0"
  vsoShp.Cells("TxtLocPinX").Formula = "TxtWidth"
  vsoShp.Cells("TxtWidth").Formula = "LEN(SHAPETEXT(TheText))*Char.Size"
  vsoShp.Cells("FillForegndTrans").Formula = 1
  vsoShp.Cells("Geometry1.NoFill").Formula = False
End Sub

' delete the shapes children
Public Sub DeleteChildren(ParentName as String, Parent as Variant, EndX as Double)
  Dim CellName as String : CellName = User(S_PARENT)
  For Each s in Parent.Shapes
    If s.Shapes.Count > 0 Then
      DeleteChildren ParentName, s, EndX
    ElseIf s.CellExists(CellName, visExistsLocally) = True Then
      If s.Cells(CellName).ResultStr("") = ParentName And _
        s.Cells("PinX").Result("") >= EndX Then s.Delete
    End If
  Next
End Sub

' delete the child at the position X, Y
Public Sub DeleteChildAt(X as Double, Y as Double, ParentName as String, Parent as Variant)
  Dim CellName as String : CellName = User(S_PARENT)
  For Each s in Parent.Shapes
    If s.Shapes.Count > 0 Then
      DeleteChildAt X, Y, ParentName, s
    ElseIf s.CellExists(CellName, visExistsLocally) = True Then
      If s.Cells(CellName).ResultStr("") = ParentName And _
        s.Cells("PinX").Result("") = X And s.Cells("PinY").Result("") = Y Then s.Delete
    End If
  Next
End Sub

' Reset the shape geometry
Public Sub ResetGeometry(Optional s as Shape = Nothing)
  If Not s is Nothing Then
    Set vsoShp = s
    Set shp.vsoShape = s
  End If

  vsoShp.DeleteSection visSectionConnectionPts
  DeleteChildren vsoShp.Name, vsoShp.ContainingPage, shp.PinX
  vsoShp.DeleteSection visSectionScratch

  Do While vsoShp.GeometryCount > 1
    vsoShp.DeleteSection visSectionFirstComponent + 1
  Loop

  Do While vsoShp.RowCount(visSectionFirstCompont) > 3
    ' delete row 2 to keep X1=0 and X2=Width
    vsoShp.DeleteRow visSectionFirstComponent, visRowComponent + 2
  Loop

  ' specifiying the section will prevent it from being added
  AddGeometry Section:=visSectionFirstComponent
  if sType = SignalType.Bus Then AddGeometry NoFill:=True

  SetRowXY "Width*0", 0, visRowComponent + 1, visSectionFirstComponent, visTagMoveTo
  SetRowXY "Width*1", 0, visRowComponent + 2, visSectionFirstComponent, visTagLineTo

  If sType = SignalType.Clock Then ResetClock
End Sub

Public Sub ResetClock(Optional s as Shape = Nothing)
  Dim clk as vw_Clock_c
  if Not s is Nothing then
    set vsoShp = s
    sType = SignalType.Clock
  End If

  set clk = New vw_Clock_c
  clk.Start Me
End Sub

Public Sub AddGeometry(Optional Section as Integer = -1, Optional NoShow as Variant = False, _
                       Optional NoFill as Boolean = False, Optional NoLine as Boolean = False, _
                       Optional NoSnap as Boolean = False)
  If Section < 0 then
    Section = vsoShp.AddSection(visSectionFirstComponent + vsoShp.GeometryCount)
    vsoShp.AddRow Section, visRowComponent, visTagComponent
  End If
  vsoShp.CellsSRC(Section, visRowComponent, visNoFill + 0).Formula = NoFill
  vsoShp.CellsSRC(Section, visRowComponent, visNoFill + 1).Formula = NoLine
  vsoShp.CellsSRC(Section, visRowComponent, visNoFill + 2).Formula = NoShow
  vsoShp.CellsSRC(Section, visRowComponent, visNoFill + 3).Formula = NoSnap
End Sub

Public Sub SetRowXY(x as Variant, y as Variant, Row as Integer, _
                    Optional Section as Integer = visSectionFirstComponent, _
                    Optional Tag as Integer = visTagLineTo, Optional Insert as Boolean = False)
  Dim xCell as Cell
  Dim yCell as Cell
  Dim bRow as Integer

  If vsoShp.RowExists(Section, Row, visExistsLocally) = False or Insert = True Then _
    vsoShp.AddRow Section, Row, Tag

  Set xCell = vsoShp.CellsSRC(Section, Row, visX)
  Set yCell = vsoShp.CellsSRC(Section, Row, visY)

  xCell.Formula = x
  yCell.Formula = y

  If Section = visSectionFirstComponent Then
    If sType = SignalType.Clock or sType = SignalType.Bit Then
      SetRowXY xCell.Name, yCell.Name, Row - 1, visSectionConnectionPts, visTagDefault, Insert
    elseif sType = SignalType.Bus Then
      SetRowXY xCell.Name, "Height-" & yCell.Name, Row, visSectionFirstComponent + 1, Tag, Insert
      Select Case VW_BUS_YANCHOR
       case visAlignBottom, visAlignTop
        bRow = ((Row-1) *(vsoShp.RowCount(visSectionFirstComponent)-1))
        SetRowXY xCell.Name, yCell.Name, bRow + 0, visSectionConnectionPts, visTagDefault, Insert
        SetRowXY xCell.Name, vsoShp.CellsSRC(visSectionFirstComponent+1, Row, visY).Name, _
                bRow + 1, visSectionConnectionPts, visTagDefault, Insert
       case visAlignMiddle
        ' //TODO connection points needs to be centered properly
        SetRowXY xCell.Name, "Height*0.5", Row - 1, visSectionConnectionPts, visTagDefault, Insert
      End Select
    End if
  End If

End Sub

' events are added to the scratch section
Public Sub AddEvent(x as Variant, pType as EventType)
  Dim EventRow as Integer
  Dim Row as Integer
  Dim y as Variant
  Dim lblShp As Shape
  Dim TestCell as Cell
  Dim ScratchCell as Cell

  EventRow = visRowLast
  y = "Height"
  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    Set ScratchCell = vsoShp.CellsSRC(visSectionScratch, i, visScratchX)
    Set TestCell = vsoShp.Cells("User.Test")
    TestCell.Formula = x
    '//TODO should nodes and labels share the same row???
    If ScratchCell.Formula = TestCell.Formula or ScratchCell.Result("") = TestCell.Result("") Then
      Row = i
      y = vsoShp.CellsSRC(visSectionScratch, i, visScratchY).Formula
    ElseIf ScratchCell.Result("") < TestCell.Result("") Then
      Row = i + 1
      y = "Height-" & vsoShp.CellsSRC(visSectionScratch, i, visScratchY).Name
    End If
  Next i

  If EventRow = visRowLast Then Row = vsoShp.AddRow(visSectionScratch, Row, visTagDefault)

  SetRowXY x, y, Row, visSectionScratch, visTagDefault
  vsoShp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Formula = pType
End Sub

' remove events and geometry > width
Public Sub CleanUp()
  Dim GeoCnt as Integer

  GeoCnt = 1
  If sType = SignalType.Bus Then GeoCnt = 2

  ' delete children >= Width
  DeleteChildren vsoShp.Name, vsoShp.ContainingPage, shp.EndX
  ' delete events >= Width
  for i = vsoShp.RowCount(visSectionScratch) - 1 to 0 Step -1
    if vsoShp.CellsSRC(visSectionScratch, i, visScratchX).Result("") >= vsoShp.Cells("Width") Then
      vsoShp.DeleteSection visSectionFirstComponent + GeoCnt + i
      vsoShp.DeleteRow visSectionScratch, i
    end if
  Next i

  ' delete connection points > Width
  for i = vsoShp.RowCount(visSectionConnectionPts) - 1 to 0 Step -1
    if vsoShp.CellsSRC(visSectionConnectionPts, i, visX).Result("") > vsoShp.Cells("Width") then _
      vsoShp.DeleteRow visSectionConnectionPts, i
  Next i

  ' delete geometry > width
  For i = vsoShp.RowCount(visSectionFirstComponent) - 1 to 0 Step -1
    if vsoShp.CellsSRC(visSectionFirstComponent, i, visX).Result("") > vsoShp.Cells("Width") Then
      If sType = SignalType.Bus Then
        vsoShp.DeleteRow visSectionFirstComponent + 1, i
      End If
      vsoShp.DeleteRow visSectionFirstComponent, i
    end if
  Next i

End Sub

' get existing labels
Private Sub GetLabels(ByRef Labels as Collection, Parent as Variant)
  For Each s in Parent.Shapes
    If s.Shapes.Count = 0 And s.CellExists("User.Parent", visExistsLocally) = True Then
      If s.Cells("User.Parent").ResultStr("") = vsoShp.Name And s.Cells("User.Type").ResultStr("") = "Label" Then Labels.Add s
    ElseIf s.Shapes.Count > 0 Then
      GetLabels Labels, s
    End If
  Next
End Sub

Private Function GetEdgeType(Row as Integer) as String
  If vsoShp.Cells("Prop.ActiveLow").Result("") = False Then
    GetEdgeType = IIf(vsoShp.CellsSRC(visSectionScratch, Row, visY).Result("") > 0, "Positive", "Negative")
  Else
    GetEdgeType = IIf(vsoShp.CellsSRC(visSectionScratch, Row, visY).Result("") > 0, "Negative", "Positive")
  End If
End Function

'//TODO. what is making this process so slow for All
' handle the labels of the shape
Public Sub DoLabels()
  ' collection of labels that may or may not exists
  Dim Labels as Collection
  Dim Index as Integer
  Dim LblIndex as Integer
  Dim IsLblRow as Boolean
  Dim Edges as String
  Dim win as Window
  Dim Selection as Collection
  Dim iPos as Integer
  Dim iNeg as Integer
  Dim CurEdge as String

  ' get the window for selection
  For Each w in Application.Windows
    If w.Page = vsoShp.ContainingPage And w.Shape.Type = visTypePage Then Set win = w
  Next

  ' get the selected shapes
  If Not win is Nothing Then
    Set Selection = New Collection
    For Each s in win.selection
      Selection.Add s
    Next
  End If

  Set Labels = New Collection
  GetLabels Labels, vsoShp.ContainingPage

  vw_cfg.Configure
  Index = VW_LABEL_INDEX0
  LblIndex = 1
  Edges = LCase(vsoShp.Cells("Prop.LabelEdges").ResultStr(""))
  ' determine which of the scratch rows should have labels
  For i = 0 to vsoShp.RowCount(visSectionScratch) - 1
    If vsoShp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).Result("") = EventType.Edge Then
      vsoShp.CellsSRC(visSectionScratch, i, VW_COL_LABEL_HIDE).Formula = "STRSAME(Prop.LabelEdges,""None"", True)"
      ' count positive and negative for pos# and neg#
      CurEdge = GetEdgeType(CInt(i))
      If CurEdge = "Positive" Then
        iPos = iPos + 1
      Else
        iNeg = iNeg + 1
      End If

      ' decide if a label should be shown
      If Edges = "none" Then
        IsLblRow = False
      ElseIf CurEdge = "Positive" And Left(Edges,3) = "pos" And IsNumeric(Mid(Edges,4)) Then
        IsLblRow = CBool(((iPos-1) Mod Cint(Mid(Edges, 4))) = 0)
      ElseIf CurEdge = "Positive" And Left(Edges,1) = "p" And IsNumeric(Mid(Edges,2)) Then
        IsLblRow = CBool(((iPos-1) Mod Cint(Mid(Edges, 2))) = 0)
      ElseIf CurEdge = "Negative" And Left(Edges,3) = "neg" And IsNumeric(Mid(Edges,4)) Then
        IsLblRow = CBool(((iNeg-1) Mod Cint(Mid(Edges, 4))) = 0)
      ElseIf CurEdge = "Negative" And Left(Edges,1) = "n" And IsNumeric(Mid(Edges,2)) Then
        IsLblRow = CBool(((iNeg-1) Mod Cint(Mid(Edges, 2))) = 0)
      ElseIf Left(Edges,3) = "mod" And IsNumeric(Mid(Edges,4)) Then
        IsLblRow = CBool((i Mod Cint(Mid(Edges, 4))) = 0)
      ElseIf Left(Edges,1) = "m" And IsNumeric(Mid(Edges,2)) Then
        IsLblRow = CBool((i Mod Cint(Mid(Edges, 2))) = 0)
      Else
        IsLblRow = CBool((Edges = LCase(CurEdge)) Or (Edges = "all"))
      End If

      ' move the existing shape accordingly
      If LblIndex <= Labels.Count and IsLblRow = True Then
        Labels(LblIndex).Cells("PinX").Formula = vsoShp.Name & "!PinX + " & vsoShp.Name & "!Scratch.X" & Cstr(i+1)
        Labels(LblIndex).Cells("Geometry1.NoShow").Formula = vsoShp.Name & "!" & vsoShp.CellsSRC(visSectionScratch, i, VW_COL_LABEL_HIDE).Name
        Labels(LblIndex).Text = Cstr(Index)
        LblIndex = LblIndex + 1
        Index = Index + 1
      ' else create the shape
      ElseIf IsLblRow = True Then
        MakeLabel CInt(i), Index
        Index = Index + 1
      End If
    End If
  Next i

  ' delete the excess labels
  If LblIndex > 1 Then
    For i = Labels.Count to LblIndex Step -1
      Labels(i).Delete
    Next i
  End If
  Set Labels = Nothing

  ' restore the original selection
  If Not win is Nothing Then
    win.DeselectAll
    For Each s in Selection
      win.Select s, visSelect
    Next
  End If
End Sub

'//TODO. incomplete function
' add/remove a node on the shape
Private Function ModifyNode(ScratchRow as Integer, Optional Boolean as ForLabel = False) as Shape
  Dim Node as Shape
  Dim x1 as Double
  Dim x2 as Double
  Dim y1 as Double
  Dim y2 as Double

  ' create a new row for this node
  If ForLabel = False Then
    If vsoShp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS) = False Then
      vsoShp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS) = True

    ' if the row exists we want to delete this node
    ElseIf vsoShp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS).Result("") <> False Then
      vsoShp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS) = False
      x1 = vsoShp.Cells("PinX").Result("") + vsoShp.CellsSRC(visSectionScratch, ScratchRow, visX).Result("")
      y1 = vsoShp.Cells("PinY").Result("") + vsoShp.Cells("Height*0.5")
      DeleteChildAt x1, y1, vsoShp.Name, vsoShp.ContainingPage
      Exit Function

    End If
  End If

  x1 = vsoShp.Cells("PinX").Result("")
  x2 = x1 + vsoShp.Cells("Prop.LabelSize").Result("")
  y1 = vsoShp.Cells("PinY").Result("")
  y2 = y1 + vsoShp.Cells("Prop.LabelSize").Result("")

  Set Node = vsoShp.Parent.DrawRectangle(x1, x2, y1, y2)

  Node.Cells("Rounding").Formula = "0.5*Width"
  ' add user data
  Node.AddNamedRow visSectionUser, "Parent", visTagDefault
  Node.Cells("User.Parent").Formula = Chr(34) & vsoShp.Name & Chr(34)
  Node.AddNamedRow visSectionUser, "Type", visTagDefault
  Node.Cells("User.Type").Formula = """Node"""
  ' transform the label
  Node.Cells("Width").Formula = vsoShp.Name & "!Prop.LabelSize*0.75"
  Node.Cells("Height").Formula = vsoShp.Name & "!Prop.LabelSize*0.75"
  Node.Cells("PinX").Formula = vsoShp.Name & "!PinX+" & vsoShp.Name & "!Scratch.X" & Cstr(ScratchRow+1)
  Node.Cells("PinY").Formula = vsoShp.Name & "!PinY+" & vsoShp.Name & "!Scratch.Y" & CStr(ScratchRow+1)
  Node.Cells("LocPinX").Formula = "Width*0.5"
  Node.Cells("LocPinY").Formula = "Height*0.5"

  ' line and character
  Node.Cells("LinePattern").Formula = 1
  Node.Cells("LineColor").Formula = 0
  Node.Cells("LineWeight").Formula = "0.12 pt"
  Node.Cells("Char.Size").Formula = vsoShp.Name & "!Prop.LabelFont*0.75"
  Node.Cells("Char.Font").Formula = vsoShp.Name & "!Char.Font"
  Node.Cells("Char.Color").Formula = vsoShp.Name & "!Char.Color"

  Node.Cells("TxtWidth").Formula = "(LEN(SHAPETEXT(TheText))+1+" & VW_0 & ")*Char.Size"
  'Node.Text = Cstr(Index)

  Set ModifyNode = Node
End Function

' this actually draws the label on the page
Private Sub MakeLabel(ScratchRow as Integer, Index as Integer)
  Dim lbl as Shape
  Set lbl = ModifyNode(ScratchRow, True)

  Select Case VW_LABEL_SHAPE
   Case S_LBL_RECTANGLE, S_LBL_SQUARE
   Case S_LBL_RND_RECTANGLE, S_LBL_RND_SQUARE
    lbl.Cells("Rounding").Formula = "0.2 * Width"
   Case S_LBL_DIAMOND
    lbl.Cells("Angle").Formula = "45 deg"
   Case S_LBL_RND_DIAMOND
    lbl.Cells("Angle").Formula = "45 deg"
    lbl.Cells("Rounding").Formula = "0.2 * Width"
   Case S_LBL_OVAL, S_LBL_CIRCLE
    lbl.Cells("Rounding").Formula = "0.5 * Width"
  End Select

  ' add user data
  lbl.Cells("User.Type").Formula = """Label"""

  ' transform the label
  lbl.Cells("Width").Formula = vsoShp.Name & "!Prop.LabelSize"
  lbl.Cells("Height").Formula = vsoShp.Name & "!Prop.LabelSize"
  lbl.Cells("PinX").Formula = vsoShp.Name & "!PinX + " & vsoShp.Name & "!Scratch.X" & Cstr(ScratchRow+1)
  lbl.Cells("PinY").Formula = vsoShp.Name & "!PinY + " & vsoShp.Name & "!Height + Height"

  ' control visibility
  lbl.Cells("Geometry1.NoShow").Formula = vsoShp.Name & "!" & vsoShp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_LABEL_HIDE).Name
  lbl.Cells("HideText").Formula = "Geometry1.NoShow"
  lbl.Cells("Char.Size").Formula = vsoShp.Name & "!" & "Prop.LabelFont"

  ' line and character
  lbl.Cells("LinePattern").Formula = vsoShp.Name & "!LinePattern"
  lbl.Cells("LineColor").Formula = vsoShp.Name & "!LineColor"
  lbl.Cells("LineWeight").Formula = "0.24 pt"

  lbl.Text = Cstr(Index)
  '//TODO. This does not work without sheet protection
  lbl.Cells("LockSelect").Formula = True
End Sub