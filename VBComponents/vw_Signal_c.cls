VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_Signal_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Class module to wrap around a Visio Shape used to convert line shapes into signals

' handle to the shape being modified
Private shp as Shape
' the type of signal
Private sType as SignalType

Private Sub Class_Initialize()
  vw_cfg.Configure
End Sub

Private Sub Class_Terminate()
  Set shp = Nothing
End Sub

Public Property Get vsoShp as Shape
  set vsoShp = shp
End Property

' Create a new signal of the given type on the input page
Public Sub NewSignal(p as Page, t as SignalType, x1 as double, y1 as double)
  set shp = p.DrawLine(x1, y1, x1 + VW_WIDTH, y1)
  Init t
End Sub

' Convert a shape to the given type on the input page
Public Sub ConvertSignal(s as Shape, t as SignalType)
  set shp = s
  RemoveSections
  Init t
End Sub

' Take a shape but don't do anything
Public Sub Register(s as Shape, t as SignalType)
  set shp = s
  sType = t
End Sub

' Initialize the shape
Private Sub Init(t as SignalType)
  sType = t

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
  shp.DeleteSection visSectionProp
  shp.DeleteSection visSectionUser
  shp.DeleteSection visSectionControls
  shp.DeleteSection visSectionActions
  shp.DeleteSection visSectionConnectionPts
End Sub

' Add the User-defined Cells
Private Sub AddUserCells()
  shp.AddNamedRow visSectionUser, "Type", visTagDefault
  shp.AddNamedRow visSectionUser, "ChildOffset", visTagDefault
  if sType = SignalType.Bus or sType = SignalType.Bit then
    shp.AddNamedRow visSectionUser, "BusWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "SkewWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "Edges", visTagDefault
  elseif sType = SignalType.Clock Then
    shp.AddNamedRow visSectionUser, "ActiveWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "SkewWidth", visTagDefault
    shp.AddNamedRow visSectionUser, "Pulses", visTagDefault
  end if
  shp.AddNamedRow visSectionUser, "Test", visTagDefault
End Sub

' Set the User-defined Cells
Private Sub SetUserCells()
  shp.Cells("User.Type").Formula = chr(34) & VW_TYPE_STR(sType) & chr(34)
  shp.Cells("User.ChildOffset").Formula = "BlockSizeY"
  if sType = SignalType.Clock then
    shp.Cells("User.ActiveWidth").Formula = VW_0 & "+ (Prop.Period * Prop.DutyCycle)"
    shp.Cells("User.SkewWidth").Formula = VW_0 & "+ (Prop.Skew * Prop.Period * 0.5)"
    shp.Cells("User.Pulses").Formula = "CEILING((Width-Geometry1.X1)/Prop.Period)"
  elseif sType = SignalType.Bit then
    shp.Cells("User.BusWidth").Formula = 1
    shp.Cells("User.Edges").Formula = False
  elseif sType = SignalType.Bus then
    shp.Cells("User.BusWidth").Formula = 8
    shp.Cells("User.Edges").Formula = False
  end if
  shp.Cells("User.Test.Prompt").Formula = """Spare cell to test formulae and values"""
End Sub

' Add the Shape Data Cells
Private Sub AddDataCells()
  shp.AddNamedRow visSectionProp, "Name", visTagDefault
  shp.Cells("Prop.Name.Verify").Formula = True
  shp.AddNamedRow visSectionProp, "Clock", visTagDefault
  shp.AddNamedRow visSectionProp, "Signal", visTagDefault
  if sType = SignalType.Clock or sType = SignalType.Bit then
    shp.AddNamedRow visSectionProp, "ActiveLow", visTagDefault
  end if
  if sType = SignalType.Clock then
    shp.AddNamedRow visSectionProp, "Period", visTagDefault
    shp.AddNamedRow visSectionProp, "Skew", visTagDefault
  end if
  shp.AddNamedRow visSectionProp, "Delay", visTagDefault
  if sType = SignalType.Clock then
    shp.AddNamedRow visSectionProp, "DutyCycle", visTagDefault
    shp.AddNamedRow visSectionProp, "SignalSkew", visTagDefault
  end if
  shp.AddNamedRow visSectionProp, "EventType", visTagDefault
  shp.AddNamedRow visSectionProp, "EventTrigger", visTagDefault
  shp.AddNamedRow visSectionProp, "EventPosition", visTagDefault
  shp.AddNamedRow visSectionProp, "LabelEdges", visTagDefault
  shp.AddNamedRow visSectionProp, "LabelSize", visTagDefault
  shp.AddNamedRow visSectionProp, "LabelFont", visTagDefault
End Sub

' Set the Shape Data Cell Defaults
Public Sub ResetData(Optional s as Shape = Nothing)
  if Not s Is Nothing then set shp = s

  if sType = SignalType.Clock or sType = SignalType.Bit then
    shp.Cells("Prop.ActiveLow").Formula = False
  end if
  if sType = SignalType.Clock then
    shp.Cells("Prop.Period").Formula = "BlockSizeX*2"
    shp.Cells("Prop.Skew").Formula = "10 %"
    shp.Cells("Prop.Delay").Formula = "BlockSizeX*0.5"
    shp.Cells("Prop.DutyCycle").Formula = "50 %"
    shp.Cells("Prop.SignalSkew").Formula = "BlockSizeX*0.1"
  else
    shp.Cells("Prop.Delay").Formula = 0
  end if
  shp.Cells("Prop.EventType").Formula = "Index(0, Prop.EventType.Format)"
  shp.Cells("Prop.LabelEdges").Formula = "Index(0,Prop.LabelEdges.Format)"
  shp.Cells("Prop.LabelSize").Formula = "BlockSizeX*0.667"
  shp.Cells("Prop.LabelFont").Formula = "8 pt"
End Sub

' Set the Shape Data Cell Prompts
Private Sub SetDataPrompts()
  shp.Cells("Prop.Name.Prompt").Formula = """Signal Name"""
  shp.Cells("Prop.Clock.Prompt").Formula = """Reference Clock"""
  shp.Cells("Prop.Signal.Prompt").Formula = """Reference Signal"""
  if sType = SignalType.Clock then
    shp.Cells("Prop.Skew.Prompt").Formula = """Percentage of period/2 to delay transitions"""
    shp.Cells("Prop.Delay.Prompt").Formula = """Initial clock delay"""
    shp.Cells("Prop.SignalSkew.Prompt").Formula = """Amount of skew to apply to dependent signals"""
  else
    shp.Cells("Prop.Delay.Prompt").Formula = """Amount of delay before each transition"""
  end if
  shp.Cells("Prop.EventType.Prompt").Formula = """Select an event type to add/modify"""
  shp.Cells("Prop.EventTrigger.Prompt").Formula = """Input the trigger type for this event e.g. Posedge"""
  shp.Cells("Prop.EventPosition.Prompt").Formula = """Input the position based on the trigger"""
  shp.Cells("Prop.LabelEdges.Prompt").Formula = """Select which labels to show on transitions"""
End Sub

' Set the Shape Data Cell Labels
Private Sub SetDataLabels()
  If sType = SignalType.Bit or sType = SignalType.Clock Then
    shp.Cells("Prop.ActiveLow.Label").Formula = """Active Low"""
  end if
  if sType = SignalType.Clock then
    shp.Cells("Prop.Skew.Label").Formula = """Skew %"""
    shp.Cells("Prop.SignalSkew.Label").Formula = """Signal Skew"""
    shp.Cells("Prop.DutyCycle.Label").Formula = """Duty Cycle %"""
  end if
  shp.Cells("Prop.LabelEdges.Label").Formula = """Label Edges"""
  shp.Cells("Prop.LabelSize.Label").Formula = """Label Size"""
  shp.Cells("Prop.LabelFont.Label").Formula = """Label Font Pt"""
  shp.Cells("Prop.EventType.Label").Formula = """Event Type"""
  shp.Cells("Prop.EventTrigger.Label").Formula = """Trigger"""
  shp.Cells("Prop.EventPosition.Label").Formula = """Position"""
End Sub

' Set the Shape Data Cell Types
Private Sub SetDataTypes()
  shp.Cells("Prop.Name.Type").Formula = visPropTypeString
  shp.Cells("Prop.Clock.Type").Formula = visPropTypeListFix
  shp.Cells("Prop.Signal.Type").Formula = visPropTypeListFixG
  if sType = SignalType.Clock or sType = SignalType.Bit then
    shp.Cells("Prop.ActiveLow.Type").Formula = visPropTypeBool
  end if
  if sType = SignalType.Clock then
    shp.Cells("Prop.Period.Type").Formula = visPropTypeNumber
    shp.Cells("Prop.Skew.Type").Formula = visPropTypeNumber
  end if
  shp.Cells("Prop.Delay.Type").Formula = visPropTypeNumber
  if sType = SignalType.Clock then
    shp.Cells("Prop.DutyCycle.Type").Formula = visPropTypeNumber
    shp.Cells("Prop.SignalSkew.Type").Formula = visPropTypeNumber
  end if
  shp.Cells("Prop.EventType.Type").Formula = visPropTypeListFix
  shp.Cells("Prop.EventTrigger.Type").Formula = visPropTypeListFix
  shp.Cells("Prop.EventPosition.Type").Formula = visPropTypeNumber
  shp.Cells("Prop.LabelEdges.Type").Formula = visPropTypeListVar
  shp.Cells("Prop.LabelSize.Type").Formula = visPropTypeNumber
  shp.Cells("Prop.LabelFont.Type").Formula = visPropTypeNumber
End Sub

' Set the Shape Data Cell Formats
Private Sub SetDataFormats()
  if sType = SignalType.Clock then
    shp.Cells("Prop.Period.Format").Formula = """0.00 u"""
    shp.Cells("Prop.Skew.Format").Formula = """0.0 u"""
    shp.Cells("Prop.SignalSkew.Format").Formula = """0.000 u"""
    shp.Cells("Prop.DutyCycle.Format").Formula = """0.0 u"""
    shp.Cells("Prop.LabelEdges.Format").Formula = """None;Positive;Negative;All"""
    shp.Cells("Prop.EventType.Format").Formula = """;Delay;Node;GateX;GateZ;Gate0;Gate1"""
  end if
  shp.Cells("Prop.LabelFont.Format").Formula = """0 u"""
  shp.Cells("Prop.Delay.Format").Formula = """0.000 u"""
  shp.Cells("Prop.EventTrigger.Format").Formula = """Absolute;Posedge;Negedge"""
End Sub

' Set the Shape Data Cells Invisible field
Private Sub SetDataHidden()
  shp.Cells("Prop.Clock.Invisible").Formula = "STRSAME("""", Prop.Clock.Format)"
  shp.Cells("Prop.Signal.Invisible").Formula = "STRSAME("""", Prop.Signal.Format)"
  shp.Cells("Prop.LabelSize.Invisible").Formula = "OR(Prop.LabelEdges.Invisible, STRSAME(""None"", Prop.LabelEdges, TRUE))"
  shp.Cells("Prop.LabelFont.Invisible").Formula = "OR(Prop.LabelEdges.Invisible, STRSAME(""None"", Prop.LabelEdges, TRUE))"
  shp.Cells("Prop.EventTrigger.Invisible").Formula = "STRSAME("""", Prop.EventType)"
  shp.Cells("Prop.EventPosition.Invisible").Formula = "STRSAME("""", Prop.EventType)"
End Sub

' Reset the Visio Shape Data
Public Sub ResetVisioData(Optional s as Shape = Nothing)
  if Not s is Nothing then set shp = s

  ' only available in 1D shapes
  if shp.CellExists("EndY", visExistsLocally) <> False then
    shp.Cells("EndY").FormulaForce = "Guard(BeginY)"
    shp.Cells("Width").FormulaForce = "Guard(EndX-BeginX)"
  end if

  shp.Cells("Height").Formula = "BlockSizeY"
  shp.Cells("PinX").FormulaForce = "Guard(BeginX)"
  shp.Cells("PinY").Formula = "Guard(BeginY)"
  shp.Cells("LocPinX").Formula = 0
  shp.Cells("LocPinY").Formula = "If(FlipY,Height*1,Height*0)"
  if sType = SignalType.Bus then
    select case VW_BUS_YANCHOR
     case visAlignMiddle
      shp.Cells("LocPinY").Formula = "Height*0.5"
     case visAlignTop
      shp.Cells("LocPinY").Formula = "If(FlipY,Height*0,Height*1)"
    end select
  end if
  if sType = SignalType.Bit or sType = SignalType.Clock then
    shp.Cells("FlipY").Formula = "Prop.ActiveLow"
  end if
  shp.Cells("Para.HorzAlign").Formula = visHorzRight
  shp.Cells("TxtPinX").Formula = "Width*0"
  shp.Cells("TxtLocPinX").Formula = "TxtWidth"
  shp.Cells("TxtWidth").Formula = "LEN(SHAPETEXT(TheText))*Char.Size"
  shp.Cells("FillForegndTrans").Formula = 1
  shp.Cells("Geometry1.NoFill").Formula = False
End Sub

' delete the shapes children
Public Sub DeleteChildren(ParentName as String, Parent as Variant, EndX as Double)
  For Each s in Parent.Shapes
    If s.Shapes.Count > 0 Then
      DeleteChildren ParentName, s, EndX
    ElseIf s.CellExists("User.Parent", visExistsLocally) = True Then
      If s.Cells("User.Parent").ResultStr("") = ParentName And _
        s.Cells("PinX").Result("") >= EndX Then s.Delete
    End If
  Next
End Sub

' delete the child at the position X, Y
Public Sub DeleteChildAt(X as Double, Y as Double, ParentName as String, Parent as Variant)
  For Each s in Parent.Shapes
    If s.Shapes.Count > 0 Then
      DeleteChildAt X, Y, ParentName, s
    ElseIf s.CellExists("User.Parent", visExistsLocally) = True Then
      If s.Cells("User.Parent").ResultStr("") = ParentName And _
        s.Cells("PinX").Result("") = X And s.Cells("PinY").Result("") = Y Then s.Delete
    End If
  Next
End Sub

' Reset the shape geometry
Public Sub ResetGeometry(Optional s as Shape = Nothing)
  if Not s is Nothing then set shp = s

  shp.DeleteSection visSectionConnectionPts
  DeleteChildren shp.Name, shp.ContainingPage, shp.Cells("PinX").Result("")
  shp.DeleteSection visSectionScratch

  Do While shp.GeometryCount > 1
    shp.DeleteSection visSectionFirstComponent + 1
  Loop

  Do While shp.RowCount(visSectionFirstCompont) > 3
    ' delete row 2 to keep X1=0 and X2=Width
    shp.DeleteRow visSectionFirstComponent, visRowComponent + 2
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
    set shp = s
    sType = SignalType.Clock
  End If

  set clk = New vw_Clock_c
  clk.Start Me
End Sub

Public Sub AddGeometry(Optional Section as Integer = -1, Optional NoShow as Variant = False, _
                       Optional NoFill as Boolean = False, Optional NoLine as Boolean = False, _
                       Optional NoSnap as Boolean = False)
  If Section < 0 then
    Section = shp.AddSection(visSectionFirstComponent + shp.GeometryCount)
    shp.AddRow Section, visRowComponent, visTagComponent
  End If
  shp.CellsSRC(Section, visRowComponent, visNoFill + 0).Formula = NoFill
  shp.CellsSRC(Section, visRowComponent, visNoFill + 1).Formula = NoLine
  shp.CellsSRC(Section, visRowComponent, visNoFill + 2).Formula = NoShow
  shp.CellsSRC(Section, visRowComponent, visNoFill + 3).Formula = NoSnap
End Sub

Public Sub SetRowXY(x as Variant, y as Variant, Row as Integer, _
                    Optional Section as Integer = visSectionFirstComponent, _
                    Optional Tag as Integer = visTagLineTo, Optional Insert as Boolean = False)
  Dim xCell as Cell
  Dim yCell as Cell
  Dim bRow as Integer

  If shp.RowExists(Section, Row, visExistsLocally) = False or Insert = True Then _
    shp.AddRow Section, Row, Tag

  Set xCell = shp.CellsSRC(Section, Row, visX)
  Set yCell = shp.CellsSRC(Section, Row, visY)

  xCell.Formula = x
  yCell.Formula = y

  If Section = visSectionFirstComponent Then
    If sType = SignalType.Clock or sType = SignalType.Bit Then
      SetRowXY xCell.Name, yCell.Name, Row - 1, visSectionConnectionPts, visTagDefault, Insert
    elseif sType = SignalType.Bus Then
      SetRowXY xCell.Name, "Height-" & yCell.Name, Row, visSectionFirstComponent + 1, Tag, Insert
      Select Case VW_BUS_YANCHOR
       case visAlignBottom, visAlignTop
        bRow = ((Row-1) *(shp.RowCount(visSectionFirstComponent)-1))
        SetRowXY xCell.Name, yCell.Name, bRow + 0, visSectionConnectionPts, visTagDefault, Insert
        SetRowXY xCell.Name, shp.CellsSRC(visSectionFirstComponent+1, Row, visY).Name, _
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
  For i = 0 to shp.RowCount(visSectionScratch) - 1
    Set ScratchCell = shp.CellsSRC(visSectionScratch, i, visScratchX)
    Set TestCell = shp.Cells("User.Test")
    TestCell.Formula = x
    '//TODO should nodes and labels share the same row???
    If ScratchCell.Formula = TestCell.Formula or ScratchCell.Result("") = TestCell.Result("") Then
      Row = i
      y = shp.CellsSRC(visSectionScratch, i, visScratchY).Formula
    ElseIf ScratchCell.Result("") < TestCell.Result("") Then
      Row = i + 1
      y = "Height-" & shp.CellsSRC(visSectionScratch, i, visScratchY).Name
    End If
  Next i

  If EventRow = visRowLast Then Row = shp.AddRow(visSectionScratch, Row, visTagDefault)

  SetRowXY x, y, Row, visSectionScratch, visTagDefault
  shp.CellsSRC(visSectionScratch, Row, VW_COL_EVENT_TYPE).Formula = pType
End Sub

' remove events and geometry > width
Public Sub CleanUp()
  Dim GeoCnt as Integer

  GeoCnt = 1
  If sType = SignalType.Bus Then GeoCnt = 2

  ' delete children >= Width
  DeleteChildren shp.Name, shp.ContainingPage, shp.Cells("EndX").Result("")
  ' delete events >= Width
  for i = shp.RowCount(visSectionScratch) - 1 to 0 Step -1
    if shp.CellsSRC(visSectionScratch, i, visScratchX).Result("") >= shp.Cells("Width") Then
      shp.DeleteSection visSectionFirstComponent + GeoCnt + i
      shp.DeleteRow visSectionScratch, i
    end if
  Next i

  ' delete connection points > Width
  for i = shp.RowCount(visSectionConnectionPts) - 1 to 0 Step -1
    if shp.CellsSRC(visSectionConnectionPts, i, visX).Result("") > shp.Cells("Width") then _
      shp.DeleteRow visSectionConnectionPts, i
  Next i

  ' delete geometry > width
  For i = shp.RowCount(visSectionFirstComponent) - 1 to 0 Step -1
    if shp.CellsSRC(visSectionFirstComponent, i, visX).Result("") > shp.Cells("Width") Then
      If sType = SignalType.Bus Then
        shp.DeleteRow visSectionFirstComponent + 1, i
      End If
      shp.DeleteRow visSectionFirstComponent, i
    end if
  Next i

End Sub

' get existing labels
Private Sub GetLabels(ByRef Labels as Collection, Parent as Variant)
  For Each s in Parent.Shapes
    If s.Shapes.Count = 0 And s.CellExists("User.Parent", visExistsLocally) = True Then
      If s.Cells("User.Parent").ResultStr("") = shp.Name And s.Cells("User.Type").ResultStr("") = "Label" Then Labels.Add s
    ElseIf s.Shapes.Count > 0 Then
      GetLabels Labels, s
    End If
  Next
End Sub

Private Function GetEdgeType(Row as Integer) as String
  If shp.Cells("Prop.ActiveLow").Result("") = False Then
    GetEdgeType = IIf(shp.CellsSRC(visSectionScratch, Row, visY).Result("") > 0, "Positive", "Negative")
  Else
    GetEdgeType = IIf(shp.CellsSRC(visSectionScratch, Row, visY).Result("") > 0, "Negative", "Positive")
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
    If w.Page = shp.ContainingPage And w.Shape.Type = visTypePage Then Set win = w
  Next

  ' get the selected shapes
  If Not win is Nothing Then
    Set Selection = New Collection
    For Each s in win.selection
      Selection.Add s
    Next
  End If

  Set Labels = New Collection
  GetLabels Labels, shp.ContainingPage

  vw_cfg.Configure
  Index = VW_LABEL_INDEX0
  LblIndex = 1
  Edges = LCase(shp.Cells("Prop.LabelEdges").ResultStr(""))
  ' determine which of the scratch rows should have labels
  For i = 0 to shp.RowCount(visSectionScratch) - 1
    If shp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).Result("") = EventType.Edge Then
      shp.CellsSRC(visSectionScratch, i, VW_COL_LABEL_HIDE).Formula = "STRSAME(Prop.LabelEdges,""None"", True)"
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
        Labels(LblIndex).Cells("PinX").Formula = shp.Name & "!PinX + " & shp.Name & "!Scratch.X" & Cstr(i+1)
        Labels(LblIndex).Cells("Geometry1.NoShow").Formula = shp.Name & "!" & shp.CellsSRC(visSectionScratch, i, VW_COL_LABEL_HIDE).Name
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
    If shp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS) = False Then
      shp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS) = True

    ' if the row exists we want to delete this node
    ElseIf shp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS).Result("") <> False Then
      shp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_NODE_EXISTS) = False
      x1 = shp.Cells("PinX").Result("") + shp.CellsSRC(visSectionScratch, ScratchRow, visX).Result("")
      y1 = shp.Cells("PinY").Result("") + shp.Cells("Height*0.5")
      DeleteChildAt x1, y1, shp.Name, shp.ContainingPage
      Exit Function

    End If
  End If

  x1 = shp.Cells("PinX").Result("")
  x2 = x1 + shp.Cells("Prop.LabelSize").Result("")
  y1 = shp.Cells("PinY").Result("")
  y2 = y1 + shp.Cells("Prop.LabelSize").Result("")

  Set Node = shp.Parent.DrawRectangle(x1, x2, y1, y2)

  Node.Cells("Rounding").Formula = "0.5*Width"
  ' add user data
  Node.AddNamedRow visSectionUser, "Parent", visTagDefault
  Node.Cells("User.Parent").Formula = Chr(34) & shp.Name & Chr(34)
  Node.AddNamedRow visSectionUser, "Type", visTagDefault
  Node.Cells("User.Type").Formula = """Node"""
  ' transform the label
  Node.Cells("Width").Formula = shp.Name & "!Prop.LabelSize*0.75"
  Node.Cells("Height").Formula = shp.Name & "!Prop.LabelSize*0.75"
  Node.Cells("PinX").Formula = shp.Name & "!PinX+" & shp.Name & "!Scratch.X" & Cstr(ScratchRow+1)
  Node.Cells("PinY").Formula = shp.Name & "!PinY+" & shp.Name & "!Scratch.Y" & CStr(ScratchRow+1)
  Node.Cells("LocPinX").Formula = "Width*0.5"
  Node.Cells("LocPinY").Formula = "Height*0.5"

  ' line and character
  Node.Cells("LinePattern").Formula = 1
  Node.Cells("LineColor").Formula = 0
  Node.Cells("LineWeight").Formula = "0.12 pt"
  Node.Cells("Char.Size").Formula = shp.Name & "!Prop.LabelFont*0.75"
  Node.Cells("Char.Font").Formula = shp.Name & "!Char.Font"
  Node.Cells("Char.Color").Formula = shp.Name & "!Char.Color"

  Node.Cells("TxtWidth").Formula = "(LEN(SHAPETEXT(TheText))+1+" & VW_0 & ")*Char.Size"
  'Node.Text = Cstr(Index)

  Set ModifyNode = Node
End Function

' this actually draws the label on the page
Private Sub MakeLabel(ScratchRow as Integer, Index as Integer)
  Dim lbl as Shape
  Set lbl = ModifyNode(ScratchRow, True)

  Select Case VW_LABEL_SHAPE
   Case "Rectangle", "Square"
   Case "RoundedRectangle", "RoundedSquare"
    lbl.Cells("Rounding").Formula = "0.2 * Width"
   Case "Diamond"
    lbl.Cells("Angle").Formula = "45 deg"
   Case "RoundedDiamond"
    lbl.Cells("Angle").Formula = "45 deg"
    lbl.Cells("Rounding").Formula = "0.2 * Width"
   Case "Oval", "Circle"
    lbl.Cells("Rounding").Formula = "0.5 * Width"
  End Select

  ' add user data
  lbl.Cells("User.Type").Formula = """Label"""

  ' transform the label
  lbl.Cells("Width").Formula = shp.Name & "!Prop.LabelSize"
  lbl.Cells("Height").Formula = shp.Name & "!Prop.LabelSize"
  lbl.Cells("PinX").Formula = shp.Name & "!PinX + " & shp.Name & "!Scratch.X" & Cstr(ScratchRow+1)
  lbl.Cells("PinY").Formula = shp.Name & "!PinY + " & shp.Name & "!Height + Height"

  ' control visibility
  lbl.Cells("Geometry1.NoShow").Formula = shp.Name & "!" & shp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_LABEL_HIDE).Name
  lbl.Cells("HideText").Formula = "Geometry1.NoShow"
  lbl.Cells("Char.Size").Formula = shp.Name & "!" & "Prop.LabelFont"

  ' line and character
  lbl.Cells("LinePattern").Formula = shp.Name & "!LinePattern"
  lbl.Cells("LineColor").Formula = shp.Name & "!LineColor"
  lbl.Cells("LineWeight").Formula = "0.24 pt"

  lbl.Text = Cstr(Index)
  '//TODO. This does not work without sheet protection
  lbl.Cells("LockSelect").Formula = True
End Sub