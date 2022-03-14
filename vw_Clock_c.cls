VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_Clock_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' Class module to wrap around a Visio Shape to draw clocks
Private shp as Shape
Private signal as vw_Signal_c

Public Sub Class_Terminate()
    set shp = Nothing
    set signal = Nothing
End Sub

Public Sub Start(ByRef s as vw_Signal_c)
    set signal = s
    set shp = signal.vsoShp

    ' initial delay
    signal.SetRowXY "Prop.Delay", "Geometry1.Y1", 2, visSectionFirstComponent, visTagLineTo, _
                    Cbool(shp.RowCount(visSectionFirstComponent) < 4)
    ' Create events
    MyEvents
End Sub

Private Sub MyEvents()
    Dim TestCell as Cell
    ' Row to modify in scratch section
    Dim ScratchRow as Integer
    Dim y as Variant

    ' use Test Cell to compare formula results
    Set TestCell = shp.Cells("User.Test")

    If shp.RowCount(visSectionScratch) = 0 Then
        ScratchRow = 1
        TestCell.Formula = "Prop.Delay"
        y = "Height"
    Else
        ScratchRow = shp.RowCount(visSectionScratch) - 1
        TestCell.Formula = "User.ActiveWidth+Scratch.X" & Cstr(ScratchRow+1)
        If (ScratchRow and 1) = 0 Then TestCell.Formula = "Prop.Period-" & TestCell.Formula
        y = "Height-" & shp.CellsSRC(visSectionScratch, ScratchRow, visY).Name
        ScratchRow = ScratchRow + 1
    End If

    Do While TestCell.Result("") < shp.Cells("Width").Result("")
        signal.SetRowXY TestCell.Formula, y, ScratchRow - 1, visSectionScratch, visTagDefault
        shp.CellsSRC(visSectionScratch, ScratchRow - 1, VW_COL_EVENT_TYPE).Formula = _
            chr(34) & VW_EVENT_TYPE_STR(EventType.Edge) & chr(34)
        TestCell.Formula = "User.ActiveWidth+Scratch.X" & Cstr(ScratchRow)
        If (ScratchRow and 1) = 0 Then TestCell.Formula = "Prop.Period-" & TestCell.Formula
        y = "Height-" & shp.CellsSRC(visSectionScratch, ScratchRow-1, visY).Name
        ScratchRow = ScratchRow + 2
    Loop

    Redraw
End Sub

Private Sub Redraw()
    Dim InsertRow as Boolean
    Dim Row as Integer
    Row = 3

    For i = 0 to shp.RowCount(visSectionScratch) - 1
        Select Case shp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultStr("")
         Case VW_EVENT_TYPE_STR(EventType.Edge)
            InsertRow = Cbool(shp.CellsSRC(visSectionFirstComponent, Row, visX).Result("") = shp.Cells("Width"))
            If i > 0 Then
                signal.SetRowXY x:=shp.CellsSRC(visSectionScratch, i, visX).Name, _
                                y:=shp.CellsSRC(visSectionScratch, i-1, visY).Name, _
                                Row:=Row, Insert:=InsertRow
                Row = Row + 1
            End If
            InsertRow = Cbool(shp.CellsSRC(visSectionFirstComponent, Row, visX).Result("") = shp.Cells("Width"))
            signal.SetRowXY x:=shp.CellsSRC(visSectionScratch, i, visX).Name & "+User.SkewWidth", _
                            y:=shp.CellsSRC(visSectionScratch, i, visY).Name, _
                            Row:=Row, Insert:=InsertRow
            Row = Row + 1
        End Select
    Next i

    shp.CellsSRC(visSectionFirstComponent, Row, visY).Formula = "Geometry1.Y" & Cstr(Row - 1)
End Sub

Public Sub CellChanged(vsoCell as IVCell)
    If shp is Nothing Then Set shp = vsoCell.Shape
    If signal is Nothing Then
        Set signal = New vw_Signal_c
        signal.Register shp, SignalType.Clock
    End If

    Select Case vsoCell.Name
     Case "Width"
        signal.CleanUp
        MyEvents
     Case "Prop.Period"
     Case "Prop.Delay"
     Case "Prop.Skew"
     Case "Prop.DutyCycle"
     Case "Prop.Trigger"
     Case "Prop.Events"
     Case "Prop.LabelEdges"
    End Select
End Sub