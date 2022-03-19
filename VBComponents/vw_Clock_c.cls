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
    Dim Start as Integer

    ' use Test Cell to compare formula results
    Set TestCell = shp.Cells("User.Test")

    ' determine which row to begin reading from scratch and formula to use
    If shp.RowCount(visSectionScratch) = 0 Then
        ScratchRow = 0
        TestCell.Formula = "Prop.Delay"
        y = "Height"
        Start = 0
    Else
        ScratchRow = shp.RowCount(visSectionScratch)
        TestCell.Formula = "User.ActiveWidth+Scratch.X" & Cstr(ScratchRow)
        If (ScratchRow and 1) = 0 Then TestCell.Formula = "Prop.Period-" & TestCell.Formula
        y = "Height-" & shp.CellsSRC(visSectionScratch, ScratchRow-1, visY).Name
        Start = ScratchRow
    End If

    ' updates the scratch section with events
    Do While TestCell.Result("") < shp.Cells("Width").Result("")
        signal.SetRowXY TestCell.Formula, y, ScratchRow, visSectionScratch, visTagDefault
        shp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_EVENT_TYPE).Formula = EventType.Edge                
        ScratchRow = ScratchRow + 1
        TestCell.Formula = "User.ActiveWidth+Scratch.X" & Cstr(ScratchRow)
        If (ScratchRow and 1) = 0 Then TestCell.Formula = "Prop.Period-" & TestCell.Formula
        y = "Height-" & shp.CellsSRC(visSectionScratch, ScratchRow - 1, visY).Name
    Loop

    Redraw Start
End Sub

Private Sub Redraw(Optional Start as Integer = 0)
    Dim InsertRow as Boolean
    Dim Row as Integer
    Dim NoShow as Variant
    Row = 3

    If Start > 0 Then Row = 2 * (Start + 1)

    ' read the scratch section to update the geometry
    For i = Start to shp.RowCount(visSectionScratch) - 1
        Select Case shp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).Result("")
         Case EventType.Edge
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
            '//TODO. Is this common code
            ' draw an edge for line ending manipulation
            If shp.SectionExists(visSectionFirstComponent + i + 1, visExistsLocally) = False Then
                If ((i+1) and 1) = 0 Then
                    NoShow = "IF(EndArrow=0,TRUE,FALSE)"
                Else
                    NoShow = "IF(BeginArrow=0,TRUE,FALSE)"
                End If
                signal.AddGeometry NoFill:=True, NoShow:=NoShow
            End If
            signal.SetRowXY x:="Geometry1.X" & Cstr(Row), y:="Geometry1.Y" & Cstr(Row), Row:=1, _
                            Section:=visSectionFirstComponent + i + 1, Tag:=visTagMoveTo
            signal.SetRowXY x:="Scratch.X" & Cstr(i+1), y:="Height-Scratch.Y" & Cstr(i+1), Row:=2, _
                            Section:=visSectionFirstComponent + i + 1, Tag:=visTagLineTo
            signal.SetRowXY x:="Geometry1.X" & Cstr(Row), y:="Geometry1.Y" & Cstr(Row), Row:=3, _
                            Section:=visSectionFirstComponent + i + 1, Tag:=visTagLineTo

            Row = Row + 1
        End Select
    Next i

    Row = shp.RowCount(visSectionFirstComponent) - 1
    if shp.CellsSRC(visSectionFirstComponent, Row - 1, visX).Result("") = shp.CellsSRC(visSectionFirstComponent, Row, visX).Result("") then
        shp.DeleteRow visSectionFirstComponent, Row - 1
    end if

    shp.CellsSRC(visSectionFirstComponent, Row, visY).Formula = "Geometry1.Y" & Cstr(Row - 1)
End Sub

Public Sub CellChanged(vsoCell as IVCell)
    If shp is Nothing Then Set shp = vsoCell.Shape
    If signal is Nothing Then
        Set signal = New vw_Signal_c
        signal.Register shp, SignalType.Clock
    End If

    Select Case vsoCell.Name
     Case "Width", "Prop.Period", "Prop.Delay"
        signal.CleanUp
        MyEvents
     Case "Prop.Skew", "Prop.DutyCycle"
        If VW_LIMIT_PERCENT Then
            if vsoCell.Result(visUnitCodes.visPercent) > 100 then
                vsoCell.Formula = "100 %"
            elseif vsoCell.Result("") < 0 then
                vsoCell.Formula = "0 % - " & vsoCell.Result("")
            elseif vsoCell.Result(visUnitCodes.visPercent) < 1 then
                vsoCell.Formula = 100 * vsoCell.Result(visUnitCodes.visPercent) & "%"
            end if
        End If
     Case "Prop.Trigger"
     Case "Prop.Events"
     Case "Prop.LabelEdges"
    End Select
End Sub