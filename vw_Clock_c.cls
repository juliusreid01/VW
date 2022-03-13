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

Public Sub Start(ByRef signal as vw_Signal_c)
    Dim TestCell as Cell
    Dim Row as Integer

    set shp = signal.vsoShp

    ' initial delay
    If shp.RowCount(visSectionFirstComponent) < 4 Then
        'shp.AddRow visSectionFirstComponent, 2, visTagLineTo
        'shp.AddRow visSectionConnectionPts, 1, visTagDefault
    End If
    signal.SetRowXY "Geometry1.X1+Prop.Delay", "Geometry1.Y1", 2, visSectionFirstComponent, visTagLineTo, _
                    Cbool(shp.RowCount(visSectionFirstComponent) < 4)

    ' user a test cell to prevent exceeding the shape's width
    Set TestCell = shp.Cells("User.Test")
    Row = 1

    TestCell.Formula = "Prop.Delay"
    Do While TestCell.Result("") < shp.Cells("Width").Result("")
        signal.AddEvent TestCell.Formula, EventType.Edge
        TestCell.Formula = "User.ActiveWidth+Scratch.X" & Cstr(Row)
        If (Row and 1) = 0 Then TestCell.Formula = "Prop.Period-" & TestCell.Formula
        Row = Row + 1
    Loop

    signal.CleanUp

    'Redraw
End Sub

Private Sub Redraw()
    Dim x as Variant
    Dim y as Variant
    Dim Row as Integer
    Row = 3

    For i = 0 to shp.RowCount(visSectionScratch) - 1
        Select Case shp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE)
         Case VW_EVENT_TYPE_STR(EventType.Edge)
            x = shp.CellsSRC(visSectionScratch, i, visX).Name & "+Prop.SkewWidth"
            y = shp.CellsSRC(visSectionScratch, i, visX).Formula
        End Select
    Next i
End Sub