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

Public Sub DrawMe(ByRef signal as vw_Signal_c)
    Dim NextRow as Integer
    Dim PrevCnt as Integer
    Dim GeoCnt As Integer
    Dim lblShp as Shape

    Dim shp as Shape
    Dim TestCell as Cell
    Dim Row as Integer

    set shp = signal.vsoShp

    ' initial delay
    If shp.RowCount(visSectionFirstComponent) < 4 Then
        shp.AddRow visSectionFirstComponent, 2, visTagLineTo
        shp.AddRow visSectionFirstComponent, 1, visTagDefault
    End If
    signal.SetRowXY "Geometry1.X1+Prop.Delay", "Geometry1.Y1", 2, visSectionFirstComponent, visTagLineTo

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

End Sub