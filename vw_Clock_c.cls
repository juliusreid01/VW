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
    Dim TestCell as Cell
    Dim Row as Integer

    set signal = s
    set shp = signal.vsoShp

    ' initial delay
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

    Redraw True
End Sub

Private Sub Redraw(Init as Boolean)
    Dim Row as Integer
    Row = 3

    For i = 0 to shp.RowCount(visSectionScratch) - 1
        Select Case shp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).ResultStr("")
         Case VW_EVENT_TYPE_STR(EventType.Edge)
            If i > 0 Then
                signal.SetRowXY x:=shp.CellsSRC(visSectionScratch, i, visX).Name, _
                                y:=shp.CellsSRC(visSectionScratch, i-1, visY).Name, _
                                Row:=Row, Insert:=Init
                Row = Row + 1
            End If
            signal.SetRowXY x:=shp.CellsSRC(visSectionScratch, i, visX).Name & "+User.SkewWidth", _
                            y:=shp.CellsSRC(visSectionScratch, i, visY).Name, _
                            Row:=Row, Insert:=Init
            Row = Row + 1
        End Select
    Next i
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
     Case "Prop.Period"
     Case "Prop.Delay"
     Case "Prop.Skew"
     Case "Prop.DutyCycle"
     Case "Prop.Trigger"
     Case "Prop.Events"
     Case "Prop.LabelEdges"
    End Select
End Sub