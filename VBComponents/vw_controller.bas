Attribute VB_Name = "vw_controller"

Public Sub CellChanged(vsoCell as IVCell)
    Dim shp as Shape
    Dim clk as vw_Clock_c
    Dim sig as vw_Signal_c

    vw_cfg.Configure
    Set shp = vsoCell.Shape

    ' register the signal
    If shp.CellExists("User.Type", visExistsLocally) = True Then
        set sig = New vw_Signal_c
        Select Case shp.Cells("User.Type").ResultStr("")
         Case VW_TYPE_STR(SignalType.Clock)
            sig.Register shp, SignalType.Clock
         Case VW_TYPE_STR(SignalType.Bit)
            sig.Register shp, SignalType.Bit
         Case VW_TYPE_STR(SignalType.Bus)
            sig.Register shp, SignalType.Bus
        End Select
    End If

    ' global changes
    Select Case vsoCell.Name
     Case "Prop.EventType"
        If vsoCell.ResultStr("") = "Node" Then
            shp.Cells("Prop.EventTrigger.Format").Formula = """Posedge;Negedge"""
        Else
            shp.Cells("Prop.EventTrigger.Format").Formula = """Absolute;Posedge;Negedge"""
        End If
     Case "Prop.EventTrigger"
     Case "Prop.LabelEdges"
        If vsoCell.ResultStr("") <> "None" Then vw_controller.DoLabels shp
    End Select

    If shp.CellExists("User.Type", visExistsLocally) = True Then
        Select Case shp.Cells("User.Type").ResultStr("")
         Case VW_TYPE_STR(SignalType.Clock)
            set clk = New vw_Clock_c
            clk.CellChanged vsoCell
        End Select
    End If
End Sub

' refresh the parents of each shape
Public Sub UpdateParents(Container as Variant)
    Dim shp as Shape
    ' if there are no shapes in the container it itself is the shape
    If Container.Shapes.Count = 0 Then
        SetSignals Container
    Else
        For Each shp in Container.Shapes
            UpdateParents shp
        Next
    End If
End Sub

' set the Prop.Clock.Format and Prop.Signal.Format fields to enable shape dependencies
Public Sub SetSignals(Child as Shape, Optional Mode as SignalType = SignalType.Void)
    Dim CurParent as String
    Dim Parents as String
    Dim s As Shape

    Dim List() as String
    Dim Value as String
    Dim Format as String

    If Mode = SignalType.Void Then
        SetSignals Child, SignalType.Clock
        SetSignals Child, SignalType.Signal
        Exit Sub
    End If

    Value  = "Prop.Signal"
    Format = "Prop.Signal.Format"
    If Mode = SignalType.Clock Then
        Value  = "Prop.Clock"
        Format = "Prop.Clock.Format"
    End If

    If Child.CellExists(Value, visExistsLocally) = False Then Exit Sub
    CurParent = Child.Cells(Value).ResultStr("")
    vw_cfg.Configure

    For Each s in Child.ContainingPage.Shapes
        If s.CellExists("User.Type", visExistsLocally) Then
            If Mode = SignalType.Clock Then
                If s.Cells("User.Type").ResultStr("") = VW_TYPE_STR(SignalType.Clock) and _
                    s.Name <> Child.Name Then Parents = Parents & s.Name & ";"
            ElseIf Mode = SignalType.Signal Then
                If s.Cells("User.Type").ResultStr("") = VW_TYPE_STR(SignalType.Bit) and _
                    s.Name <> Child.Name Then Parents = Parents & s.Name & ";"
                '//TODO. How can a bus be a parent
                'If s.Cells("User.Type").ResultStr("") = VW_TYPE_STR(SignalType.Bus) and _
                '    s.Name <> Child.Name Then Parents = s.Name & ";"
            End If
        ElseIf s.Shapes.Count > 0 Then
            If Mode = SignalType.Clock Then
                Parents = Parents & GetShapes(SignalType.Clock, s, Child.Name)
            ElseIf Mode = SignalType.Signal Then
                Parents = Parents & GetShapes(SignalType.Bit, s, Child.Name)
                '//TODO. How can a bus be a parent
                'Parents = Parents & GetShapes(SignalType.Bus, s, Child.Name)
            End If
        End If
    Next

    Child.Cells(Format).Formula = chr(34) & Parents & chr(34)
    List = Split(Parents, ";")
    Child.Cells(Value).Formula = "INDEX(" & UBound(List) & "," & Format & ")"
    For i = 0 to UBound(List)
        If CurParent = List(i) Then Child.Cells(Value).Formula = "INDEX(" & Cstr(i) & "," & Format & ")"
    Next i
End Sub

' gets all shapes of the given sType that are not the child shape
Private Function GetShapes(sType as SignalType, Parent as Shape, ChildName as String) as String
    Dim s as Shape

    If Parent.CellExists("User.Type", visExistsLocally) Then
        If Parent.Cells("User.Type").ResultStr("") = VW_TYPE_STR(sType) and _
            Parent.Name <> ChildName Then GetShapes = Parent.Name & ";"
    End If

    For Each s in Parent.Shapes
        GetShapes = GetShapes & GetShapes(sType, s, ChildName)
    Next
End Function

' handle the labels of the shape
Public Sub DoLabels(Parent as Shape)
    ' collection of labels that may or may not exists
    Dim Labels as Collection
    Dim shp As Shape
    Dim Index as Integer
    Dim LblIndex as Integer
    Dim IsLblRow as Boolean
    Dim Edges as String
    Dim win as Window
    Dim Selection as Collection

    ' Error protection
    If Parent.CellExists("Prop.LabelEdges", visExistsLocally) = False Then
        Err.Raise vbObjectError + 513, ThisDocument.Name & ":vw_controller:DoLabels", _
            "Shape is invalid shape type for this operation"
        Exit Sub
    End If

    ' get the window for selection
    For Each w in Application.Windows
        If w.Page = Parent.ContainingPage And w.Shape.Type = visTypePage Then Set win = w
    Next
    ' get the selected shapes
    If Not win is Nothing Then
        Set Selection = New Collection
        For Each s in win.selection
            Selection.Add s
        Next
    End If

    Set Labels = New Collection
    ' get the labels that exists already
    For Each shp In Parent.Parent.Shapes
        If shp.CellExists("User.Parent", visExistsLocally) = True Then
            If shp.Cells("User.Parent").ResultStr("") = Parent.Name And _
               shp.Cells("User.Type").ResultStr("") = "Label" Then Labels.Add shp
        End If
    Next

    vw_cfg.Configure
    Index = VW_LABEL_INDEX0
    LblIndex = 1
    Set shp = Parent
    Edges = shp.Cells("Prop.LabelEdges").ResultStr("")
    For i = 0 to shp.RowCount(visSectionScratch) - 1
        If shp.CellsSRC(visSectionScratch, i, VW_COL_EVENT_TYPE).Result("") = EventType.Edge Then
            shp.CellsSRC(visSectionScratch, i, VW_COL_LABEL_HIDE).Formula = "STRSAME(Prop.LabelEdges,""None"")"
            IsLblRow = True
            ' determine if this is a visible row
            If (Edges = "Positive" And shp.Cells("Prop.ActiveLow").Result("") = False) Or _
               (Edges = "Negative" And shp.Cells("Prop.ActiveLow").Result("") <> False) Then
               IsLblRow = Cbool(((i+1) And 1) <> 0)
            ElseIf (Edges = "Positive" And shp.Cells("Prop.ActiveLow").Result("") <> False) Or _
               (Edges = "Negative" And shp.Cells("Prop.ActiveLow").Result("") = False) Then
               IsLblRow = Cbool(((i+1) And 1) = 0)
            ElseIf Edges = "None" Then
                IsLblRow = False
            ElseIf Left(Edges,3) = "Mod" Then
               IsLblRow = CBool((i Mod Cint(Mid(Edges, 3))) = 0)
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
                MakeLabel shp, CInt(i), Index
                Index = Index + 1
            End If
        End If
    Next i

    If LblIndex > 1 Then
        For i = Labels.Count to LblIndex Step -1
            Labels(i).Delete
        Next i
    End If

    Set Labels = Nothing

    If Not win is Nothing Then
        win.DeselectAll
        For Each s in Selection
            win.Select s, visSelect
        Next
    End If
End Sub

' this actually draws the label on the page
Private Sub MakeLabel(shp as Shape, ScratchRow as Integer, Index as Integer)
    Dim lbl as Shape

    x1 = shp.Cells("PinX").Result("")
    x2 = x1 + shp.Cells("Prop.LabelSize").Result("")
    y1 = shp.Cells("PinY").Result("")
    y2 = y1 + shp.Cells("Prop.LabelSize").Result("")

    Set lbl = shp.Parent.DrawRectangle(x1, x2, y1, y2)

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
    lbl.AddNamedRow visSectionUser, "Parent", visTagDefault
    lbl.Cells("User.Parent").Formula = Chr(34) & shp.Name & Chr(34)
    lbl.AddNamedRow visSectionUser, "Type", visTagDefault
    '//TODO. Nodes share all of the same details except this field and Y postion
    lbl.Cells("User.Type").Formula = """Label"""

    ' transform the label
    lbl.Cells("Width").Formula = shp.Name & "!Prop.LabelSize"
    lbl.Cells("Height").Formula = shp.Name & "!Prop.LabelSize"
    lbl.Cells("PinX").Formula = shp.Name & "!PinX + " & shp.Name & "!Scratch.X" & Cstr(ScratchRow+1)
    '//TODO. Nodes do not share this position
    lbl.Cells("PinY").Formula = shp.Name & "!PinY + " & shp.Name & "!Height + Height"
    lbl.Cells("LocPinX").Formula = "Width*0.5"
    lbl.Cells("LocPinY").Formula = "Height*0.5"

    ' control visibility
    lbl.Cells("Geometry1.NoShow").Formula = shp.Name & "!" & shp.CellsSRC(visSectionScratch, ScratchRow, VW_COL_LABEL_HIDE).Name
    lbl.Cells("HideText").Formula = "Geometry1.NoShow"

    CopyParentFeatures shp, lbl
    lbl.Cells("TxtWidth").Formula = "(LEN(SHAPETEXT(TheText))+1+" & VW_0 & ")*Char.Size"
    lbl.Cells("Char.Size").Formula = shp.Name & "!" & "Prop.LabelFont"
    lbl.Text = Cstr(Index)
    '//TODO. This does not work without sheet protection
    lbl.Cells("LockSelect").Formula = True

End Sub

' makes a child shape copy certain features of the parent shape
Public Sub CopyParentFeatures(Parent as Shape, Child as Shape)
    Child.Cells("LinePattern").Formula = Parent.Name & "!LinePattern"
    Child.Cells("LineWeight").Formula = Parent.Name & "!LineWeight"
    Child.Cells("LineColor").Formula = Parent.Name & "!LineColor"
    Child.Cells("Char.Size").Formula = Parent.Name & "!Char.Size"
    Child.Cells("Char.Font").Formula = Parent.Name & "!Char.Font"
    Child.Cells("Char.Color").Formula = Parent.Name & "!Char.Color"
End Sub
