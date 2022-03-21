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
   Case "Prop.ActiveLow"
    sig.DoLabels
   Case "Prop.EventType"
    If vsoCell.ResultStr("") = "Node" Then
      shp.Cells("Prop.EventPosition.Type").Formula = visPropTypeListFix
      shp.Cells("Prop.EventTrigger.Format").Formula = """Posedge;Negedge"""
    ElseIf vsoCell.ResultStr("") = "Delay" Then
      shp.Cells("Prop.EventPosition.Type").Formula = visPropTypeListFix
      shp.Cells("Prop.EventTrigger.Format").Formula = """Pulse"""
    Else
      shp.Cells("Prop.EventTrigger.Format").Formula = """Absolute;Posedge;Negedge"""
    End If
   Case "Prop.EventTrigger"
    If shp.Cells("Prop.EventTrigger").ResultStr("") = "Absolute" Then
      shp.Cells("Prop.EventPosition.Type").Formula = visPropTypeListVar
    Else
      shp.Cells("Prop.EventPosition.Type").Formula = visPropTypeListFix
    End If
   Case "Prop.LabelEdges"
    If vsoCell.ResultStr("") <> "None" Then sig.DoLabels
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
    Set shp = Container
    SetSignals shp
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
    If s.CellExists("User.Type", visExistsLocally) And s.Name <> Child.Name Then
      If s.CellExists("Prop.Clock", visExistsLocally) = True Then
        If s.Cells("Prop.Clock").ResultStr("") <> Child.Name And s.Cells("Prop.Signal").ResultStr("") <> Child.Name Then
          If Mode = SignalType.Clock Then
            If s.Cells("User.Type").ResultStr("") = VW_TYPE_STR(SignalType.Clock) Then Parents = Parents & s.Name & ";"
          ElseIf Mode = SignalType.Signal Then
            If s.Cells("User.Type").ResultStr("") = VW_TYPE_STR(SignalType.Bit) Then Parents = Parents & s.Name & ";"
            '//TODO. How can a bus be a parent
            'If s.Cells("User.Type").ResultStr("") = VW_TYPE_STR(SignalType.Bus) and _
            '    s.Name <> Child.Name Then Parents = s.Name & ";"
          End If
        End If
      End If
    ElseIf s.Shapes.Count > 0 Then
      Parents = Parents & GetShapes(Mode, s, Child.Name)
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

