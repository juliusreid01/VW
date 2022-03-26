VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "visio_shape_wrapper_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' wrapper for visio shapes to reduce code line length
Private s As Shape

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
  Set s = Nothing
End Sub

Public Property Set vsoShape(vsoShp As Shape)
  Set s = vsoShp
End Property

Public Sub Delete()
  s.Delete
  Set s = Nothing
End Sub

Public Property Get Name() As String
  Name = s.Name
End Property

Public Property Let Name(v As String)
  s.Name = v
End Property

Public Property Get Width() As Variant
  Width = s.Cells("Width").Result("")
End Property

Public Property Let Width(v As Variant)
  s.Cells("Width").FormulaForce = v
End Property

Public Property Get Height() As Variant
  Height = s.Cells("Height").Result("")
End Property

Public Property Let Height(v As Variant)
  s.Cells("Height").FormulaForce = v
End Property

Public Property Get Angle() As Variant
  Angle = s.Cells("Angle").Result("")
End Property

Public Property Let Angle(v As Variant)
  s.Cells("Angle").FormulaForce = v
End Property

Public Property Get PinX() As Variant
  PinX = s.Cells("PinX").Result("")
End Property

Public Property Let PinX(v As Variant)
  s.Cells("PinX").FormulaForce = v
End Property

Public Property Get PinY() As Variant
  PinY = s.Cells("PinY").Result("")
End Property

Public Property Let PinY(v As Variant)
  s.Cells("PinY").FormulaForce = v
End Property

Public Property Get LocPinX() As Variant
  LocPinX = s.Cells("LocPinX").Result("")
End Property

Public Property Let LocPinX(v As Variant)
  s.Cells("LocPinX").FormulaForce = v
End Property

Public Property Get LocPinY() As Variant
  LocPinY = s.Cells("LocPinY").Result("")
End Property

Public Property Let LocPinY(v As Variant)
  s.Cells("LocPinY").FormulaForce = v
End Property

Public Property Get FlipY() As Variant
  FlipY = s.Cells("FlipY").Result("")
End Property

Public Property Let FlipY(v As Variant)
  s.Cells("FlipY").FormulaForce = v
End Property

' 1-D Shape Cells
Public Property Get BeginX() As Variant
  If s.CellExists("BeginX", visExistsLocally) = False Then
    BeginX = PinX - LocPinX
  Else
    BeginX = s.Cells("BeginX").Result("")
  End If
End Property

Public Property Let BeginX(v As Variant)
  If s.CellExists("BeginX", visExistsLocally) = False Then
    s.Cells("PinX").FormulaForce = v & "+" & LocPinX
  Else
    s.Cells("BeginX").FormulaForce = v
  End If
End Property

Public Property Get BeginY() As Variant
  If s.CellExists("BeginY", visExistsLocally) = False Then
    BeginY = PinY - LocPinY
  Else
    BeginY = s.Cells("BeginY").Result("")
  End If
End Property

Public Property Let BeginY(v As Variant)
  If s.CellExists("BeginY", visExistsLocally) = False Then
    s.Cells("PinY").FormulaForce = v & "+" & LocPinY
  Else
    s.Cells("BeginY").FormulaForce = v
  End If
End Property

Public Property Get EndX() As Variant
  If s.CellExists("EndX", visExistsLocally) = False Then
    EndX = PinX + LocPinX
  Else
    EndX = s.Cells("EndX").Result("")
  End If
End Property

Public Property Let EndX(v As Variant)
  If s.CellExists("EndX", visExistsLocally) = False Then
    s.Cells("PinX").FormulaForce = v & "-" & LocPinX
  Else
    s.Cells("EndX").FormulaForce = v
  End If
End Property

Public Property Get EndY() As Variant
  If s.CellExists("EndY", visExistsLocally) = False Then
    EndY = PinY + LocPinY
  Else
    EndY = s.Cells("EndY").Result("")
  End If
End Property

Public Property Let EndY(v As Variant)
  If s.CellExists("EndY", visExistsLocally) = False Then
    s.Cells("PinY").FormulaForce = v & "-" & LocPinY
  Else
    s.Cells("EndY").FormulaForce = v
  End If
End Property

Public Function GetCellRefName(CellName As String) As String
  GetCellRefName = s.Name & "!" & CellName
End Function

Public Sub SetPoint(Section As Integer, Row As Integer, X As Variant, Y As Variant)
  s.CellsSRC(Section, Row, visX).FormulaForce = X
  s.CellsSRC(Section, Row, visY).FormulaForce = Y
End Sub

' Expanded functions for vw_waves
Public Sub AddUserCell(Name As String, Optional Prompt As String = "")
  If s.CellExists("User." & Name, visExistsLocally) = False Then _
    s.AddNamedRow visSectionUser, Name, visTagDefault

  If Prompt <> "" Then s.Cells("User." & Name & ".Prompt").Formula = Prompt
End Sub

Public Sub SetUserCell(Name As String, v As Variant)
  s.Cells("User." & Name).Formula = v
End Sub

Public Sub AddDataCell(Name As String, Optional Prompt As String = "")
  If s.CellExists("Prop." & Name, visExistsLocally) = False Then _
    s.AddNamedRow visSectionProp, Name, visTagDefault

  If Prompt <> "" Then s.Cells("Prop." & Name & ".Prompt").Formula = Prompt
End Sub

Public Sub SetDataCell(Name As String, v As Variant, Optional Position As String = "")
  If Position = "" Then
    s.Cells("Prop." & Name).Formula = v
  Else
    s.Cells("Prop." & Name & "." & Position).Formula = v
  End If
End Sub

Public Sub SetCell(Name As String, v As Variant)
  If s.CellExists(Name, visExistsLocally) <> False Then
    s.Cells(Name).FormulaForce = v
  ElseIf s.CellExists("Prop." & Name, visExistsLocally) <> False Then
    s.Cells("Prop." & Name).FormulaForce = v
  ElseIf s.CellExists("User." & Name, visExistsLocally) <> False Then
    s.Cells("User." & Name).FormulaForce = v
  Else
    Err.Raise vbObjectError + 512, "visio_shape_wrapper_c", "Cell: " & Name & " not found!"
  End If
End Sub
