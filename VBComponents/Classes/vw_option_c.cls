VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vw_option_c"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' class to store option data

Private pName as String
Private pLabel as String
Private pDesc as String
Private pType as Integer
Private pFormat as String
Private pValue as Variant
Private pHidden as Variant
Private pTry as Boolean

Private Sub Class_Initialize()
  pLabel = ""
  pDesc = ""
  pType = -1
  pFormat = ""
  pValue = ""
  pHidden = False
  pTry = False
End Sub

Public Property Get Name() as String: Name = pName: End Property

Public Property Get Label() as String: Label = pLabel: End Property
Public Property Let Label(s as String): pLabel = s: End Property

Public Property Get Description() as String: Description = pDesc: End Property
Public Property Let Description(s as String): pDesc = s: End Property
Public Property Get Prompt() as String: Prompt = pDesc: End Property
Public Property Let Prompt(s as String): pDesc = s: End Property

Public Property Get TypeInt() as Integer: TypeInt = pType: End Property
Public Property Let TypeInt(i as Integer): pType = i: End Property

Public Property Get Format() as String: Format = pFormat: End Property
Public Property Let Format(s as String): pFormat = s: End Property

Public Property Get Value() as Variant: Value = pValue: End Property
Public Property Let Value(v as Variant): pValue = v: End Property

Public Property Get Hidden() as Variant: Hidden = pHidden: End Property
Public Property Let Hidden(b as Variant): pHidden = b: End Property

Public Property Let Try(b as Boolean): pTry = b: End Property

Public Sub Create(Name as String, Optional Desc as String = "")
  pName = Name
  pLabel = Name
  pDesc = Desc
End Sub

Public Property Let User(s as Shape)
  Dim Row as Integer
  If s.CellExists("User." & pName, visExistsLocally) = False Then
    Row = s.AddNamedRow(visSectionUser, pName, visTagDefault)
  Else
    Row = s.Cells("User." & pName).Row
  End If
  If pDesc <> "" Then s.CellsSRC(visSectionUser, Row, visUserPrompt).Formula = Chr(34) & pDesc &Chr(34)
  If pValue <> "" Then s.CellsSRC(visSectionuser, Row, visUserValue).Formula = pValue
End Property

Public Property Let Data(s as Shape)
  Dim Row as Integer
  If s.CellExists("Prop." & pName, visExistsLocally) = False Then
    Row = s.AddNamedRow(visSectionProp, pName, visTagDefault)
  Else
    Row = s.Cells("Prop." & pName).Row
  End If
  If pLabel <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsLabel).Formula = Chr(34) & pLabel &Chr(34)
  If pDesc <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsPrompt).Formula = Chr(34) & pDesc &Chr(34)
  If pType >= 0 Then s.CellsSRC(visSectionProp, Row, visCustPropsType).Formula = pType
  If pFormat <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsFormat).Formula = Chr(34) & pFormat &Chr(34)
  If pValue <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsValue).Formula = pValue
  s.CellsSRC(visSectionProp, Row, visCustPropsInvis).Formula = pHidden
End Property

Public Property Let DefaultData(s as Shape)
  Dim Row as Integer
  If s.CellExists("Prop." & pName, visExistsLocally) = False Then
    Row = s.AddNamedRow(visSectionProp, pName, visTagDefault)
  Else
    Row = s.Cells("Prop." & pName).Row
  End If
  If pLabel <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsLabel).Formula = Chr(34) & "Default " & pLabel &Chr(34)
  If pDesc <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsPrompt).Formula = Chr(34) & "Default " & pDesc &Chr(34)
  If pType >= 0 Then s.CellsSRC(visSectionProp, Row, visCustPropsType).Formula = pType
  If pFormat <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsFormat).Formula = Chr(34) & pFormat &Chr(34)
  '//TODOE can we take the string from the cell and use it in another cell's formula
  If pValue <> "" Then s.CellsSRC(visSectionProp, Row, visCustPropsValue).Formula = pValue
  s.CellsSRC(visSectionProp, Row, visCustPropsInvis).Formula = False
End Property