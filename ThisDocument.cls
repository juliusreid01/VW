VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' this code would go in ThisDocument of the stencil
Private WithEvents vsoApp As Visio.Application
Public BUSY as Boolean

Private Sub Document_DocumentOpened(ByVal doc as IVDocument)
  Set vsoApp = Application
  BUSY = False
End Sub

Public Sub RefreshApp()
  if vsoApp is Nothing Then Set vsoApp = Application
End Sub

Private Sub Document_BeforeDocumentClosed(ByVal doc as IVDocument)
  Set vsoApp = Nothing
End Sub

Private Sub vsoApp_CellChanged(ByVal vsoCell As IVCell)
  vw_controller.CellChanged vsoCell
End Sub

Private Sub vsoApp_BeforeShapeDelete(ByVal vsoShp As IVShape)
  'pass the shape to a Controller for label/node deletes first
End Sub

Private Sub vsoApp_ShapeAdded(ByVal Shape As IVShape)
  'pass the shape to a Controller to update parents & children
End Sub
