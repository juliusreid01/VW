VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vw_frm_cfg 
   Caption         =   "Configuration Options"
   ClientHeight    =   7380
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4632
   OleObjectBlob   =   "vw_frm_cfg.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "vw_frm_cfg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private ShowMore as Boolean
Private p as Page

Private Sub UserForm_Initialize()
  Shrink
  cfg = False
End Sub

Private Sub UserFrom_Terminate()
  Set p = Nothing
End Sub

Private Sub Shrink()
  ShowMore = True
  cmdMore.Caption = "More Info"
  Me.Height = 84.6
  lstDetails.ListIndex = -1
End Sub

Private Sub Expand()
  ShowMore = False
  cmdMore.Caption = "Less Info"
  If lstDetails.ListCount = 0 Then GenerateList
  Me.Height = 390
End Sub

Private Sub GenerateList()
  Dim idx as Integer
  idx = 0

  AddItem S_CHILDOFFSET, "BlockSizeY (" & ActivePage.PageSheet.Cells("BlockSizeY").Result("") & ")", False, idx
  AddItem S_LABELSIZE, "BlockSizeX*2/3 (" & (ActivePage.PageSheet.Cells("BlockSizeX").Result("") * 0.67) & ")", True, idx
  AddItem S_LABELFONT, "8 pt", True, idx
  AddItem S_LBL_SHAPE, S_LBL_SQUARE, False, idx
  AddItem S_NODESIZEMULT, "0.5", False, idx
  AddItem S_NODEFONT, "8 pt", False, idx

End Sub

Private Function AddItem(Name as String, Value as Variant, Always as Boolean, ByRef idx as Integer)
  lstDetails.AddItem Name
  lstDetails.List(idx, 1) = Value
  lstDetails.List(idx, 2) = IIF(Always, "Yes", "No")
  idx = idx + 1
End Function

Private Sub cmdNo_Click()
  If p is Nothing Then Set p = ActivePage
  vw_page_cfg.InitPageCfg p, False
  Unload Me
End Sub

Private Sub cmdYes_Click()
  If p is Nothing Then Set p = ActivePage
  vw_page_cfg.InitPageCfg p, True
  Unload Me
End Sub

Private Sub cmdMore_Click()
  If ShowMore = True Then
    Expand
  Else
    Shrink
  End If
End Sub

Private Sub lstDetails_Change()
  If lblDetails.ListIndex < 0 Then
    lblDetails.Caption = "Click any option above to get more information"
    Exit Sub
  End If
  Select Case lstDetails.List(lstDetails.ListIndex, 0)
   Case S_CHILDOFFSET : lblDetails.Caption = _
    "When a signal depends on another this is the minimum distance between them"
   Case S_LABELSIZE: lblDetails.Caption = _
    "The height of the labels on the page"
   Case S_LABELFONT: lblDetails.Caption = _
    "The font size of the labels created"
   Case S_LBL_SHAPE: lblDetails.Caption = _
    "The shape of the labels created e.g. Square, Circle, Diamond"
   Case S_NODESIZEMULT: lblDetails.Caption = _
    "The multiplier for the size of nodes, the actual height of the node is the parent shape's height X this number"
   Case S_NODEFONT: lblDetails.Caption = _
    "The font size of the nodes created"
   Case Else : lblDetails.Caption = "TBD"
  End Select
End Sub

