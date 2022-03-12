Attribute VB_Name = "vw_cfg"

Public VW_WIDTH as double
Public VW_HEIGHT as double
Public Const VW_0IN as string = "0 in"

Public Sub Configure()
    VW_WIDTH = 3
    VW_HEIGHT = ActiveWindow.Shape.Cells("BlockSizeY").Result("")
End Sub