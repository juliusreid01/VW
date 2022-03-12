Attribute VB_Name = "vw_cfg"

Public Const VW_0IN as string = "0 in"

Public VW_WIDTH as double
Public VW_HEIGHT as double
Public VW_TYPE_STR(0 to 2) as string
Public Const VW_BUS_YANCHOR as Integer = visAlignBottom

Public Sub Configure()
    VW_WIDTH = 3
    VW_HEIGHT = ActiveWindow.Shape.Cells("BlockSizeY").Result("")
    VW_TYPE_STR(SignalType.Clock) = """Clock"""
    VW_TYPE_STR(SignalType.Bit) = """Bit"""
    VW_TYPE_STR(SignalType.Bus) = """Bus"""
End Sub