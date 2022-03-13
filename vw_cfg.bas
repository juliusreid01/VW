Attribute VB_Name = "vw_cfg"

' unit dependent
Public Const VW_0 as string = "0 in"

Public VW_WIDTH as double
Public VW_HEIGHT as double
Public VW_TYPE_STR(0 to 2) as string
Public VW_EVENT_TYPE_STR(0 to 10) as string

Public Const VW_BUS_YANCHOR as Integer = visAlignBottom
'Public Const VW_BUS_YANCHOR as Integer = visAlignMiddle
'Public Const VW_BUS_YANCHOR as Integer = visAlignTop

Public Const VW_COL_EVENT_TYPE as Integer = visScratchA
Public Const VW_COL_LABEL_NAME as Integer = visScratchB
Public Const VW_COL_NODE_NAME as Integer = visScratchC
Public Const VW_COL_EVENT_SHOW as Integer = visScratchD

Public Sub Configure()
    VW_WIDTH = 3
    VW_HEIGHT = ActiveWindow.Shape.Cells("BlockSizeY").Result("")
    VW_TYPE_STR(SignalType.Clock) = """Clock"""
    VW_TYPE_STR(SignalType.Bit) = """Bit"""
    VW_TYPE_STR(SignalType.Bus) = """Bus"""

    VW_EVENT_TYPE_STR(EventType.Edge) = "Edge"
    VW_EVENT_TYPE_STR(EventType.Gate0) = "Gate0"
    VW_EVENT_TYPE_STR(EventType.Gate1) = "Gate1"
    VW_EVENT_TYPE_STR(EventType.GateX) = "GateX"
    VW_EVENT_TYPE_STR(EventType.GateZ) = "GateZ"
End Sub