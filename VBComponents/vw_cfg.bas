Attribute VB_Name = "vw_cfg"

' unit dependent
Public Const VW_0 as string = "0 in"

' default signal width
Public VW_WIDTH as double
' default signal height
Public VW_HEIGHT as double
' convert the enumeration to a string
Public VW_TYPE_STR(0 to 2) as string

' control if DutyCycle and Skew should be limited by Percent
Public Const VW_LIMIT_PERCENT as Boolean = True
' control if a Rectangle or Oval is used for labels
Public Const VW_LABEL_SHAPE as String = "Oval"
Public Const VW_LABEL_INDEX0 as Integer = 1

' controls LocPinY and Connection points
Public Const VW_BUS_YANCHOR as Integer = visAlignBottom
'Public Const VW_BUS_YANCHOR as Integer = visAlignMiddle
'Public Const VW_BUS_YANCHOR as Integer = visAlignTop

Public Const VW_COL_EVENT_TYPE as Integer = visScratchA
Public Const VW_COL_LABEL_HIDE as Integer = visScratchB
Public Const VW_COL_NODE_NAME as Integer = visScratchC
'Public Const VW_COL_EVENT_SHOW as Integer = visScratchD

' configure the defaults call before using them
Public Sub Configure()
    VW_WIDTH = 3
    VW_HEIGHT = ActiveWindow.Shape.Cells("BlockSizeY").Result("")
    VW_TYPE_STR(SignalType.Clock) = "Clock"
    VW_TYPE_STR(SignalType.Bit) = "Bit"
    VW_TYPE_STR(SignalType.Bus) = "Bus"
End Sub