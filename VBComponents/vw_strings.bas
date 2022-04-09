Attribute VB_Name = "vw_strings"

' user cell strings
Public Const S_TYPE as String = "Type"
Public Const S_CHILDOFFSET as String = "ChildOffset"
Public Const S_BUSWIDTH as String = "BusWidth"
Public Const S_SKEWWIDTH as String = "SkewWidth"
Public Const S_EDGES as String = "Edges"
Public Const S_ACTIVEWIDTH as String = "ActiveWidth"
Public Const S_PULSES as String = "Pulses"
Public Const S_TEST as String = "Test"
Public Const S_PARENT as String = "Parent"

' data cell strings
Public Const S_NAME as String = "Name"
Public Const S_CLOCK as String = "Clock"
Public Const S_SIGNAL as String = "Signal"
Public Const S_ACTIVELOW as String = "ActiveLow"
Public Const S_PERIOD as String = "Period"
Public Const S_SKEW as String = "Skew"
Public Const S_DELAY as String = "Delay"
Public Const S_DUTYCYCLE as String = "DutyCycle"
Public Const S_SIGNALSKEW as String = "SignalSkew"
Public Const S_EVENTTYPE as String = "EventType"
Public Const S_EVENTTRIGGER as String = "EventTrigger"
Public Const S_EVENTPOSITION as String = "EventPosition"
Public Const S_LABELEDGES as String = "LabelEdges"
Public Const S_LABELSIZE as String = "LabelSize"
Public Const S_LABELFONT as String = "LabelFont"

' label shape strings
Public Const S_LBL_RECTANGLE as String = "Rectangle"
Public Const S_LBL_SQUARE as String = "Square"
Public Const S_LBL_DIAMOND as String = "Diamond"
Public Const S_LBL_RND_RECTANGLE as String = "RoundedRectangle"
Public Const S_LBL_RND_SQUARE as String = "RoundedSquare"
Public Const S_LBL_RND_DIAMOND as String = "RoundedDiamond"
Public Const S_LBL_OVAL as String = "Oval"
Public Const S_LBL_CIRCLE as String = "Circle"

' list items
Public Const S_LIST_NONE as String = "None"
Public Const S_LIST_ALL as String = "All"
Public Const S_LIST_POSEDGE as String = "Posedge"
Public Const S_LIST_NEGEDGE as String = "Negedge"
Public Const S_LIST_NODE as String = "Node"
Public Const S_LIST_EDGE as String = "Edge"
Public Const S_LIST_SPACER as String = "Gap"
Public Const S_LIST_DRIVE_X as String = "DriveX"
Public Const S_LIST_DRIVE_Z as String = "DriveZ"
Public Const S_LIST_DRIVE_0 as String = "Drive0"
Public Const S_LIST_DRIVE_1 as String = "Drive1"
Public Const S_LIST_ABSOLUTE as String = "Absolute"
Public Const S_LIST_DELETE as String = "Delete"

Public Function GenList(ParamArray items() as Variant)
  GenList = items(LBound(items))
  For i = LBound(items) + 1 to UBound(items)
    GenList = GenList & ";" & items(i)
  Next i
End Function