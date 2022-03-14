Attribute VB_Name = "vw_Types"

Public Enum SignalType
    Clock  = 0
    Bit    = 1
    Bus    = 2
    Signal = 3
    Void   = -1
End Enum

Public Enum EventPosition
    Absolute = 0
    Posedge  = 1
    Negedge  = 2
End Enum

Public Enum EventType
    Edge  = 1
    Delay = 2
    Gate0 = 3
    Gate1 = 4
    GateX = 5
    GateZ = 6
End Enum
    