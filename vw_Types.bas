Attribute VB_Name = "vw_Types"

Public Enum SignalType
    Clock = 0
    Bit   = 1
    Bus   = 2
End Enum

Public Enum EventPosition
    Absolute = 0
    Posedge  = 1
    Negedge  = 2
End Enum

Public Enum EventType
    Edge  = 0
    Delay = 1
    Gate0 = 2
    Gate1 = 3
    GateX = 4
    GateZ = 5
End Enum
    