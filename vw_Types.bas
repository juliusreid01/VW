Attribute VB_Name = "vw_Types"

Public Enum SignalType
    Clock = 0
    Bit   = 1
    Bus   = 2
End Enum

Public Enum EventType
    Edge = 0
    Gate0 = 1
    Gate1 = 2
    GateX = 4
    GateZ = 8
End Enum
    