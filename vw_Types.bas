Attribute VB_Name = "vw_Types"

Public Enum SignalType
    Clock = 0
    Bit   = 1
    Bus   = 2
End Enum

Public Enum EventType
    Edge = 0
    Pull = 1
    GateX = 2
    GateZ = 4
End Enum
    