Attribute VB_Name = "vw_types"

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
  Edge  = 2 ^ 9  ' max 511 rows but works with integer
  ' assume posedge
  Posedge = (2 ^ 9)
  ' bit 11 = 0: positive/gate0, 1: negative/gate1
  Negedge = (2 ^ 9) Or (2 ^ 10)
  ' gate type will not include delay
  Gate0 = (2 ^ 10)
  Gate1 = (2 ^ 10) Or (2 ^ 11)
  GateX = (2 ^ 10) Or (2 ^ 12)
  GateZ = (2 ^ 10) Or (2 ^ 13)
  ' node/gap can or with Edge
  Node  = 2 ^ 14
  Gap   = 2 ^ 15
End Enum

Public Const ROW_MASK = EventType.Edge - 1
    