Attribute VB_Name = "vw_types"

Public Enum SignalType
  Clock  = 1
  Bit    = 2
  Bus    = 3
  Signal = 4
  Void   = 0
End Enum

Public Enum EventPosition
  Absolute = 0
  Posedge  = 1
  Negedge  = 2
End Enum

Public Enum EventType
  Edge  = 2 ^ 9  ' max 256 rows but works with integer
  Gate  = 2 ^ 10
  Node  = 2 ^ 14
  Gap   = 2 ^ 15
  ' modifiers
  m1 = 2 ^ 11
  mZ = 2 ^ 12
  mX = m1 Or mZ
  ' types II
  Posedge = Edge Or m1
  Negedge = Edge
  Gate0 = Gate
  Gate1 = Gate Or m1
  GateZ = Gate Or mZ
  GateX = Gate Or mX
End Enum

Public Const ROW_MASK = EventType.Edge - 1
    