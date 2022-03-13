Attribute VB_Name = "vw_Test"

Public Sub NewClock_Test()
   Dim u1 as Double
   Dim c as vw_Signal_c

   u1 = Application.BeginUndoScope("ClockTest")
   Set c = New vw_Signal_c

   c.NewSignal ActiveDocument.Pages(1), SignalType.Clock, 0.5, 10.5

   Application.EndUndoScope u1, True
End Sub

Public Sub NewBit_Test()
   Dim u1 as Double
   Dim b as vw_Signal_c

   u1 = Application.BeginUndoScope("BitTest")
   Set b = New vw_Signal_c

   b.NewSignal ActiveDocument.Pages(1), SignalType.Bit, 0.5, 10

   Application.EndUndoScope u1, True
End Sub

Public Sub NewBus_Test()
   Dim u1 as Double
   Dim b as vw_Signal_c

   u1 = Application.BeginUndoScope("BusTest")
   Set b = New vw_Signal_c

   b.NewSignal ActiveDocument.Pages(1), SignalType.Bus, 0.5, 9.5

   Application.EndUndoScope u1, True
End Sub

Public Sub New_Test()
   Dim u1 As Double
   Dim s0 As vw_Signal_c

   u1 = Application.BeginUndoScope("Test")
   Set s0 = New vw_Signal_c

   s0.NewSignal ActiveDocument.Pages(1), SignalType.Clock, 0.5, 10.5
   s0.NewSignal ActiveDocument.Pages(1), SignalType.Bit, 0.5, 10
   s0.NewSignal ActiveDocument.Pages(1), SignalType.Bus, 0.5, 9.5

   Application.EndUndoScope u1, True
End Sub