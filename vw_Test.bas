Attribute VB_Name = "vw_Test"

Public Sub NewSignal_Test()
   Dim u1 As Double
   Dim s0 As vw_Signal_c

   u1 = Application.BeginUndoScope("SignalTest")
   Set s0 = New vw_Signal_c

   s0.NewSignal activepage, SignalType.Clock, 0.5, 10.5
   s0.NewSignal activepage, SignalType.Bit, 0.5, 10
   s0.NewSignal activepage, SignalType.Bus, 0.5, 9.5

   Application.EndUndoScope u1, True
End Sub