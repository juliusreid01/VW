Attribute VB_Name = "vw_Test"

Public Sub NewSignal_Test()
   Dim u1 As Double
   Dim signal0 As vw_Signal_c
   
   u1 = Application.BeginUndoScope("SignalTest")
   Set signal0 = New vw_Signal_c
   signal0.NewSignal activepage, SignalType.Clock, 0.5, 10.5
   Application.EndUndoScope u1, True
End Sub
