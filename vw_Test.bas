Attribute VB_Name = "vw_Test"

'//TODO. a test should draw signals manually and then compare geometry of the drawn shapes
Private Const TEST_X as Double = 0.5

Private Const TEST_CLOCK_Y0 as Double = 10
Private Const TEST_CLOCK_Y1 as Double = 10.5

Private Const TEST_BIT_Y0 as Double = 9
Private Const TEST_BIT_Y1 as Double = 9.5

Private Const TEST_BUS_Y0 as Double = 8
Private Const TEST_BUS_Y1 as Double = 8.5

Public Sub NewClock_Test()
   Dim u1 as Double
   Dim c as vw_Signal_c

   u1 = Application.BeginUndoScope("ClockTest")
   Set c = New vw_Signal_c

   c.NewSignal ActiveDocument.Pages(1), SignalType.Clock, TEST_X, TEST_CLOCK_Y1

   Application.EndUndoScope u1, True
End Sub

Public Sub NewBit_Test()
   Dim u1 as Double
   Dim b as vw_Signal_c

   u1 = Application.BeginUndoScope("BitTest")
   Set b = New vw_Signal_c

   b.NewSignal ActiveDocument.Pages(1), SignalType.Bit, TEST_X, TEST_BIT_Y1

   Application.EndUndoScope u1, True
End Sub

Public Sub NewBus_Test()
   Dim u1 as Double
   Dim b as vw_Signal_c

   u1 = Application.BeginUndoScope("BusTest")
   Set b = New vw_Signal_c

   b.NewSignal ActiveDocument.Pages(1), SignalType.Bus, TEST_X, TEST_BUS_Y1

   Application.EndUndoScope u1, True
End Sub

Public Sub New_Test()
   Dim u1 As Double
   Dim s0 As vw_Signal_c

   u1 = Application.BeginUndoScope("Test")
   Set s0 = New vw_Signal_c

   s0.NewSignal ActiveDocument.Pages(1), SignalType.Clock, TEST_X, TEST_CLOCK_Y1
   s0.NewSignal ActiveDocument.Pages(1), SignalType.Bit, TEST_X, TEST_BIT_Y1
   s0.NewSignal ActiveDocument.Pages(1), SignalType.Bus, TEST_X, TEST_BUS_Y1

   Application.EndUndoScope u1, True
End Sub

Private Sub Manual_Clock()
   Dim shp as Shape
   Set shp = ActiveDocument.Pages(1).DrawLine(TEST_X, TEST_CLOCK_Y0, TEST_X + 3, TEST_CLOCK_Y0)

End Sub