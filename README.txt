Current Visio Waves project directory

After enabling Macro Settings -> Allow Programmatic Access to VBProject
Put this in the immediate window, replace D:\VW with your path of the files

Application.VBE.ActiveVBProject.VBComponents.Import "D:\VW\vw_cfg.bas"
Application.VBE.ActiveVBProject.VBComponents.Import "D:\VW\vw_Clock_c.cls"
Application.VBE.ActiveVBProject.VBComponents.Import "D:\VW\vw_controller.bas"
Application.VBE.ActiveVBProject.VBComponents.Import "D:\VW\vw_Signal_c.cls"
Application.VBE.ActiveVBProject.VBComponents.Import "D:\VW\vw_Test.bas"
Application.VBE.ActiveVBProject.VBComponents.Import "D:\VW\vw_Types.bas"

vw_controller.CellChanged ActivePage.Shapes("Sheet.1").Cells("Width")

