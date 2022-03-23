Attribute VB_Name = "VBAReset"

' Use after enabling Macro Settings -> Allow Programmatic Access to VBProject

Public Sub VBA_Reset()
    On Error Resume Next

   ' Replace D:\VW\ with your path of the files
    Const VW_HOME as String = "D:\VW\VBComponents\"
    Dim MyComponents as Collection
    Set MyComponents = New Collection

    ' modify this list to use in other projects
    MyComponents.Add VW_HOME & "vw_base_shape_c.cls"
    MyComponents.Add VW_HOME & "vw_base_signal_c.cls"
    MyComponents.Add VW_HOME & "vw_cfg.bas"
    MyComponents.Add VW_HOME & "vw_Clock_c.cls"
    MyComponents.Add VW_HOME & "vw_controller.bas"
    MyComponents.Add VW_HOME & "vw_Signal_c.cls"
    MyComponents.Add VW_HOME & "vw_strings.bas"
    MyComponents.Add VW_HOME & "vw_Types.bas"
    MyComponents.Add "D:\VW\Visio_Shape_Wrapper" & "vw_shape_wrapper_c.cls"

    Do While ThisDocument.VBProject.VBComponents.Count > 2
        For Each vbComp in ThisDocument.VBProject.VBComponents
            ' modify the prefix here as well
            If Left$(vbComp.Name, 3) = "vw_" Then
                ' vbext_ComponentType
                'Select Case vbComp.Type
                ' Case 1 ' vbext_ct_StdModule
                '    MyComponents.Add vbComp.Name, ".bas"
                ' Case 2 ' vbext_ct_ClassModule
                '    MyComponents.Add vbComp.Name, ".cls"
                ' Case 3 ' vbext_ct_MSForm
                '    MyComponents.Add vbComp.Name, ".frm"
                'End Select
                ThisDocument.VBProject.VBComponents.Remove vbComp
            End If
        Next
    Loop

    For Each vbComp in MyComponents
        Application.VBE.ActiveVBProject.VBComponents.Import vbComp
    Next

End Sub