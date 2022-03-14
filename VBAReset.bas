Attribute VB_Name = "VBAReset"

' Use after enabling Macro Settings -> Allow Programmatic Access to VBProject

Public Sub VBA_Reset()
    On Error Resume Next

   ' Replace D:\VW\ with your path of the files
    Const VW_HOME as String = "D:\VW\VBComponents\"
    Dim MyComponents as Object
    Set MyComponents = CreateObject("Scripting.Dictionary")

    ' modify this list to use in other projects
    MyComponents.Add "vw_cfg", ".bas"
    MyComponents.Add "vw_Clock_c", ".cls"
    MyComponents.Add "vw_controller", ".bas"
    MyComponents.Add "vw_Signal_c", ".cls"
    MyComponents.Add "vw_Test", ".bas"
    MyComponents.Add "vw_Types", ".bas"

    For Each vbComp in ThisDocument.VBProject.VBComponents
        ' modify the prefix here as well
        If Left$(vbComp.Name, 3) = "vw_" Then
            ' vbext_ComponentType
            Select Case vbComp.Type
             Case 1 ' vbext_ct_StdModule
                MyComponents.Add vbComp.Name, ".bas"
             Case 2 ' vbext_ct_ClassModule
                MyComponents.Add vbComp.Name, ".cls"
             Case 3 ' vbext_ct_MSForm
                MyComponents.Add vbComp.Name, ".frm"
            End Select
            ThisDocument.VBProject.VBComponents.Remove vbComp
        End If
    Next

    For Each key in MyComponents.Keys
        Application.VBE.ActiveVBProject.VBComponents.Import VW_HOME & key & MyComponents.Item(key)
    Next

End Sub