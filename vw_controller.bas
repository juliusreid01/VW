Attribute VB_Name = "vw_controller"

Public Sub CellChanged(vsoCell as IVCell)
    Dim shp as Shape
    Dim clk as vw_Clock_c

    vw_cfg.Configure
    Set shp = vsoCell.Shape

    If shp.CellExists("User.Type", visExistsLocally) = True Then
        Select Case shp.Cells("User.Type").ResultStr("")
         Case VW_TYPE_STR(SignalType.Clock)
            set clk = New vw_Clock_c
            clk.CellChanged vsoCell
        End Select
    End If
End Sub