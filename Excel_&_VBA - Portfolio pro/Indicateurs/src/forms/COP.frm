' Module : COP.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_MJCOP_CNX_Click()
COP.Hide
End Sub

Private Sub CB_MJCOP_VAL_Click()

Range("COP_1") = TB_COP_1.Value
Range("COP_2") = TB_COP_2.Value
Range("COP_3") = TB_COP_3.Value
Range("COP_4") = TB_COP_4.Value
Range("COP_5") = TB_COP_5.Value
Range("COP_6") = TB_COP_6.Value
Range("COP_7") = TB_COP_7.Value
Range("COP_8") = TB_COP_8.Value
Range("COP_9") = TB_COP_9.Value
Range("COP_10") = TB_COP_10.Value
Range("COP_11") = TB_COP_11.Value
Range("COP_12") = TB_COP_12.Value
Range("COP_13") = TB_COP_13.Value
Range("COP_14") = TB_COP_14.Value
Range("COP_15") = TB_COP_15.Value

Sheets("DATA_BASE").Visible = True

Range("H3:H17").Select
    ActiveWorkbook.Worksheets("DATA_BASE").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DATA_BASE").Sort.SortFields.Add Key:=Range("H3"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA_BASE").Sort
        .SetRange Range("H3:H17")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Sheets("DATA_BASE").Visible = xlSheetVeryHidden

COP.Hide

End Sub

Private Sub Userform_Activate()

TB_COP_1.Value = Range("COP_1").Text
TB_COP_2.Value = Range("COP_2").Text
TB_COP_3.Value = Range("COP_3").Text
TB_COP_4.Value = Range("COP_4").Text
TB_COP_5.Value = Range("COP_5").Text
TB_COP_6.Value = Range("COP_6").Text
TB_COP_7.Value = Range("COP_7").Text
TB_COP_8.Value = Range("COP_8").Text
TB_COP_9.Value = Range("COP_9").Text
TB_COP_10.Value = Range("COP_10").Text
TB_COP_11.Value = Range("COP_11").Text
TB_COP_12.Value = Range("COP_12").Text
TB_COP_13.Value = Range("COP_13").Text
TB_COP_14.Value = Range("COP_14").Text
TB_COP_15.Value = Range("COP_15").Text

End Sub
