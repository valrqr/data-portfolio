' Module : FAT.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_ANN_FNA_Click()
FAT.Hide
End Sub
Private Sub CB_VAL_FNA_Click()

Dim ADD_FAT, ADD_NOT, ADD_MAIP, ADD_SAIP, ADD_MRR, ADD_GRU, ADD_138, ADD_139, ADD_140, ADD_141, ADD_142, ADD_143

ADD_FAT = Range("ADD_MX_138").Value
ADD_NOT = Range("ADD_MX_139").Value
ADD_MAIP = Range("ADD_MX_140").Value
ADD_SAIP = Range("ADD_MX_141").Value
ADD_MRR = Range("ADD_MX_142").Value
ADD_GRU = Range("ADD_MX_143").Value

ADD_138 = Range("ADD_MCR_138").Value
ADD_139 = Range("ADD_MCR_139").Value
ADD_140 = Range("ADD_MCR_140").Value
ADD_141 = Range("ADD_MCR_141").Value
ADD_142 = Range("ADD_MCR_142").Value
ADD_143 = Range("ADD_MCR_143").Value

Sheets("LOG_ANNUEL").Range(ADD_FAT) = TB_FAT.Text
Sheets("LOG_ANNUEL").Range(ADD_NOT) = TB_NOT.Text
Sheets("LOG_ANNUEL").Range(ADD_MAIP) = TB_MAIP.Text
Sheets("LOG_ANNUEL").Range(ADD_SAIP) = TB_SAIP.Text
Sheets("LOG_ANNUEL").Range(ADD_MRR) = TB_MRR.Text
Sheets("LOG_ANNUEL").Range(ADD_GRU) = TB_GRU.Text

Sheets("LOG_M (2)").Range(ADD_138) = TB_FAT.Text
Sheets("LOG_M (2)").Range(ADD_139) = TB_NOT.Text
Sheets("LOG_M (2)").Range(ADD_140) = TB_MAIP.Text
Sheets("LOG_M (2)").Range(ADD_141) = TB_SAIP.Text
Sheets("LOG_M (2)").Range(ADD_142) = TB_MRR.Text
Sheets("LOG_M (2)").Range(ADD_143) = TB_GRU.Text


FAT.Hide

Range("MLG_FAT") = TB_FAT.Text
Range("MLG_NOT") = TB_NOT.Text
Range("MLG_FATA") = TB_FATA.Text
Range("MLG_AIP1") = TB_SAIP.Text
Range("MLG_AIP2") = TB_MAIP.Text
Range("MLG_MRR") = TB_MRR.Text
Range("MLG_GRU") = TB_GRU.Text


Range("CR_FATT") = TB_FATT.Text
Range("CR_NOTA") = TB_NOT.Text
Range("CR_MRRT") = TB_MRRT.Text
Range("CR_GRUA") = TB_GRU.Text
Range("CR_GRUT") = TB_GRUT.Text

If TB_FATA.Text = "" Then

Range("CR_FATA") = 0
Else

Range("CR_FATA") = TB_FATA.Text

End If

If TB_MRR.Text = "" Then

Range("CR_MRRA") = 0

Else

Range("CR_MRRA") = TB_MRR.Text

End If

If TB_GRU.Text = "" Then

Range("CR_GRUA") = 0

Else

Range("CR_GRUA") = TB_GRU.Text

End If


Range("ETS_TVX") = 1

ActiveWorkbook.Save

End Sub
Private Sub MOIS_FNA_Change()

Range("MOIS_X") = MOIS_FNA.ListIndex + 1

TB_FAT.Text = Range("DMX_138")
TB_NOT.Text = Range("DMX_139")
TB_MAIP.Text = Range("DMX_140")
TB_SAIP.Text = Range("DMX_141")
TB_MRR.Text = Range("DMX_142")
TB_GRU.Text = Range("DMX_143")

End Sub
