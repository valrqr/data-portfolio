' Module : CARBU.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_ANN_CAL_Click()

TB_CA_LF.BackColor = &H80000005
TB_CA_MF.BackColor = &H80000005
TB_CA_LG.BackColor = &H80000005
TB_CA_MG.BackColor = &H80000005
TB_CA_LS.BackColor = &H80000005
TB_CA_MS.BackColor = &H80000005
TB_CA_AS.BackColor = &H80000005
TB_CA_MAS.BackColor = &H80000005
TB_CA_LA.BackColor = &H80000005
TB_CA_MLA.BackColor = &H80000005

CARBU.Hide

End Sub

Private Sub CB_M_CARB_Change()

Range("MOIS_X") = CB_M_CARB.ListIndex + 1

TB_CA_LF.Text = Range("DMX_152")
TB_CA_MF.Text = Range("DMX_153")
TB_CA_LG.Text = Range("DMX_154")
TB_CA_MG.Text = Range("DMX_155")
TB_CA_LS.Text = Range("DMX_156")
TB_CA_MS.Text = Range("DMX_157")
TB_CA_AS.Text = Range("DMX_158")
TB_CA_MAS.Text = Range("DMX_159")
TB_CA_LA.Text = Range("DMX_160")
TB_CA_MLA.Text = Range("DMX_161")


End Sub
Private Sub CB_VAL_CAL_Click()

Dim ADD_LS, ADD_MS, ADD_LF, ADD_MF, ADD_LG, ADD_MG, ADD_AS, ADD_MAS, ADD_LA, ADD_MLA, ADD_152, ADD_153, ADD_154, ADD_155, ADD_156, ADD_157, ADD_158, ADD_159, ADD_160, ADD_161

ADD_LS = Range("ADD_MX_156").Value
ADD_MS = Range("ADD_MX_157").Value
ADD_LF = Range("ADD_MX_152").Value
ADD_MF = Range("ADD_MX_153").Value
ADD_LG = Range("ADD_MX_154").Value
ADD_MG = Range("ADD_MX_155").Value
ADD_AS = Range("ADD_MX_158").Value
ADD_MAS = Range("ADD_MX_159").Value
ADD_LA = Range("ADD_MX_160").Value
ADD_MLA = Range("ADD_MX_161").Value

ADD_156 = Range("ADD_MCR_156").Value
ADD_157 = Range("ADD_MCR_157").Value
ADD_152 = Range("ADD_MCR_152").Value
ADD_153 = Range("ADD_MCR_153").Value
ADD_154 = Range("ADD_MCR_154").Value
ADD_155 = Range("ADD_MCR_155").Value
ADD_158 = Range("ADD_MCR_158").Value
ADD_159 = Range("ADD_MCR_159").Value
ADD_160 = Range("ADD_MCR_160").Value
ADD_161 = Range("ADD_MCR_161").Value


If TB_CA_LF.Text = "" Or TB_CA_MF.Text = "" Or TB_CA_LG.Text = "" Or TB_CA_MG.Text = "" Or TB_CA_LS.Text = "" Or TB_CA_MS.Text = "" Or TB_CA_AS.Text = "" Or TB_CA_MAS.Text = "" Or TB_CA_LA.Text = "" Or TB_CA_MLA.Text = "" Then

MsgBox "Un ou plusieurs champs sont vides, complêtez pour valider"

Else

Sheets("LOG_ANNUEL").Range(ADD_LS) = TB_CA_LS.Value
Sheets("LOG_ANNUEL").Range(ADD_MS) = TB_CA_MS.Value
Sheets("LOG_ANNUEL").Range(ADD_LF) = TB_CA_LF.Value
Sheets("LOG_ANNUEL").Range(ADD_MF) = TB_CA_MF.Value
Sheets("LOG_ANNUEL").Range(ADD_LG) = TB_CA_LG.Value
Sheets("LOG_ANNUEL").Range(ADD_MG) = TB_CA_MG.Value
Sheets("LOG_ANNUEL").Range(ADD_AS) = TB_CA_AS.Value
Sheets("LOG_ANNUEL").Range(ADD_MAS) = TB_CA_MAS.Value
Sheets("LOG_ANNUEL").Range(ADD_LA) = TB_CA_LA.Value
Sheets("LOG_ANNUEL").Range(ADD_MLA) = TB_CA_MLA.Value

Sheets("LOG_M (2)").Range(ADD_156) = TB_CA_LS.Value
Sheets("LOG_M (2)").Range(ADD_157) = TB_CA_MS.Value
Sheets("LOG_M (2)").Range(ADD_152) = TB_CA_LF.Value
Sheets("LOG_M (2)").Range(ADD_153) = TB_CA_MF.Value
Sheets("LOG_M (2)").Range(ADD_154) = TB_CA_LG.Value
Sheets("LOG_M (2)").Range(ADD_155) = TB_CA_MG.Value
Sheets("LOG_M (2)").Range(ADD_158) = TB_CA_AS.Value
Sheets("LOG_M (2)").Range(ADD_159) = TB_CA_MAS.Value
Sheets("LOG_M (2)").Range(ADD_160) = TB_CA_LA.Value
Sheets("LOG_M (2)").Range(ADD_161) = TB_CA_MLA.Value


CARBU.Hide

End If

Range("ETS_CAR") = 1

ActiveWorkbook.Save

End Sub


Private Sub CommandButton1_Click()
Range("ESSAI_F244") = TB_CA_LF.Value
End Sub

Private Sub TB_CA_LF_Change()
If TB_CA_LF.Text = "" Then
TB_CA_LF.BackColor = &H8080FF
Else
TB_CA_LF.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MF_Change()
If TB_CA_MF.Text = "" Then
TB_CA_MF.BackColor = &H8080FF
Else
TB_CA_MF.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_LG_Change()
If TB_CA_LG.Text = "" Then
TB_CA_LG.BackColor = &H8080FF
Else
TB_CA_LG.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MG_Change()
If TB_CA_MG.Text = "" Then
TB_CA_MG.BackColor = &H8080FF
Else
TB_CA_MG.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_LS_Change()
If TB_CA_LS.Text = "" Then
TB_CA_LS.BackColor = &H8080FF
Else
TB_CA_LS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MS_Change()
If TB_CA_MS.Text = "" Then
TB_CA_MS.BackColor = &H8080FF
Else
TB_CA_MS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_AS_Change()
If TB_CA_AS.Text = "" Then
TB_CA_AS.BackColor = &H8080FF
Else
TB_CA_AS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MAS_Change()
If TB_CA_MAS.Text = "" Then
TB_CA_MAS.BackColor = &H8080FF
Else
TB_CA_MAS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_LA_Change()
If TB_CA_LA.Text = "" Then
TB_CA_LA.BackColor = &H8080FF
Else
TB_CA_LA.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MLA_Change()
If TB_CA_MLA.Text = "" Then
TB_CA_MLA.BackColor = &H8080FF
Else
TB_CA_MLA.BackColor = &H80000005
End If
End Sub
