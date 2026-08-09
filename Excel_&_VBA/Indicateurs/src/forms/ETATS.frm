' Module : ETATS.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub DPU_PLUS_Click()
TB_DPU.Value = TB_DPU.Value + 1
End Sub

Private Sub ETA_PLUS_Click()
TB_ETA2.Value = TB_ETA2.Value + 1
End Sub

Private Sub ETC_PLUS_Click()
TB_ETC2.Value = TB_ETC2.Value + 1
End Sub

Private Sub ETV_PLUS_Click()
TB_ETV2.Value = TB_ETV2.Value + 1
End Sub

Private Sub Userform_Activate()

TB_ETV2.BackColor = &H80000005
CB_ETV2.Value = False
TB_ETV2.Visible = False
TB_ETV2.Text = ""

TB_ETA2.BackColor = &H80000005
CB_ETA2.Value = False
TB_ETA2.Visible = False
TB_ETA2.Text = ""

TB_ETC2.BackColor = &H80000005
CB_ETC2.Value = False
TB_ETC2.Visible = False
TB_ETC2.Text = ""

TB_DPU.BackColor = &H80000005
CB_DPU.Value = False
TB_DPU.Visible = False
TB_DPU.Text = ""

End Sub

Private Sub CB_ET_CNX_Click()

TB_ETV2.BackColor = &H80000005
CB_ETV2.Value = False
TB_ETV2.Visible = False
TB_ETV2.Text = ""

TB_ETA2.BackColor = &H80000005
CB_ETA2.Value = False
TB_ETA2.Visible = False
TB_ETA2.Text = ""

TB_ETC2.BackColor = &H80000005
CB_ETC2.Value = False
TB_ETC2.Visible = False
TB_ETC2.Text = ""

TB_DPU.BackColor = &H80000005
CB_DPU.Value = False
TB_DPU.Visible = False
TB_DPU.Text = ""

ETATS.Hide

End Sub
Private Sub CB_ET_VAL_Click()

If CB_ETV2.Value = True And TB_ETV2.Text = "" Then

MsgBox "Le nombre d'Etat de Veille est manquant"
TB_POR.BackColor = &H8080FF

ElseIf CB_ETA2.Value = True And TB_ETA2.Text = "" Then

MsgBox "Le nombre d'Etat d'Alerte est manquant"
TB_VTF.BackColor = &H8080FF

ElseIf CB_ETC2.Value = True And TB_ETC2.Text = "" Then

MsgBox "Le nombre d'Etat d'Accident est manquant"
TB_NEI.BackColor = &H8080FF

ElseIf CB_DPU.Value = True And TB_DPU.Text = "" Then

MsgBox "Le nombre de Déclenchement Plan d'Urgence est manquant"
TB_AUT3.BackColor = &H8080FF

Else

Dim ADD_ETV, ADD_ETA, ADD_ETC, ADD_DPU

ADD_ETV = Range("ADD_JJ_16").Value
ADD_ETA = Range("ADD_JJ_17").Value
ADD_ETC = Range("ADD_JJ_18").Value
ADD_DPU = Range("ADD_JJ_19").Value

Sheets("LOG_M").Range(ADD_ETV) = TB_ETV2.Text
Sheets("LOG_M").Range(ADD_ETA) = TB_ETA2.Text
Sheets("LOG_M").Range(ADD_ETC) = TB_ETC2.Text
Sheets("LOG_M").Range(ADD_DPU) = TB_DPU.Text



End If

ETATS.Hide

End Sub

Private Sub CB_ETV2_Click()

If CB_ETV2.Value = True Then

TB_ETV2.Visible = True
ETV_PLUS.Visible = True
TB_ETV2.Text = "1"

Else

TB_ETV2.Text = ""
TB_ETV2.Visible = False
ETV_PLUS.Visible = False
TB_ETV2.BackColor = &H80000005

End If

End Sub
Private Sub CB_ETA2_Click()

If CB_ETA2.Value = True Then

TB_ETA2.Visible = True
ETA_PLUS.Visible = True
TB_ETA2.Text = "1"

Else

TB_ETA2.Text = ""
TB_ETA2.Visible = False
ETA_PLUS.Visible = False
TB_ETA2.BackColor = &H80000005

End If

End Sub
Private Sub CB_ETC2_Click()

If CB_ETC2.Value = True Then

TB_ETC2.Visible = True
ETC_PLUS.Visible = True
TB_ETC2.Text = "1"

Else

TB_ETC2.Text = ""
TB_ETC2.Visible = False
ETC_PLUS.Visible = False
TB_ETC2.BackColor = &H80000005

End If

End Sub
Private Sub CB_DPU_Click()

If CB_DPU.Value = True Then

TB_DPU.Visible = True
DPU_PLUS.Visible = True
TB_DPU.Text = "1"

Else

TB_DPU.Text = ""
TB_DPU.Visible = False
DPU_PLUS.Visible = False
TB_DPU.BackColor = &H80000005

End If

End Sub
Private Sub TB_ETV2_Change()
If TB_ETV2.Text = "" Then
TB_ETV2.BackColor = &H8080FF
Else
TB_ETV2.BackColor = &H80000005
End If
End Sub
Private Sub TB_ETA2_Change()
If TB_ETA2.Text = "" Then
TB_ETA2.BackColor = &H8080FF
Else
TB_ETA2.BackColor = &H80000005
End If
End Sub
Private Sub TB_ETC2_Change()
If TB_ETC2.Text = "" Then
TB_ETC2.BackColor = &H8080FF
Else
TB_ETC2.BackColor = &H80000005
End If
End Sub
Private Sub TB_DPU_Change()
If TB_DPU.Text = "" Then
TB_DPU.BackColor = &H8080FF
Else
TB_DPU.BackColor = &H80000005
End If
End Sub
