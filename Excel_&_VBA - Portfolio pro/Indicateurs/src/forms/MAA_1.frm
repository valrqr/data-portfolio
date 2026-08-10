' Module : MAA_1.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub AUT3_PLUS_Click()
TB_AUT3.Value = TB_AUT3.Value + 1
End Sub

Private Sub CDM_PLUS_Click()
TB_CDMR.Value = TB_CDMR.Value + 1
End Sub

Private Sub NEI_PLUS_Click()
TB_NEI.Value = TB_NEI.Value + 1
End Sub

Private Sub POR_PLUS_Click()
TB_POR.Value = TB_POR.Value + 1
End Sub

Private Sub Userform_Activate()

CB_POR.Value = False
TB_POR.Visible = False
TB_POR.Text = ""

CB_VTF.Value = False
TB_VTF.Visible = False
TB_VTF.Text = ""

CB_NEI.Value = False
TB_NEI.Visible = False
TB_NEI.Text = ""

CB_AUT3.Value = False
TB_AUT3.Visible = False
TB_AUT3.Text = ""

CB_CDM.Value = False
TB_CDMR.Visible = False
TB_CDMR.Text = ""

LB_CDM.Visible = False
TB_ICDM.Visible = False
TB_ICDM.Text = ""
TB_TICDM.Visible = False
TB_TICDM.Text = ""

End Sub

Private Sub CB_MAA_CNX_Click()

TB_POR.BackColor = &H80000005
CB_POR.Value = False
TB_POR.Visible = False
TB_POR.Text = ""

TB_VTF.BackColor = &H80000005
CB_VTF.Value = False
TB_VTF.Visible = False
TB_VTF.Text = ""

TB_NEI.BackColor = &H80000005
CB_NEI.Value = False
TB_NEI.Visible = False
TB_NEI.Text = ""

TB_AUT3.BackColor = &H80000005
CB_AUT3.Value = False
TB_AUT3.Visible = False
TB_AUT3.Text = ""

TB_CDMR.BackColor = &H80000005
CB_CDM.Value = False
TB_CDMR.Visible = False
TB_CDMR.Text = ""

LB_CDM.Visible = False
TB_ICDM.BackColor = &H80000005
TB_ICDM.Visible = False
TB_ICDM.Text = ""

TB_TICDM.BackColor = &H80000005
TB_TICDM.Visible = False
TB_TICDM.Text = ""

MAA_1.Hide

End Sub

Private Sub CB_MAA_VAL_Click()


If CB_POR.Value = True And TB_POR.Text = "" Then

MsgBox "Le nombre de MAA Pluie/Orage est manquant"
TB_POR.BackColor = &H8080FF

ElseIf CB_VTF.Value = True And TB_VTF.Text = "" Then

MsgBox "Le nombre de MAA Vent Fort est manquant"
TB_VTF.BackColor = &H8080FF

ElseIf CB_NEI.Value = True And TB_NEI.Text = "" Then

MsgBox "Le nombre de MAA Neige est manquant"
TB_NEI.BackColor = &H8080FF

ElseIf CB_AUT3.Value = True And TB_AUT3.Text = "" Then

MsgBox "Le nombre de MAA Autre est manquant"
TB_AUT3.BackColor = &H8080FF

ElseIf CB_CDM.Value = True And TB_CDMR.Text = "" Then

MsgBox "Le nombre de MAA Coup de Mer est manquant"
TB_CDM.BackColor = &H8080FF

ElseIf CB_CDM.Value = True And TB_ICDM.Text = "" Then

MsgBox "Vous devez saisir le nombre d'inspection(s) Coup de Mer pour poursuivre"
TB_ICDM.BackColor = &H8080FF

ElseIf CB_CDM.Value = True And TB_TICDM.Text = "" Then

MsgBox "Vous devez saisir le nombre d'inspection(s) Coup de Mer pour poursuivre"
TB_TICDM.BackColor = &H8080FF


Else

Dim ADD_PO, ADD_VF, ADD_NE, ADD_CM, ADD_IM, ADD_AU, ADD_TCDM

ADD_PO = Range("ADD_JJ_145").Value
ADD_VF = Range("ADD_JJ_146").Value
ADD_NE = Range("ADD_JJ_147").Value
ADD_CM = Range("ADD_JJ_149").Value
ADD_IM = Range("ADD_JJ_150").Value
ADD_TCDM = Range("ADD_JJ_151").Value
ADD_AU = Range("ADD_JJ_148").Value

Sheets("LOG_M").Range(ADD_PO) = TB_POR.Text
Sheets("LOG_M").Range(ADD_VF) = TB_VTF.Text
Sheets("LOG_M").Range(ADD_NE) = TB_NEI.Text
Sheets("LOG_M").Range(ADD_CM) = TB_CDMR.Text
Sheets("LOG_M").Range(ADD_IM) = TB_ICDM.Text
Sheets("LOG_M").Range(ADD_TCDM) = TB_TICDM.Text
Sheets("LOG_M").Range(ADD_AU) = TB_AUT3.Text

End If

MAA_1.Hide

End Sub
Private Sub CB_POR_Click()

If CB_POR.Value = True Then

TB_POR.Visible = True
POR_PLUS.Visible = True
TB_POR.Text = "1"

Else

TB_POR.Text = ""
TB_POR.Visible = False
POR_PLUS.Visible = False
TB_POR.BackColor = &H80000005

End If
End Sub
Private Sub CB_VTF_Click()

If CB_VTF.Value = True Then

TB_VTF.Visible = True
VTF_PLUS.Visible = True
TB_VTF.Text = "1"

Else

TB_VTF.Text = ""
TB_VTF.Visible = False
VTF_PLUS.Visible = False
TB_VTF.BackColor = &H80000005

End If
End Sub
Private Sub CB_NEI_Click()

If CB_NEI.Value = True Then

TB_NEI.Visible = True
NEI_PLUS.Visible = True
TB_NEI.Text = "1"

Else

TB_NEI.Text = ""
TB_NEI.Visible = False
NEI_PLUS.Visible = False
TB_NEI.BackColor = &H80000005

End If
End Sub
Private Sub CB_AUT3_Click()

If CB_AUT3.Value = True Then

TB_AUT3.Visible = True
AUT3_PLUS.Visible = True
TB_AUT3.Text = "1"

Else

TB_AUT3.Text = ""
TB_AUT3.Visible = False
AUT3_PLUS.Visible = False
TB_AUT3.BackColor = &H80000005

End If
End Sub
Private Sub CB_CDM_Click()

If CB_CDM.Value = True Then

TB_CDMR.Visible = True
TB_ICDM.Visible = True
TB_TICDM.Visible = True
LB_CDM.Visible = True
CDM_PLUS.Visible = True
TB_CDMR.Text = "1"

Else

TB_CDMR.Text = ""
TB_CDMR.Visible = False
TB_CDMR.BackColor = &H80000005
TB_ICDM.Visible = False
TB_TICDM.Visible = False
LB_CDM.Visible = False
CDM_PLUS.Visible = False

End If
End Sub
Private Sub TB_CDM_Change()
If TB_CDMR.Value <> "" Then
TB_ICDM.Visible = True
TB_TICDM.Visible = True
LB_CDM.Visible = True
Else
TB_ICDM.Visible = False
TB_TICDM.Visible = False
LB_CDM.Visible = False
End If
End Sub

Private Sub TB_POR_Change()
If TB_POR.Text = "" Then
TB_POR.BackColor = &H8080FF
Else
TB_POR.BackColor = &H80000005
End If
End Sub
Private Sub TB_VTF_Change()
If TB_VTF.Text = "" Then
TB_VTF.BackColor = &H8080FF
Else
TB_VTF.BackColor = &H80000005
End If
End Sub
Private Sub TB_NEI_Change()
If TB_NEI.Text = "" Then
TB_NEI.BackColor = &H8080FF
Else
TB_NEI.BackColor = &H80000005
End If
End Sub
Private Sub TB_AUT3_Change()
If TB_AUT3.Text = "" Then
TB_AUT3.BackColor = &H8080FF
Else
TB_AUT3.BackColor = &H80000005
End If
End Sub
Private Sub TB_CDMR_Change()
If TB_CDMR.Text = "" Then
TB_CDMR.BackColor = &H8080FF
Else
TB_CDMR.BackColor = &H80000005
End If
End Sub
Private Sub TB_ICDM_Change()
If TB_ICDM.Text = "" Then
TB_ICDM.BackColor = &H8080FF
Else
TB_ICDM.BackColor = &H80000005
End If
End Sub

Private Sub VTF_PLUS_Click()
TB_VTF.Value = TB_VTF.Value + 1
End Sub
