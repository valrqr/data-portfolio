' Module : ADMIN.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub Userfrom_Activate()

TB_ADMIN.Value = ""
TB_VRU.Value = ""
TB_MAJ.Value = ""

TB_ADMIN.Text = ""
TB_VRU.Text = ""
TB_MAJ.Text = ""

End Sub

Private Sub CB_ADM_F_Click()

ADMIN.Hide

End Sub

Private Sub CB_ADMIN_Click()
If TB_ADMIN.Text = "MAFALDA" Then

Sheets("SAISIE").Unprotect

Sheets("LOG_M").Visible = True
Sheets("LOG_ST").Visible = True
Sheets("LOG_ANNUEL").Visible = True
Sheets("LOG_ANNUEL_A-1").Visible = True
Sheets("DATA_ST").Visible = True
Sheets("DATA_ST3").Visible = True
Sheets("DATA_MAIN").Visible = True
Sheets("DATA_BASE").Visible = True
Sheets("LOG_OPS").Visible = True
Sheets("SYNTH").Visible = True
Sheets("CR").Visible = True
Sheets("CR_PP").Visible = True
Sheets("CON+ADD").Visible = True

ADMIN.Hide

Else
MsgBox ("Le mot de passe saisie est erroné")
End If
End Sub

Private Sub CB_MAJ_Click()

If TB_MAJ.Text = "B737_CONTENEURISE" Then

MAJ_DB1.Show

ADMIN.Hide

Else
MsgBox ("Le mot de passe saisie est erroné")
End If

End Sub
Private Sub CB_VRU_Click()

If TB_VRU.Text = "MAFALDA" Then

Sheets("SAISIE").Protect

Sheets("LOG_M").Visible = xlSheetVeryHidden
Sheets("LOG_ST").Visible = xlSheetVeryHidden
Sheets("LOG_ANNUEL").Visible = xlSheetVeryHidden
Sheets("LOG_ANNUEL_A-1").Visible = xlSheetVeryHidden
Sheets("DATA_ST").Visible = xlSheetVeryHidden
Sheets("DATA_ST3").Visible = xlSheetVeryHidden
Sheets("DATA_MAIN").Visible = xlSheetVeryHidden
Sheets("DATA_BASE").Visible = xlSheetVeryHidden
Sheets("LOG_OPS").Visible = xlSheetVeryHidden
Sheets("SYNTH").Visible = xlSheetVeryHidden
Sheets("CR").Visible = xlSheetVeryHidden
Sheets("CR_PP").Visible = xlSheetVeryHidden
Sheets("CON+ADD").Visible = xlSheetVeryHidden


ADMIN.Hide

Else

MsgBox ("Le mot de passe saisie est erroné")

End If

End Sub
