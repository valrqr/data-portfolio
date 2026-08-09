' Module : SURTOU_1.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_AUJ_Click()

If CB_AUJ.Value = True Then

CB_ALSDAT.Value = False

Range("ST_ALT_DATE").Value = 0

Else

CB_ALSDAT.Value = True

End If

End Sub

Private Sub CB_COP_ADD_Click()
COP.Show
End Sub

Private Sub CB_COP_Change()
Range("ST_COP").Value = CB_COP.ListIndex + 1
End Sub

Private Sub CommandButton1_Click()
Dim ALT_DATE, ALT_ENT
ALT_ENT = TB_SDAT.Value
ALT_DATE = Format(ALT_ENT, "dd/mm/yyyy")

Sheets("LOG_ST").Range("H6").End(xlDown).Offset(1, 0) = ALT_DATE

End Sub

Private Sub TB_SDAT_afterupdate()
Range("ST_ENT_DATE").Value = TB_SDAT.Value
End Sub

Private Sub Userform_Activate()

CB_AUJ.Value = False
CB_ALSDAT.Value = False

CB_SCIES.Value = ""
CB_ST_APP.Value = ""
CB_COP.Value = ""

TB_SDAT.Text = ""
TB_TRIC.Text = ""
TB_IMM.Text = ""
TB_PKG.Text = ""
TB_TRIA.Text = ""
TB_ACOP.Text = ""
TB_BON.Text = ""
TB_BOFF.Text = ""

Range("ST_LI_CIE").Value = ""
Range("ST_LI_APP").Value = ""
Range("ST_EY_TCIE").Value = ""
Range("ST_EY_TAPP").Value = ""
Range("ST_EY_IMM").Value = ""
Range("ST_ASS_CR").Value = ""
Range("ST_PKG").Value = ""
Range("ST_COP").Value = ""

End Sub
Private Sub CB_ALSDAT_Click()

If CB_ALSDAT.Value = True Then

Range("ST_ALT_DATE").Value = 1

CB_AUJ.Value = False
LB_SDAT.Visible = True
TB_SDAT.Visible = True

Else

Range("ST_ALT_DATE").Value = 0

CB_AUJ.Value = True
LB_SDAT.Visible = False
TB_SDAT.Visible = False


End If
End Sub
Private Sub CB_ANS1_Click()
SURTOU_1.Hide
End Sub
Private Sub CB_ASSC_Click()
If CB_ASSC.Value = True Then
CB_CASS.Visible = True
Else
CB_CASS.Visible = False
Range("ST_ASS_CR").Value = ""
End If
End Sub
Private Sub CB_CASS_Change()

Range("ST_ASS_CR").Value = CB_CASS.ListIndex + 1

Dim ADD_ASS_CR

ADD_ASS_CR = Range("ADD_ASS_CR").Value

Sheets("DATA_ST").Range(ADD_ASS_CR) = Range("ST_HDG_3").Value

End Sub
Private Sub CB_SCIES_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub
Private Sub CB_SCIES_AfterUpdate()
Sheets("DATA_MAIN").Range("ST_LI_CIE").Value = CB_SCIES.ListIndex
Sheets("DATA_ST").Range("ST_LI_CIE3").Value = CB_SCIES.ListIndex
TB_TRIC.Text = Range("ST_MC_TCIE").Value
LB_SASS.Caption = Range("ST_HDG_2").Value

End Sub
Private Sub CB_ST_APP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub
Private Sub CB_ST_APP_AfterUpdate()
Sheets("DATA_MAIN").Range("ST_LI_APP").Value = CB_ST_APP.ListIndex
TB_TRIA.Text = Range("ST_MC_TAPP").Value
End Sub

Private Sub CB_SURS_Click()

If Range("ST_LI_CIE").Value = "" Then

MsgBox "Le champs Compagnie est manquant, saisir pour poursuivre"

ElseIf Range("ST_COP").Value = "" Then

MsgBox "Le champs COP est manquant, saisir pour poursuivre"

ElseIf CB_AUJ.Value = False And CB_ALSDAT.Value = False Then

MsgBox "Saisir une date pour poursuivre"

ElseIf TB_IMM.Text = "" Then

MsgBox "Le champs Immatriculation est manquant, saisir pour poursuivre"

ElseIf TB_PKG.Text = "" Then

MsgBox "Le champs Parking est manquant, saisir pour poursuivre"

ElseIf Range("ST_LI_APP").Value = "" Then

MsgBox "Le champs Appareil est manquant, saisir pour poursuivre"

ElseIf TB_ACOP.Text = "" Then

MsgBox "Le champs Arrivée COP est manquant, saisir pour poursuivre"

ElseIf TB_BON.Text = "" Then

MsgBox "Le champs ON BT est manquant, saisir pour poursuivre"

ElseIf TB_BOFF.Text = "" Then

MsgBox "Le champs OFF BT est manquant, saisir pour poursuivre"

Else

Sheets("LOG_ST").Range("G6").End(xlDown).Offset(1, 0) = Range("ST_HDG_4").Value
Sheets("LOG_ST").Range("K6").End(xlDown).Offset(1, 0) = Range("ST_RLI_CIE").Value
Sheets("LOG_ST").Range("L6").End(xlDown).Offset(1, 0) = Range("ST_RLI_APP").Value
Sheets("LOG_ST").Range("M6").End(xlDown).Offset(1, 0) = TB_IMM.Text
Sheets("LOG_ST").Range("N6").End(xlDown).Offset(1, 0) = Range("ST_PKG").Value
Sheets("LOG_ST").Range("O6").End(xlDown).Offset(1, 0) = TB_BON.Text
Sheets("LOG_ST").Range("P6").End(xlDown).Offset(1, 0) = TB_BOFF.Text
Sheets("LOG_ST").Range("I6").End(xlDown).Offset(1, 0) = Range("ST_COP").Value
Sheets("LOG_ST").Range("J6").End(xlDown).Offset(1, 0) = TB_ACOP.Text

If Range("ST_ALT_DATE").Value = 1 Then

Sheets("LOG_ST").Range("H6").End(xlDown).Offset(1, 0) = Range("ST_DATE_ALT").Value

ElseIf Range("ST_ALT_DATE").Value = 0 Then

Sheets("LOG_ST").Range("H6").End(xlDown).Offset(1, 0) = Range("DATE_G").Value

End If

SURTOU_2.Show

End If
End Sub
Private Sub TB_PKG_Change()
Range("ST_PKG").Value = TB_PKG.Text
End Sub
Private Sub TB_PKG_AfterUpdate()
If Range("ST_PKG_CT") = 1 Then
MsgBox "Ce parking n'existe pas ou n'est pas encore répertorié, recommencer."
End If
End Sub

Private Sub TB_TRIA_AfterUpdate()

Range("ST_EY_TAPP").Value = TB_TRIA.Text

If Range("ST_EY_APP_CT").Value = "1" Then
MsgBox "Ce tigramme ne correspond à aucun appareil, recommencer."
Else
CB_ST_APP.Text = Range("ST_EY_MC_APP").Value
End If

End Sub
Private Sub TB_TRIC_AfterUpdate()

Range("ST_EY_TCIE").Value = TB_TRIC.Text

If Range("ST_EY_CIE_CT").Value = "1" Then
MsgBox "Ce tigramme ne correspond à aucune compagnie, recommencer."
Else
LB_SASS.Caption = Range("ST_HDG_2").Value
End If

End Sub
