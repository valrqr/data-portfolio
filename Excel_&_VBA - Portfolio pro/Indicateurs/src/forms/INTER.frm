' Module : INTER.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub AUT2_PLUS_Click()

TB_AUT2.Value = TB_AUT2.Value + 1

End Sub

Private Sub BAL_PLUS_Click()

TB_BAL.Value = TB_BAL.Value + 1

End Sub

Private Sub INF_PLUS_Click()

TB_INF.Value = TB_INF.Value + 1

End Sub
Private Sub Userform_Activate()

TB_INF.BackColor = &H80000005
CB_INF.Value = False
TB_INF.Visible = False
TB_INF.Text = ""

TB_BAL.BackColor = &H80000005
CB_BAL.Value = False
TB_BAL.Visible = False
TB_BAL.Text = ""

TB_AUT2.BackColor = &H80000005
CB_AUT2.Value = False
TB_AUT2.Visible = False
TB_AUT2.Text = ""

TB_APRE.BackColor = &H80000005
TB_APRE.Visible = False
TB_APRE.Text = ""

End Sub

Private Sub CB_INF_Click()

If CB_INF.Value = True Then

MsgBox ("Rappel!: Si un ou plusieurs déversement(s) ont eu lieu, les interventions INFRA correspondantes sont déjà comptabilisées, inutile de les re-saisir!")

TB_INF.Visible = True
INF_PLUS.Visible = True
TB_INF.Text = "1"

Else

TB_INF.Text = ""
TB_INF.Visible = False
INF_PLUS.Visible = False
TB_INF.BackColor = &H80000005

End If

End Sub
Private Sub CB_BAL_Click()

If CB_BAL.Value = True Then

TB_BAL.Visible = True
BAL_PLUS.Visible = True
TB_BAL.Text = "1"

Else

TB_BAL.Text = ""
TB_BAL.Visible = False
BAL_PLUS.Visible = False
TB_BAL.BackColor = &H80000005

End If
End Sub
Private Sub CB_AUT2_Click()

If CB_AUT2.Value = True Then

TB_AUT2.Visible = True
AUT2_PLUS.Visible = True
TB_AUT2.Text = "1"
TB_APRE.Visible = True
LB_APRE.Visible = True
TB_APRE.Text = ""

Else

TB_AUT2.Text = ""
TB_AUT2.Visible = False
AUT2_PLUS.Visible = False
TB_AUT2.BackColor = &H80000005
TB_APRE.Visible = False
LB_APRE.Visible = False

End If
End Sub
Private Sub CB_INT_CNX_Click()

TB_INF.BackColor = &H80000005
CB_INF.Value = False
TB_INF.Visible = False
TB_INF.Text = ""

TB_BAL.BackColor = &H80000005
CB_BAL.Value = False
TB_BAL.Visible = False
TB_BAL.Text = ""

TB_AUT2.BackColor = &H80000005
CB_AUT2.Value = False
TB_AUT2.Visible = False
TB_AUT2.Text = ""

TB_APRE.BackColor = &H80000005
TB_APRE.Visible = False
TB_APRE.Text = ""

INTER.Hide

End Sub
Private Sub CB_INT_VAL_Click()

If CB_INF.Value = True And TB_INF.Text = "" Then

MsgBox "Le nombre d'intervention INFRA est manquant"
TB_INF.BackColor = &H8080FF

ElseIf CB_BAL.Value = True And TB_BAL.Text = "" Then

MsgBox "Le nombre d'intervention BAL est manquant"
TB_BAL.BackColor = &H8080FF

ElseIf CB_AUT2.Value = True And TB_AUT2.Text = "" Then

MsgBox "Le nombre d'intervention AUTRE est manquant"
TB_AUT2.BackColor = &H8080FF

ElseIf CB_AUT2.Value = True And TB_APRE.Text = "" Then

MsgBox "Préciser le service concerné par AUTRE pour poursuivre."
TB_APRE.BackColor = &H8080FF

Else

Dim ADD_INF, ADD_BAL, ADD_AUT2, ADD_APRE

ADD_INF = Range("ADD_JJ_13").Value
ADD_BAL = Range("ADD_JJ_12").Value
ADD_AUT2 = Range("ADD_JJ_14").Value
ADD_APRE = Range("ADD_JJ_15").Value

Sheets("LOG_M").Range(ADD_INF) = TB_INF.Text
Sheets("LOG_M").Range(ADD_BAL) = TB_BAL.Text
Sheets("LOG_M").Range(ADD_AUT2) = TB_AUT2.Text
Sheets("LOG_M").Range(ADD_APRE) = TB_APRE.Text


End If

INTER.Hide

End Sub
Private Sub TB_APRE_Change()
If TB_APRE.Text = "" Then
TB_APRE.BackColor = &H8080FF
Else
TB_APRE.BackColor = &H80000005
End If
End Sub
Private Sub TB_AUT2_Change()

If TB_AUT2.Value <> "" Then
TB_APRE.Visible = True
LB_APRE.Visible = True
Else
TB_APRE.Visible = False
LB_APRE.Visible = False
End If

If TB_AUT2.Text = "" Then
TB_AUT2.BackColor = &H8080FF
Else
TB_AUT2.BackColor = &H80000005
End If

End Sub
Private Sub TB_BAL_Change()
If TB_BAL.Text = "" Then
TB_BAL.BackColor = &H8080FF
Else
TB_BAL.BackColor = &H80000005
End If
End Sub

Private Sub TB_INF_Change()
If TB_INF.Text = "" Then
TB_INF.BackColor = &H8080FF
Else
TB_INF.BackColor = &H80000005
End If
End Sub
