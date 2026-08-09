' Module : DIV.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub BDD_PLUS_Click()

TB_BDD.Value = TB_BDD.Value + 1

End Sub


Private Sub CB_DIV_AN2_Click()

TB_MDG.BackColor = &H80000005
CB_MDG.Value = False
TB_MDG.Visible = False
TB_MDG.Text = ""

TB_BDD.BackColor = &H80000005
CB_BDD.Value = False
TB_BDD.Visible = False
TB_BDD.Text = ""

DIV.Hide

End Sub

Private Sub CB_DIV_VAL2_Click()

If CB_MDG.Value = True And TB_MDG.Text = "" Then

MsgBox "Le nombre de Mesure de Glissance est manquant"
TB_MDG.BackColor = &H8080FF

ElseIf CB_BDD.Value = True And TB_BDD.Text = "" Then

MsgBox "Le nombre de Balise de Détresse est manquant"
TB_BDD.BackColor = &H8080FF

Else

Dim ADD_MDG, ADD_BDD

ADD_MDG = Range("ADD_JJ_25").Value
ADD_BDD = Range("ADD_JJ_24").Value


Sheets("LOG_M").Range(ADD_MDG) = TB_MDG.Text
Sheets("LOG_M").Range(ADD_BDD) = TB_BDD.Text

End If

DIV.Hide

End Sub

Private Sub CB_MDG_Click()

If CB_MDG.Value = True Then

TB_MDG.Visible = True
MDG_PLUS.Visible = True
TB_MDG.Text = "1"

Else

TB_MDG.Text = ""
TB_MDG.Visible = False
MDG_PLUS.Visible = False
TB_MDG.BackColor = &H80000005

End If
End Sub
Private Sub CB_BDD_Click()

If CB_BDD.Value = True Then

TB_BDD.Visible = True
BDD_PLUS.Visible = True
TB_BDD.Text = "1"

Else

TB_BDD.Text = ""
TB_BDD.Visible = False
BDD_PLUS.Visible = False
TB_BDD.BackColor = &H80000005

End If
End Sub

Private Sub MDG_PLUS_Click()

TB_MDG.Value = TB_MDG.Value + 1

End Sub


Private Sub TB_BDD_Change()
If TB_BDD.Text = "" Then
TB_BDD.BackColor = &H8080FF
Else
TB_BDD.BackColor = &H80000005
End If
End Sub

Private Sub TB_MDG_Change()
If TB_MDG.Text = "" Then
TB_MDG.BackColor = &H8080FF
Else
TB_MDG.BackColor = &H80000005
End If
End Sub
