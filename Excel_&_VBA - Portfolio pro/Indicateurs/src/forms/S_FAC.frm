' Module : S_FAC.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_FACT_ASS_Change()

Range("ADD_FACT_HDG").Value = CB_FACT_ASS.ListIndex + 1

End Sub
Private Sub Userform_Activate()

TB_FAC.Text = ""
CB_FACT_MAIN.Text = ""
CB_FACT_ASS.Visible = False

End Sub
Private Sub CB_FAC_CNX_Click()

S_FAC.Hide

End Sub

Private Sub CB_FAC_VAL_Click()

If TB_FAC.Text = 0 Or CB_FACT_MAIN.Text = "" Or CB_FACT_ASS.Text = "" Then

MsgBox "Complêter tout les champs pour poursuivre"

ElseIf Range("ADD_FACT_MAIN").Value = 2 Or Range("ADD_FACT_MAIN").Value = 5 Then

Dim ADD_FACT, ADD_FASS

ADD_FACT = Range("ADD_JJ_FACT").Value
ADD_FASS = Range("ADD_JJ_FASS").Value

Sheets("LOG_M").Range(ADD_FACT) = TB_FAC.Text
Sheets("LOG_M").Range(ADD_FASS) = TB_FAC.Text

Else

Dim ADD_FACT1

ADD_FACT1 = Range("ADD_JJ_FACT").Value

Sheets("LOG_M").Range(ADD_FACT1) = TB_FACT.Text

End If

S_FAC.Hide

End Sub
Private Sub CB_FACT_MAIN_Change()

Range("ADD_FACT_MAIN").Value = CB_FACT_MAIN.ListIndex + 1

If Range("ADD_FACT_MAIN").Value = 2 Or Range("ADD_FACT_MAIN").Value = 3 Or Range("ADD_FACT_MAIN").Value = 5 Then

CB_FACT_ASS.Visible = True

ElseIf Range("ADD_FACT_MAIN").Value = 1 Or Range("ADD_FACT_MAIN").Value = 4 Or Range("ADD_FACT_MAIN").Value = 6 Then

CB_FACT_ASS.Visible = False

End If

End Sub
Private Sub CB_NSAISIE_FACT_Click()

If TB_FAC.Text = 0 Or CB_FACT_MAIN.Text = "" Then

MsgBox "Complêter tout les champs pour poursuivre"

Else

Dim ADD_FACT

ADD_FACT = Range("ADD_JJ_FACT").Value

Sheets("LOG_M").Range(ADD_FACT) = TB_FAC.Text

End If

CB_FACT_MAIN.Text = ""

End Sub

Private Sub FAC_PLUS_Click()

TB_FAC.Value = TB_FAC.Value + 1

End Sub

Private Sub TB_FAC_Change()

If TB_FAC.Text = "" Then
TB_FAC.BackColor = &H8080FF
FAC_PLUS.Visible = False
Else
FAC_PLUS.Visible = True
TB_FAC.BackColor = &H80000005
End If

End Sub
