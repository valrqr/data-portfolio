' Module : TRAF_2.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub Userform_Activate()

TB_TRAF.Text = 0
CB_TRAF_MAIN.Text = ""
CB_TRAF_ASS.Text = ""

End Sub
Private Sub CB_NSAISIE_TRAF_Click()

If TB_TRAF.Text = 0 Or CB_TRAF_MAIN.Text = "" Or CB_TRAF_ASS.Text = "" Then

MsgBox "Complêter tout les champs pour poursuivre"

Else

Dim ADD_TRAF, ADD_TASS

ADD_TRAF = Range("ADD_JJ_TMAIN").Value
ADD_TASS = Range("ADD_JJ_TMAFR").Value

Sheets("LOG_M").Range(ADD_TRAF) = TB_TRAF.Text
Sheets("LOG_M").Range(ADD_TASS) = TB_TRAF.Text

End If

TB_TRAF.Text = 0
CB_TRAF_MAIN.Text = ""
CB_TRAF_ASS.Text = ""

End Sub
End Sub
Private Sub CB_TRA2_VAL_Click()

If TB_TRAF.Text = 0 Or CB_TRAF_MAIN.Text = "" Or CB_TRAF_ASS.Text = "" Then

MsgBox "Complêter tout les champs pour poursuivre"

Else

Dim ADD_TRAF, ADD_TASS

ADD_TRAF = Range("ADD_JJ_TMAIN").Value
ADD_TASS = Range("ADD_JJ_TMAFR").Value

Sheets("LOG_M").Range(ADD_TRAF) = TB_TRAF.Text
Sheets("LOG_M").Range(ADD_TASS) = TB_TRAF.Text

End If

TRAF_2.Hide

End Sub
Private Sub CB_TRAF_ASS_Change()
Range("ADD_TRAF_HDG").Value = CB_TRAF_ASS.ListIndex + 1
End Sub
Private Sub CB_TRAF_MAIN_Change()
Range("ADD_TRA_MAIN").Value = CB_TRAF_MAIN.ListIndex + 1
End Sub
Private Sub CB_TRA2_CNX_Click()
TRAF_2.Hide
End Sub
Private Sub TRA2FAC_Click()
S_FAC.Show
End Sub
Private Sub TRAF_PLUS_Click()
TB_TRAF.Value = TB_TRAF.Value + 1
End Sub
