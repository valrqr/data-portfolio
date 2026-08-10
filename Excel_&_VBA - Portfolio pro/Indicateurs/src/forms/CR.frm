' Module : CR.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_RAZ_Click()

TB_FMA1.Text = ""
TB_FMA2.Text = ""
TB_FMA3.Text = ""
TB_FMA4.Text = ""
TB_FMA5.Text = ""
TB_FMA6.Text = ""
TB_FMA7.Text = ""
TB_FMA8.Text = ""
TB_CLB.Text = ""

End Sub

Private Sub Userform_Activate()

TB_FMA1.Text = Range("FAITM1")
TB_FMA2.Text = Range("FAITM2")
TB_FMA3.Text = Range("FAITM3")
TB_FMA4.Text = Range("FAITM4")
TB_FMA5.Text = Range("FAITM5")
TB_FMA6.Text = Range("FAITM6")
TB_FMA7.Text = Range("FAITM7")
TB_FMA8.Text = Range("FAITM8")

If Range("MOIS_M").Value = 10 Or Range("MOIS_M").Value = 11 Or Range("MOIS_M").Value = 12 Or Range("MOIS_M").Value = 1 Or Range("MOIS_M").Value = 2 Or Range("MOIS_M").Value = 3 Then

TB_CLB.Text = Range("FAITM_CV").Text

Else

TB_CLB.Text = Range("CHAMPS_LIB")

End If


End Sub

Private Sub CB_CR_FERM_Click()

CR.Hide

End Sub

Private Sub CB_SAVE_FMA_Click()

If TB_FMA1.Text = "" Then

Range("FAITM1") = ""

Else

Range("FAITM1") = TB_FMA1.Text

End If

If TB_FMA2.Text = "" Then

Range("FAITM2") = ""

Else

Range("FAITM2") = TB_FMA2.Text

End If

If TB_FMA3.Text = "" Then

Range("FAITM3") = ""

Else

Range("FAITM3") = TB_FMA3.Text

End If

If TB_FMA4.Text = "" Then

Range("FAITM4") = ""

Else

Range("FAITM4") = TB_FMA4.Text

End If

If TB_FMA5.Text = "" Then

Range("FAITM5") = ""

Else

Range("FAITM5") = TB_FMA5.Text

End If

If TB_FMA6.Text = "" Then

Range("FAITM6") = ""

Else

Range("FAITM6") = TB_FMA6.Text

End If

If TB_FMA7.Text = "" Then

Range("FAITM7") = ""

Else

Range("FAITM7") = TB_FMA7.Text

End If

If TB_FMA8.Text = "" Then

Range("FAITM8") = ""

Else

Range("FAITM8") = TB_FMA8.Text

End If

If TB_CLB.Text = "" Then

Range("CHAMPS_LIB") = ""

Else

Range("CHAMPS_LIB") = TB_CLB.Text

End If

End Sub
