' Module : PRIX_PRES.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub Userform_Activate()

TB_CGP.Text = Range("PRIX_CGP").Text
TB_PBSF.Text = Range("PRIX_PBS").Text
TB_FMSH.Text = Range("PRIX_MSH").Text
TB_SAT.Text = Range("PRIX_SAT").Text
TB_JET.Text = Range("PRIX_JET").Text
TB_FEMT.Text = Range("PRIX_EMT").Text
TB_TRACT.Text = Range("PRIX_TRACT").Text

TB_PERMI.Text = Range("PRIX_PERMI").Text
TB_PERMR.Text = Range("PRIX_PERMR").Text
TB_EXAMT.Text = Range("PRIX_EXAMT").Text
TB_PERC.Text = Range("PRIX_PERC").Text
TB_PNA.Text = Range("PRIX_PNA").Text
TB_FOT.Text = Range("PRIX_FOT").Text
TB_FOM.Text = Range("PRIX_FOM").Text

End Sub
Private Sub CB_VAL_PRIX_Click()

Range("PRIX_CGP").Value = TB_CGP.Text
Range("PRIX_PBS").Value = TB_PBSF.Text
Range("PRIX_MSH").Value = TB_FMSH.Text
Range("PRIX_SAT").Value = TB_SAT.Text
Range("PRIX_JET").Value = TB_JET.Text
Range("PRIX_EMT").Value = TB_FEMT.Text
Range("PRIX_TRACT").Value = TB_TRACT.Text

Range("PRIX_PERMI").Value = TB_PERMI.Text
Range("PRIX_PERMR").Value = TB_PERMR.Text
Range("PRIX_EXAMT").Value = TB_EXAMT.Text
Range("PRIX_PERC").Value = TB_PERC.Text
Range("PRIX_PNA").Value = TB_PNA.Text
Range("PRIX_FOT").Value = TB_FOT.Text
Range("PRIX_FOM").Value = TB_FOM.Text

PRIX_PRES.Hide

End Sub
Private Sub CB_CNX_PRIX_Click()

PRIX_PRES.Hide

End Sub
