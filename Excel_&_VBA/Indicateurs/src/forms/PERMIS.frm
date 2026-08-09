' Module : PERMIS.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_CNX_PER_Click()
PERMIS.Hide
End Sub

Private Sub CB_M_PERMIS_Change()

Range("MOIS_X") = CB_M_PERMIS.ListIndex + 1

TB_FOT.Text = Range("DMX_174")
TB_EXT.Text = Range("DMX_173")
TB_FIM.Text = Range("DMX_175")
TB_FRM.Text = Range("DMX_176")
TB_FOC.Text = Range("DMX_177")
TB_PTD.Text = Range("DMX_179")
TB_PPNA.Text = Range("DMX_181")
TB_PCD.Text = Range("DMX_180")
TB_PTRD.Text = Range("DMX_182")

TB_AUT.Text = Range("DMX_162")
TB_AUM.Text = Range("DMX_163")
TB_AUC.Text = Range("DMX_164")
TB_AUP.Text = Range("DMX_165")
TB_FFT.Text = Range("DMX_166")
TB_REDT.Text = Range("DMX_167")
TB_REDM.Text = Range("DMX_168")
TB_REDC.Text = Range("DMX_169")
TB_REDP.Text = Range("DMX_170")
TB_RFFT.Text = Range("DMX_171")
TB_TH.Text = Range("DMX_172")


End Sub
Private Sub CB_VAL_PERM_Click()

Dim ADD_FOT, ADD_PTRC, ADD_EXT, ADD_FIM, ADD_FRM, ADD_FOC, ADD_PTD, ADD_PNAD, ADD_PCD, ADD_AUT, ADD_AUM, ADD_AUP, ADD_AUC, ADD_FFT, ADD_REDM, ADD_REDT, ADD_REDP, ADD_REDC, ADD_RFFT, ADD_TH, ADD_162, ADD_163, ADD_164, ADD_165, ADD_166, ADD_167, ADD_168, ADD_169, ADD_170, ADD_171, ADD_172, ADD_173, ADD_174, ADD_175, ADD_176, ADD_177, ADD_178, ADD_179, ADD_180, ADD_181, ADD_181, ADD_182

ADD_FOT = Range("ADD_MX_174").Value
ADD_EXT = Range("ADD_MX_173").Value
ADD_FIM = Range("ADD_MX_175").Value
ADD_FRM = Range("ADD_MX_176").Value
ADD_FOC = Range("ADD_MX_177").Value
ADD_PTD = Range("ADD_MX_179").Value
ADD_PMD = Range("ADD_MX_178").Value
ADD_PNAD = Range("ADD_MX_181").Value
ADD_PCD = Range("ADD_MX_180").Value
ADD_PTRC = Range("ADD_MX_182").Value

ADD_AUT = Range("ADD_MX_162").Value
ADD_AUM = Range("ADD_MX_163").Value
ADD_AUC = Range("ADD_MX_164").Value
ADD_AUP = Range("ADD_MX_165").Value
ADD_FFT = Range("ADD_MX_166").Value

ADD_REDM = Range("ADD_MX_167").Value
ADD_REDT = Range("ADD_MX_168").Value
ADD_REDC = Range("ADD_MX_169").Value
ADD_REDP = Range("ADD_MX_170").Value
ADD_RFFT = Range("ADD_MX_171").Value

ADD_TH = Range("ADD_MX_172").Value

Sheets("LOG_ANNUEL").Range(ADD_FOT) = TB_FOT.Text
Sheets("LOG_ANNUEL").Range(ADD_EXT) = TB_EXT.Text
Sheets("LOG_ANNUEL").Range(ADD_FIM) = TB_FIM.Text
Sheets("LOG_ANNUEL").Range(ADD_FRM) = TB_FRM.Text
Sheets("LOG_ANNUEL").Range(ADD_FOC) = TB_FOC.Text
Sheets("LOG_ANNUEL").Range(ADD_PTD) = TB_PTD.Text
Sheets("LOG_ANNUEL").Range(ADD_PMD) = TB_AUM.Text
Sheets("LOG_ANNUEL").Range(ADD_PCD) = TB_PCD.Text
Sheets("LOG_ANNUEL").Range(ADD_PNAD) = TB_PPNA.Text
Sheets("LOG_ANNUEL").Range(ADD_PTRC) = TB_PTRD.Text

Sheets("LOG_ANNUEL").Range(ADD_AUT) = TB_AUT.Text
Sheets("LOG_ANNUEL").Range(ADD_AUM) = TB_AUM.Text
Sheets("LOG_ANNUEL").Range(ADD_AUC) = TB_AUC.Text
Sheets("LOG_ANNUEL").Range(ADD_AUP) = TB_AUP.Text
Sheets("LOG_ANNUEL").Range(ADD_FFT) = TB_FFT.Text

Sheets("LOG_ANNUEL").Range(ADD_REDT) = TB_REDT.Text
Sheets("LOG_ANNUEL").Range(ADD_REDM) = TB_REDM.Text
Sheets("LOG_ANNUEL").Range(ADD_REDC) = TB_REDC.Text
Sheets("LOG_ANNUEL").Range(ADD_REDP) = TB_REDP.Text
Sheets("LOG_ANNUEL").Range(ADD_RFFT) = TB_RFFT.Text

Sheets("LOG_ANNUEL").Range(ADD_TH) = TB_TH.Text


Sheets("LOG_M (2)").Range(ADD_174) = TB_FOT.Text
Sheets("LOG_M (2)").Range(ADD_173) = TB_EXT.Text
Sheets("LOG_M (2)").Range(ADD_175) = TB_FIM.Text
Sheets("LOG_M (2)").Range(ADD_176) = TB_FRM.Text
Sheets("LOG_M (2)").Range(ADD_177) = TB_FOC.Text
Sheets("LOG_M (2)").Range(ADD_179) = TB_PTD.Text
Sheets("LOG_M (2)").Range(ADD_178) = TB_AUM.Text
Sheets("LOG_M (2)").Range(ADD_181) = TB_PCD.Text
Sheets("LOG_M (2)").Range(ADD_180) = TB_PPNA.Text
Sheets("LOG_M (2)").Range(ADD_182) = TB_PTRD.Text

Sheets("LOG_M (2)").Range(ADD_162) = TB_AUT.Text
Sheets("LOG_M (2)").Range(ADD_163) = TB_AUM.Text
Sheets("LOG_M (2)").Range(ADD_164) = TB_AUC.Text
Sheets("LOG_M (2)").Range(ADD_165) = TB_AUP.Text
Sheets("LOG_M (2)").Range(ADD_166) = TB_FFT.Text

Sheets("LOG_M (2)").Range(ADD_167) = TB_REDT.Text
Sheets("LOG_M (2)").Range(ADD_168) = TB_REDM.Text
Sheets("LOG_M (2)").Range(ADD_169) = TB_REDC.Text
Sheets("LOG_M (2)").Range(ADD_170) = TB_REDP.Text
Sheets("LOG_M (2)").Range(ADD_171) = TB_RFFT.Text

Sheets("LOG_M (2)").Range(ADD_172) = TB_TH.Text



Range("CR_PERT") = TB_AUT.Text
Range("CR_PERM") = TB_AUM.Text
Range("CR_PERC") = TB_AUC.Text
Range("CR_PERP") = TB_AUP.Text
Range("CR_PERH") = TB_TH.Text

PERMIS.Hide

Range("ETS_PER") = 1

ActiveWorkbook.Save

End Sub
