' Module : SAISIE_J.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_DEV_Click()

Range("DV_CS1").Value = ""
Range("DV_CS2").Value = ""
Range("DV_CS3").Value = ""
Range("DV_CS4").Value = ""
Range("DV_CS5").Value = ""
Range("DV_CS6").Value = ""
Range("DV_CS7").Value = ""
Range("DV_CS8").Value = ""
Range("DV_CS9").Value = ""
Range("DV_CS10").Value = ""
Range("DV_CS11").Value = ""
Range("DV_CS12").Value = ""

DEV.Show
End Sub

Private Sub CB_DIV_LCH_Click()
DIV.Show
End Sub
Private Sub CB_EPU_Click()
ETATS.Show
End Sub

Private Sub CB_FRET_Click()

Range("FRET_ENT") = ""
Range("FRET_SOR") = ""
Range("FRET_M30") = ""
Range("FRET_P30") = ""
Range("FRET_1H") = ""

FRET.Show



End Sub

Private Sub CB_INTER_Click()
INTER.Show
End Sub

Private Sub CB_SJ_FACT_Click()

S_FAC.Show

End Sub

Private Sub CB_SJ_VAL_Click()

Range("ACT_MOD") = 1

If Range("JOUR_J").Value = 31 And Range("MOIS_M").Value = 12 Then

Sheets("SAISIE").Select
Range("N18").Select
Selection.Copy
Sheets("CR_PP").Visible = True
    Sheets("CR_PP").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Sheets("CR_PP").Visible = xlSheetVeryHidden
Sheets("SAISIE").Select

Else

Call MAIN_S

End If

SAISIE_J.Hide

ActiveWorkbook.Save

End Sub
Private Sub MAIN_S()

Dim ADD_SGS, ADD_IPG, ADD_CVH, ADD_FOD, ADD_FOA, ADD_FPA

ADD_SGS = Range("ADD_JJ_20").Value
ADD_IPG = Range("ADD_JJ_1").Value
ADD_CVH = Range("ADD_JJ_144").Value
ADD_FOD = Range("ADD_JJ_21").Value
ADD_FOA = Range("ADD_JJ_22").Value
ADD_FPA = Range("ADD_JJ_23").Value

Sheets("LOG_M").Range(ADD_SGS) = TB_SGS.Text
Sheets("LOG_M").Range(ADD_IPG) = TB_IPG.Text
Sheets("LOG_M").Range(ADD_CVH) = TB_CVH.Text
Sheets("LOG_M").Range(ADD_FOD) = TB_FOD.Text
Sheets("LOG_M").Range(ADD_FOA) = TB_FOA.Text
Sheets("LOG_M").Range(ADD_FPA) = TB_FPA.Text


End Sub
Private Sub CB_TRA_Click()

TRAF_2.Show

End Sub
Private Sub CONVOY_S_Click()

CONVOY.Show

End Sub

Private Sub CYV_PLUS_Click()

TB_CVH.Value = TB_CVH.Value + 1

End Sub

Private Sub FA_PLUS_Click()

TB_FOA.Value = TB_FOA.Value + 1

End Sub

Private Sub FNA_PLUS_Click()

TB_FOD.Value = TB_FOD.Value + 1

End Sub

Private Sub FPA_PLUS_Click()

TB_FPA.Value = TB_FPA.Value + 1

End Sub

Private Sub I_NPS_Click()

I_NP.Show

End Sub
Private Sub Saisie_ANN_Click()

SAISIE_J.Hide

End Sub
Private Sub SGS_PLUS_Click()

TB_SGS.Value = TB_SGS.Value + 1

End Sub
Private Sub TB_CVH_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TB_CVH.Text = 1
CYV_PLUS.Visible = True

End Sub
Private Sub TB_CVH_afterupdate()

If TB_CVH.Text = "" Then

CYV_PLUS.Visible = False

End If

End Sub
Private Sub TB_FOA_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TB_FOA.Text = 1
FA_PLUS.Visible = True

End Sub
Private Sub TB_FOA_afterupdate()

If TB_FOA.Text = "" Then

FA_PLUS.Visible = False

End If

End Sub
Private Sub TB_FOD_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TB_FOD.Text = 1
FNA_PLUS.Visible = True

End Sub
Private Sub TB_FOD_afterupdate()

If TB_FOD.Text = "" Then

FNA_PLUS.Visible = False

End If

End Sub
Private Sub TB_FPA_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TB_FPA.Text = 1
FPA_PLUS.Visible = True

End Sub
Private Sub TB_FPA_afterupdate()

If TB_FPA.Text = "" Then

FPA_PLUS.Visible = False

End If

End Sub
Private Sub TB_SGS_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

TB_SGS.Text = 1
SGS_PLUS.Visible = True

End Sub
Private Sub TB_SGS_afterupdate()

If TB_SGS.Text = "" Then

SGS_PLUS.Visible = False

End If

End Sub
Private Sub Userform_Activate()

If Range("ACT_MOD") = 1 Then

MsgBox ("L'ajout de données via cette interface post SAISIE JOURNALIERE effectuée n'est pas disponible, utiliser CONSULTATION DONNEES et modifier les données déjà saisies du jour")

SAISIE_J.Hide
CONSUL_MAIN.Show

End If

SGS_PLUS.Visible = False
CYV_PLUS.Visible = False
FNA_PLUS.Visible = False
FA_PLUS.Visible = False
FPA_PLUS.Visible = False

TB_SGS.Text = ""
TB_FOD.Text = ""
TB_FOA.Text = ""
TB_FPA.Text = ""
TB_CVH.Text = ""

Saisie_ANN.Caption = "ANNULER"

SGS_PLUS.Visible = False
CYV_PLUS.Visible = False
FNA_PLUS.Visible = False
FA_PLUS.Visible = False
FPA_PLUS.Visible = False



End Sub
Private Sub SAISIE_MAA_Click()
MAA_1.Show
End Sub
