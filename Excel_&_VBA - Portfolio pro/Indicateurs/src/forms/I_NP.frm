' Module : I_NP.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub Userform_Activate()

Range("JOUR_X") = Range("JOUR_J")

TB_PLUS.Text = 1
Range("DATA_VAL_MAIN").Value = 1
TGB_MS.Visible = False

CB_SB.Value = False
TB_SBS.Visible = False
TB_SBS.Text = ""

CB_BA.Value = False
TB_BSA.Visible = False
TB_BSA.Text = ""

CB_AA.Value = False
TB_ACA.Visible = False
TB_ACA.Text = ""

CB_EV.Value = False
TB_ETV.Visible = False
TB_ETV.Text = ""

CB_EA.Value = False
TB_ETA.Visible = False
TB_ETA.Text = ""

CB_EC.Value = False
TB_ACC.Visible = False
TB_ACC.Text = ""

CB_IN.Value = False
TB_IIB.Visible = False
TB_IIB.Text = ""

CB_TV.Value = False
TB_TVX.Visible = False
TB_TVX.Text = ""

CB_DI.Value = False
TB_AUT.Visible = False
TB_AUT.Text = ""

Range("INP_SB").Value = ""
Range("INP_BA").Value = ""
Range("INP_AA").Value = ""
Range("INP_EV").Value = ""
Range("INP_EA").Value = ""
Range("INP_EC").Value = ""
Range("INP_IN").Value = ""
Range("INP_TV").Value = ""
Range("INP_DI").Value = ""

Range("INP_MSB").Value = ""
Range("INP_MBA").Value = ""
Range("INP_MAA").Value = ""
Range("INP_MEV").Value = ""
Range("INP_MEA").Value = ""
Range("INP_MEC").Value = ""
Range("INP_MIN").Value = ""
Range("INP_MTV").Value = ""
Range("INP_MDI").Value = ""

End Sub
Private Sub CB_INP_CNX_Click()

TB_SBS.BackColor = &H80000005
CB_SB.Value = False
TB_SBS.Visible = False
TB_SBS.Text = ""

TB_BSA.BackColor = &H80000005
CB_BA.Value = False
TB_BSA.Visible = False
TB_BSA.Text = ""

TB_ACA.BackColor = &H80000005
CB_AA.Value = False
TB_ACA.Visible = False
TB_ACA.Text = ""

TB_ETV.BackColor = &H80000005
CB_EV.Value = False
TB_ETV.Visible = False
TB_ETV.Text = ""

TB_ETA.BackColor = &H80000005
CB_EA.Value = False
TB_ETA.Visible = False
TB_ETA.Text = ""

TB_ACC.BackColor = &H80000005
CB_EC.Value = False
TB_ACC.Visible = False
TB_ACC.Text = ""

TB_IIB.BackColor = &H80000005
CB_IN.Value = False
TB_IIB.Visible = False
TB_IIB.Text = ""

TB_TVX.BackColor = &H80000005
CB_TV.Value = False
TB_TVX.Visible = False
TB_TVX.Text = ""

TB_AUT.BackColor = &H80000005
CB_DI.Value = False
TB_AUT.Visible = False
TB_AUT.Text = ""

Call INP_RESET

I_NP.Hide

End Sub
Private Sub CB_INP_VAL_Click()

If Range("INP_CNT").Value = 1 Then

MsgBox "Le total d'inspections NP saisi diffère du total des types d'inspections NP"

ElseIf CB_SB.Value = True And TB_SBS.Text = "" Then

MsgBox "Le nombre d'INP Suspicion Bird Strike est manquant"
TB_SBS.BackColor = &H8080FF

ElseIf CB_BA.Value = True And TB_BSA.Text = "" Then

MsgBox "Le nombre d'INP Bird Strike Avéré est manquant"
TB_BSA.BackColor = &H8080FF

ElseIf CB_AA.Value = True And TB_ACA.Text = "" Then

MsgBox "Le nombre d'INP Accélération / Arrêt est manquant"
TB_ACA3.BackColor = &H8080FF

ElseIf CB_EV.Value = True And TB_ETV.Text = "" Then

MsgBox "Le nombre d'INP Etat de Veille est manquant"
TB_ETV.BackColor = &H8080FF

ElseIf CB_EA.Value = True And TB_ETA.Text = "" Then

MsgBox "Le nombre d'INP Etat d'Alerte est manquant"
TB_ETA.BackColor = &H8080FF

ElseIf CB_EC.Value = True And TB_ACC.Text = "" Then

MsgBox "Le nombre d'INP Etat d'Accident est manquant"
TB_ACC.BackColor = &H8080FF

ElseIf CB_IN.Value = True And TB_IIB.Text = "" Then

MsgBox "Le nombre d'INP Suite Intervention INFRA/BAL est manquant"
TB_IIB.BackColor = &H8080FF

ElseIf CB_TV.Value = True And TB_TVX.Text = "" Then

MsgBox "Le nombre d'INP Suite Travaux est manquant"
TB_TVX.BackColor = &H8080FF

ElseIf CB_DI.Value = True And TB_AUT.Text = "" Then

MsgBox "Le nombre d'INP Etat de Veille est manquant"
TB_AUT.BackColor = &H8080FF

Else

Call INP_JJ_VAL

End If

I_NP.Hide

End Sub

Private Sub CB_SB_Click()
'Si l'on a cliqué sur le bouton CB_SB le textbox correspondant s'affiche
If CB_SB.Value = True Then
'La cellule correspondant à la valeur du bouton SB, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_SB") = "1"
TB_SBS.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton SB est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_SBS.Visible = False
TB_SBS.BackColor = &H80000005
Range("INP_SB") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_BA_Click()
'Si l'on a cliqué sur le bouton CB_BA le textbox correspondant s'affiche
If CB_BA.Value = True Then
'La cellule correspondant à la valeur du bouton BA, 1 ou 0, est activé en feuille DATA_MAINet la textbox correspondante est rendue visible
Range("INP_BA") = "1"
TB_BSA.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton BA est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_BSA.Visible = False
TB_BSA.BackColor = &H80000005
Range("INP_BA") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_AA_Click()
'Si l'on a cliqué sur le bouton CB_AA le textbox correspondant s'affiche
If CB_AA.Value = True Then
'La cellule correspondant à la valeur du bouton AA, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_AA") = "1"
TB_ACA.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton AA est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_ACA.Visible = False
TB_ACA.BackColor = &H80000005
Range("INP_AA") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_EV_Click()
'Si l'on a cliqué sur le bouton CB_EV le textbox correspondant s'affiche
If CB_EV.Value = True Then
'La cellule correspondant à la valeur du bouton EV, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_EV") = "1"
TB_ETV.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton EV est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_ETV.Visible = False
TB_ETV.BackColor = &H80000005
Range("INP_EV") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_EA_Click()
'Si l'on a cliqué sur le bouton CB_EA le textbox correspondant s'affiche
If CB_EA.Value = True Then
'La cellule correspondant à la valeur du bouton EA, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_EA") = "1"
TB_ETA.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton EA est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_ETA.Visible = False
TB_ETA.BackColor = &H80000005
Range("INP_EA") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_EC_Click()
'Si l'on a cliqué sur le bouton CB_EC le textbox correspondant s'affiche
If CB_EC.Value = True Then
'La cellule correspondant à la valeur du bouton EC, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_EC") = "1"
TB_ACC.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton EC est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_ACC.Visible = False
TB_ACC.BackColor = &H80000005
Range("INP_EC") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_IN_Click()
'Si l'on a cliqué sur le bouton CB_IN le textbox correspondant s'affiche
If CB_IN.Value = True Then
'La cellule correspondant à la valeur du bouton IN, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_IN") = "1"
TB_IIB.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton IN est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_IIB.Visible = False
TB_IIB.BackColor = &H80000005
Range("INP_IN") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_TV_Click()
'Si l'on a cliqué sur le bouton CB_TV le textbox correspondant s'affiche
If CB_TV.Value = True Then
'La cellule correspondant à la valeur du bouton TV, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_TV") = "1"
TB_TVX.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton TV est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_TVX.Visible = False
TB_TVX.BackColor = &H80000005
Range("INP_TV") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub
Private Sub CB_DI_Click()
'Si l'on a cliqué sur le bouton CB_DI le textbox correspondant s'affiche
If CB_DI.Value = True Then
'La cellule correspondant à la valeur du bouton DI, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("INP_DI") = "1"
TB_AUT.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call INP_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton DI est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_AUT.Visible = False
TB_AUT.BackColor = &H80000005
Range("INP_DI") = ""
'Les valeurs des textbox sont mises à jour via la procédure INP_TEXT_MAJ
Call INP_TEXT_MAJ
End If
End Sub

Private Sub INP_JJ_VAL()

Dim ADD_SB, ADD_BA, ADD_AA, ADD_EV, ADD_EA, ADD_EC, ADD_IN, ADD_TV, ADD_DI, ADD_EV2, ADD_EA2, ADD_EC2

ADD_SB = Range("AD_JJ_3").Value
ADD_BA = Range("AD_JJ_4").Value
ADD_AA = Range("AD_JJ_5").Value
ADD_EV = Range("AD_JJ_6").Value
ADD_EA = Range("AD_JJ_7").Value
ADD_EC = Range("AD_JJ_8").Value
ADD_IN = Range("AD_JJ_9").Value
ADD_TV = Range("AD_JJ_10").Value
ADD_DI = Range("AD_JJ_11").Value

ADD_EV2 = Range("ADD_JJ_16").Value
ADD_EA2 = Range("ADD_JJ_17").Value
ADD_EC2 = Range("ADD_JJ_18").Value

Sheets("LOG_M").Range(ADD_SB) = Range("INP_PSB").Value
Sheets("LOG_M").Range(ADD_BA) = Range("INP_PBA").Value
Sheets("LOG_M").Range(ADD_AA) = Range("INP_PAA").Value
Sheets("LOG_M").Range(ADD_EV) = Range("INP_PEV").Value
Sheets("LOG_M").Range(ADD_EA) = Range("INP_PEA").Value
Sheets("LOG_M").Range(ADD_EC) = Range("INP_PEC").Value
Sheets("LOG_M").Range(ADD_IN) = Range("INP_PIN").Value
Sheets("LOG_M").Range(ADD_TV) = Range("INP_PTV").Value
Sheets("LOG_M").Range(ADD_DI) = Range("INP_PDI").Value

Sheets("LOG_M").Range(ADD_EV2) = Range("INP_PEV").Value
Sheets("LOG_M").Range(ADD_EA2) = Range("INP_PEA").Value
Sheets("LOG_M").Range(ADD_EC2) = Range("INP_PEC").Value


Call INP_RESET

End Sub

Private Sub TB_ACA_AfterUpdate()

Range("INP_MAA").Value = TB_ACA.Text


If TB_ACA.Text = "" Then
TB_ACA.BackColor = &H8080FF
Else
TB_ACA.BackColor = &H80000005
End If

End Sub
Private Sub TB_ACC_AfterUpdate()

Range("INP_MEC").Value = TB_ACC.Text


If TB_ACC.Text = "" Then
TB_ACC.BackColor = &H8080FF
Else
TB_ACC.BackColor = &H80000005
End If

End Sub
Private Sub TB_AUT_AfterUpdate()

Range("INP_MDI").Value = TB_AUT.Text


If TB_AUT.Text = "" Then
TB_AUT.BackColor = &H8080FF
Else
TB_AUT.BackColor = &H80000005
End If

End Sub
Private Sub TB_BSA_AfterUpdate()

Range("INP_MBA").Value = TB_BSA.Text

If TB_BSA.Text = "" Then
TB_BSA.BackColor = &H8080FF
Else
TB_BSA.BackColor = &H80000005
End If

End Sub
Private Sub TB_ETA_AfterUpdate()

Range("INP_MEA").Value = TB_ETA.Text


If TB_ETA.Text = "" Then
TB_ETA.BackColor = &H8080FF
Else
TB_ETA.BackColor = &H80000005
End If

End Sub
Private Sub TB_ETV_AfterUpdate()

Range("INP_MEV").Value = TB_ETV.Text


If TB_ETV.Text = "" Then
TB_ETV.BackColor = &H8080FF
Else
TB_ETV.BackColor = &H80000005
End If

End Sub
Private Sub TB_IIB_AfterUpdate()

Range("INP_MIN").Value = TB_IIB.Text


If TB_IIB.Text = "" Then
TB_IIB.BackColor = &H8080FF
Else
TB_IIB.BackColor = &H80000005
End If

End Sub
Private Sub TB_PLUS_Change()

Range("DATA_VAL_MAIN").Value = TB_PLUS.Text

End Sub
Private Sub TB_SBS_AfterUpdate()

Range("INP_MSB").Value = TB_SBS.Text

If TB_SBS.Text = "" Then
TB_SBS.BackColor = &H8080FF
Else
TB_SBS.BackColor = &H80000005
End If

End Sub

Private Sub TB_TVX_AfterUpdate()

Range("INP_MTV").Value = TB_TVX.Text


If TB_TVX.Text = "" Then
TB_TVX.BackColor = &H8080FF
Else
TB_TVX.BackColor = &H80000005
End If

End Sub

Private Sub TGB_MS_Click()

If Range("DATA_VAL_MAIN").Value = 1 Then

TGB_MS.Visible = False

Else

Range("DATA_VAL_MAIN").Value = Range("DATA_VAL_MAIN").Value - 1
TB_PLUS.Text = Range("DATA_VAL_MAIN").Value

Call INP_TEXT_MAJ

End If

End Sub

Private Sub TGB_PL_Click()

Range("DATA_VAL_MAIN").Value = Range("DATA_VAL_MAIN").Value + 1
TB_PLUS.Text = Range("DATA_VAL_MAIN").Value

Call INP_TEXT_MAJ

TGB_MS.Visible = True

End Sub
Private Sub INP_TEXT_MAJ()

TB_SBS.Text = Range("INP_ASB")
TB_BSA.Text = Range("INP_ABA")
TB_ACA.Text = Range("INP_AAA")
TB_ETA.Text = Range("INP_AEA")
TB_ETV.Text = Range("INP_AEV")
TB_ACC.Text = Range("INP_AEC")
TB_IIB.Text = Range("INP_AIN")
TB_TVX.Text = Range("INP_ATV")
TB_AUT.Text = Range("INP_ADI")

End Sub
Private Sub INP_RESET()

Range("INP_SB").Value = ""
Range("INP_BA").Value = ""
Range("INP_AA").Value = ""
Range("INP_EV").Value = ""
Range("INP_EA").Value = ""
Range("INP_EC").Value = ""
Range("INP_IN").Value = ""
Range("INP_TV").Value = ""
Range("INP_DI").Value = ""

Range("INP_MSB").Value = ""
Range("INP_MBA").Value = ""
Range("INP_MAA").Value = ""
Range("INP_MEV").Value = ""
Range("INP_MEA").Value = ""
Range("INP_MEC").Value = ""
Range("INP_MIN").Value = ""
Range("INP_MTV").Value = ""
Range("INP_MDI").Value = ""

Range("DATA_VAL_MAIN").Value = ""

End Sub
