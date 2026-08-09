' Module : CONVOY.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub SB_C_Change()

With SB_C
.Max = 99
.Min = 1
.SmallChange = 1
End With

TB_PLUS_C.Text = SB_C.Value
Range("DATA_VAL_MAIN").Value = SB_C.Value
C_TEXT_MAJ

End Sub
Private Sub C_TEXT_MAJ()

TB_ACF.Text = Range("C_AACF2")
TB_DK23.Text = Range("C_ADK232")
TB_DCT.Text = Range("C_ADCT2")
TB_PAN.Text = Range("C_APAN2")
TB_REP.Text = Range("C_AREP2")
TB_SKU.Text = Range("C_ASKU2")
TB_TVX.Text = Range("C_ATVX2")
TB_MDE.Text = Range("C_AMDE2")
TB_AG_C.Text = Range("C_AAG")
TB_AC_C.Text = Range("C_AAC")

End Sub
Private Sub C_TEXT_MAJ2()

TB_AG_C.Text = Range("C_AAG")
TB_AC_C.Text = Range("C_AAC")

End Sub

Private Sub TB_CAC_Click()
'Si l'on a cliqué sur le bouton TB_CAC le textbox correspondant s'affiche
If TB_CAC.Value = True Then
'La cellule correspondant à la valeur du bouton AC, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_AC") = "1"
TB_AC_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton AFR est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_AC_C.Visible = False
TB_AC_C.BackColor = &H80000005
Range("C_AC") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If
End Sub

Private Sub TB_CAG_Click()
'Si l'on a cliqué sur le bouton TB_CAG le textbox correspondant s'affiche
If TB_CAG.Value = True Then
'La cellule correspondant à la valeur du bouton AG, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_AG") = "1"
TB_AG_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton AFR est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_AG_C.Visible = False
TB_AG_C.BackColor = &H80000005
Range("C_AG") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If
End Sub

Private Sub Userform_Activate()

Range("JOUR_X") = Range("JOUR_J")

TB_PLUS_C.Text = 1
Range("DATA_VAL_MAIN").Value = 1

Call C_RESET

CB_ACF.Value = False
TB_ACF.Visible = False
TB_ACF.Text = ""

CB_DK23.Value = False
TB_DK23.Visible = False
TB_DK23.Text = ""

CB_DCT.Value = False
TB_DCT.Visible = False
TB_DCT.Text = ""

CB_PAN.Value = False
TB_PAN.Visible = False
TB_PAN.Text = ""

CB_REP.Value = False
TB_REP.Visible = False
TB_REP.Text = ""

CB_SKU.Value = False
TB_SKU.Visible = False
TB_SKU.Text = ""

CB_TVX.Value = False
TB_TVX.Visible = False
TB_TVX.Text = ""

CB_MDE.Value = False
TB_MDE.Visible = False
TB_MDE.Text = ""

TB_CAFR.Value = False
TB_AFR_C.Visible = False
TB_AFR_C.Text = ""

TB_CAPR.Value = False
TB_APR_C.Visible = False
TB_APR_C.Text = ""

TB_CAVP.Value = False
TB_AVP_C.Visible = False
TB_AVP_C.Text = ""

TB_CMZS.Value = False
TB_MZS_C.Visible = False
TB_MZS_C.Text = ""

TB_CGEH.Value = False
TB_GEH_C.Visible = False
TB_GEH_C.Text = ""

TB_CSFS.Value = False
TB_SFS_C.Visible = False
TB_SFS_C.Text = ""

TB_CSEA.Value = False
TB_SEA_C.Visible = False
TB_SEA_C.Text = ""

TB_CAC.Value = False
TB_AC_C.Visible = False
TB_AC_C.Text = ""

TB_CAG.Value = False
TB_AG_C.Visible = False
TB_AG_C.Text = ""

End Sub
Private Sub C_RESET()

Range("C_MAFR").Value = ""
Range("C_MAVP").Value = ""
Range("C_MMZS").Value = ""
Range("C_MGEH").Value = ""
Range("C_MAPR").Value = ""
Range("C_MSFS").Value = ""
Range("C_MSEA").Value = ""

Range("C_AFR").Value = ""
Range("C_AVP").Value = ""
Range("C_GEH").Value = ""
Range("C_MZS").Value = ""
Range("C_APR").Value = ""
Range("C_SFS").Value = ""
Range("C_SEA").Value = ""

Range("C_ACF").Value = ""
Range("C_DK23").Value = ""
Range("C_DCT").Value = ""
Range("C_PAN").Value = ""
Range("C_REP").Value = ""
Range("C_SKU").Value = ""
Range("C_TVX").Value = ""
Range("C_MDE").Value = ""

Range("C_MACF").Value = ""
Range("C_MDK23").Value = ""
Range("C_MDCT").Value = ""
Range("C_MPAN").Value = ""
Range("C_MREP").Value = ""
Range("C_MSKU").Value = ""
Range("C_MTVX").Value = ""
Range("C_MMDE").Value = ""

Range("C_AG").Value = ""
Range("C_AC").Value = ""

Range("C_MAC").Value = ""
Range("C_MAG").Value = ""

Range("C_ACF2").Value = ""
Range("C_DK232").Value = ""
Range("C_DCT2").Value = ""
Range("C_PAN2").Value = ""
Range("C_REP2").Value = ""
Range("C_SKU2").Value = ""
Range("C_TVX2").Value = ""
Range("C_MDE2").Value = ""

Range("C_MACF2").Value = ""
Range("C_MDK232").Value = ""
Range("C_MDCT2").Value = ""
Range("C_MPAN2").Value = ""
Range("C_MREP2").Value = ""
Range("C_MSKU2").Value = ""
Range("C_MTVX2").Value = ""
Range("C_MMDE2").Value = ""


End Sub
Private Sub CB_C_CNX_Click()

TB_ACF.BackColor = &H80000005
CB_ACF.Value = False
TB_ACF.Visible = False
TB_ACF.Text = ""

TB_DK23.BackColor = &H80000005
CB_DK23.Value = False
TB_DK23.Visible = False
TB_DK23.Text = ""

TB_DCT.BackColor = &H80000005
CB_DCT.Value = False
TB_DCT.Visible = False
TB_DCT.Text = ""

TB_PAN.BackColor = &H80000005
CB_PAN.Value = False
TB_PAN.Visible = False
TB_PAN.Text = ""

TB_REP.BackColor = &H80000005
CB_REP.Value = False
TB_REP.Visible = False
TB_REP.Text = ""

TB_SKU.BackColor = &H80000005
CB_SKU.Value = False
TB_SKU.Visible = False
TB_SKU.Text = ""

TB_TVX.BackColor = &H80000005
CB_TVX.Value = False
TB_TVX.Visible = False
TB_TVX.Text = ""

TB_MDE.BackColor = &H80000005
CB_MDE.Value = False
TB_MDE.Visible = False
TB_MDE.Text = ""

TB_AFR_C.BackColor = &H80000005
TB_CAFR.Value = False
TB_AFR_C.Visible = False
TB_AFR_C.Text = ""

TB_APR_C.BackColor = &H80000005
TB_CAPR.Value = False
TB_APR_C.Visible = False
TB_APR_C.Text = ""

TB_AVP_C.BackColor = &H80000005
TB_CAVP.Value = False
TB_AVP_C.Visible = False
TB_AVP_C.Text = ""

TB_MZS_C.BackColor = &H80000005
TB_CMZS.Value = False
TB_MZS_C.Visible = False
TB_MZS_C.Text = ""

TB_GEH_C.BackColor = &H80000005
TB_CGEH.Value = False
TB_GEH_C.Visible = False
TB_GEH_C.Text = ""

TB_SFS_C.BackColor = &H80000005
TB_CSFS.Value = False
TB_SFS_C.Visible = False
TB_SFS_C.Text = ""

TB_SEA_C.BackColor = &H80000005
TB_CSEA.Value = False
TB_SEA_C.Visible = False
TB_SEA_C.Text = ""

TB_AC_C.BackColor = &H80000005
TB_CAC.Value = False
TB_AC_C.Visible = False
TB_AC_C.Text = ""

TB_AG_C.BackColor = &H80000005
TB_CAG.Value = False
TB_AG_C.Visible = False
TB_AG_C.Text = ""

Call C_RESET

CONVOY.Hide

End Sub

Private Sub CB_C_VAL_Click()

If Range("C_CNT1").Value = 1 Or Range("C_CNT2").Value = 1 Then

MsgBox "Le total de convoyages saisi diffère du total des types de convoyage selectionnés "

ElseIf CB_ACF.Value = True And TB_ACF.Text = "" Then

MsgBox "Le nombre de convoyage(s) Aéronef Code F est manquant"
TB_ACF.BackColor = &H8080FF

ElseIf CB_DK23.Value = True And TB_DK23.Text = "" Then

MsgBox "Le nombre de convoyage(s) Départ Kilo 23/6 est manquant"
TB_DK23.BackColor = &H8080FF

ElseIf CB_DCT.Value = True And TB_DCT.Text = "" Then

MsgBox "Le nombre de convoyage(s) Demande Client est manquant"
TB_ACA3.BackColor = &H8080FF

ElseIf CB_PAN.Value = True And TB_PAN.Text = "" Then

MsgBox "Le nombre de convoyage(s) Panne ou Annulation est manquant"
TB_PAN.BackColor = &H8080FF

ElseIf CB_REP.Value = True And TB_REP.Text = "" Then

MsgBox "Le nombre de convoyage(s) Repositionnement est manquant"
TB_REP.BackColor = &H8080FF

ElseIf CB_SKU.Value = True And TB_SKU.Text = "" Then

MsgBox "Le nombre de convoyage(s) Saturation Kilo Unité est manquant"
TB_SKU.BackColor = &H8080FF

ElseIf CB_TVX.Value = True And TB_TVX.Text = "" Then

MsgBox "Le nombre de convoyage(s) Travaux est manquant"
TB_TVX.BackColor = &H8080FF

ElseIf CB_MDE.Value = True And TB_MDE.Text = "" Then

MsgBox "Le nombre de convoyage(s) Modification Durée d'Escale est manquant"
TB_MDE.BackColor = &H8080FF

ElseIf TB_CAG.Value = True And TB_AG_C.Text = "" Then

MsgBox "Le nombre de tractage(s) Aviation Générale est manquant"
TB_AG_C.BackColor = &H8080FF

ElseIf TB_CAC.Value = True And TB_AC_C.Text = "" Then

MsgBox "Le nombre de tractage(s) Aviation Commercial est manquant"
TB_AC_C.BackColor = &H8080FF

Else

Call C_VAL_MAIN

End If

CONVOY.Hide

End Sub
Private Sub C_VAL_MAIN()

Dim ADD_ACF, ADD_DK23, ADD_DCT, ADD_PAN, ADD_REP, ADD_SKU, ADD_TVX, ADD_COM, ADD_GEN, ADD_MDE

ADD_ACF = Range("ADD_JJ_33").Value
ADD_DK23 = Range("ADD_JJ_35").Value
ADD_DCT = Range("ADD_JJ_34").Value
ADD_PAN = Range("ADD_JJ_38").Value
ADD_REP = Range("ADD_JJ_39").Value
ADD_SKU = Range("ADD_JJ_37").Value
ADD_TVX = Range("ADD_JJ_36").Value
ADD_COM = Range("ADD_JJ_42").Value
ADD_GEN = Range("ADD_JJ_41").Value
ADD_MDE = Range("ADD_JJ_32").Value

Sheets("LOG_M").Range(ADD_ACF) = Range("C_PACF2").Value
Sheets("LOG_M").Range(ADD_DK23) = Range("C_PDK232").Value
Sheets("LOG_M").Range(ADD_DCT) = Range("C_PDCT2").Value
Sheets("LOG_M").Range(ADD_PAN) = Range("C_PPAN2").Value
Sheets("LOG_M").Range(ADD_REP) = Range("C_PREP2").Value
Sheets("LOG_M").Range(ADD_SKU) = Range("C_PSKU2").Value
Sheets("LOG_M").Range(ADD_TVX) = Range("C_PTVX2").Value
Sheets("LOG_M").Range(ADD_MDE) = Range("C_PMDE2").Value
Sheets("LOG_M").Range(ADD_COM) = Range("C_PCOM2").Value
Sheets("LOG_M").Range(ADD_GEN) = Range("C_PGEN2").Value

End Sub
Private Sub CB_MDE_Click()

'Si l'on a cliqué sur le bouton CB_MDE le textbox correspondant s'affiche
If CB_MDE.Value = True Then
'La cellule correspondant à la valeur du bouton MDE, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_MDE2") = "1"
TB_MDE.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton MDE est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_MDE.Visible = False
TB_MDE.BackColor = &H80000005
Range("C_MDE2") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ
End If

End Sub

Private Sub TB_ACF_AfterUpdate()

Range("C_MACF2").Value = TB_ACF.Text

If TB_ACF.Text = "" Then
TB_ACF.BackColor = &H8080FF
Else
TB_ACF.BackColor = &H80000005
End If

End Sub
Private Sub TB_DCT_AfterUpdate()

Range("C_MDCT2").Value = TB_DCT.Text

If TB_DCT.Text = "" Then
TB_DCT.BackColor = &H8080FF
Else
TB_DCT.BackColor = &H80000005
End If

End Sub

Private Sub TB_DK23_AfterUpdate()

Range("C_MDK232").Value = TB_DK23.Text

Call C_TEXT_MAJ2

If TB_DK23.Text = "" Then
TB_DK23.BackColor = &H8080FF
Else
TB_DK23.BackColor = &H80000005
End If

End Sub

Private Sub TB_MDE_AfterUpdate()

Range("C_MMDE2").Value = TB_MDE.Text

If TB_MDE.Text = "" Then
TB_MDE.BackColor = &H8080FF
Else
TB_MDE.BackColor = &H80000005
End If

End Sub

Private Sub TB_PAN_AfterUpdate()

Range("C_MPAN2").Value = TB_PAN.Text

If TB_PAN.Text = "" Then
TB_PAN.BackColor = &H8080FF
Else
TB_PAN.BackColor = &H80000005
End If

End Sub

Private Sub TB_REP_AfterUpdate()

Range("C_MREP2").Value = TB_REP.Text

If TB_REP.Text = "" Then
TB_REP.BackColor = &H8080FF
Else
TB_REP.BackColor = &H80000005
End If

End Sub

Private Sub TB_SKU_AfterUpdate()

Range("C_MSKU2").Value = TB_SKU.Text

Call C_TEXT_MAJ2

If TB_SKU.Text = "" Then
TB_SKU.BackColor = &H8080FF
Else
TB_SKU.BackColor = &H80000005
End If

End Sub
Private Sub TB_TVX_AfterUpdate()

Range("C_MTVX2").Value = TB_TVX.Text

If TB_TVX.Text = "" Then
TB_TVX.BackColor = &H8080FF
Else
TB_TVX.BackColor = &H80000005
End If

End Sub
Private Sub CB_ACF_Click()

'Si l'on a cliqué sur le bouton CB_ACF le textbox correspondant s'affiche
If CB_ACF.Value = True Then
'La cellule correspondant à la valeur du bouton ACF, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_ACF2") = "1"
TB_ACF.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton ACF est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_ACF.Visible = False
TB_ACF.BackColor = &H80000005
Range("C_ACF2") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ
End If

End Sub
Private Sub CB_DK23_Click()

'Si l'on a cliqué sur le bouton CB_DK23 le textbox correspondant s'affiche
If CB_DK23.Value = True Then
'La cellule correspondant à la valeur du bouton DK23, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_DK232") = "1"
TB_DK23.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Détecte les tractages comme AG et enclenche le bouton AG correspondant et une MAJ des champs AG AC
TB_CAG.Value = True
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton DK23 est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_DK23.Visible = False
TB_DK23.BackColor = &H80000005
Range("C_DK232") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ

TB_CAG.Value = False
Call C_TEXT_MAJ2

End If


End Sub
Private Sub CB_DCT_Click()

'Si l'on a cliqué sur le bouton CB_DCT le textbox correspondant s'affiche
If CB_DCT.Value = True Then
'La cellule correspondant à la valeur du bouton DCT, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_DCT2") = "1"
TB_DCT.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton DCT est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_DCT.Visible = False
TB_DCT.BackColor = &H80000005
Range("C_DCT2") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ
End If


End Sub
Private Sub CB_PAN_Click()

'Si l'on a cliqué sur le bouton CB_PAN le textbox correspondant s'affiche
If CB_PAN.Value = True Then
'La cellule correspondant à la valeur du bouton PAN, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_PAN2") = "1"
TB_PAN.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton PAN est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_PAN.Visible = False
TB_PAN.BackColor = &H80000005
Range("C_PAN2") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ
End If

End Sub
Private Sub CB_REP_Click()

'Si l'on a cliqué sur le bouton CB_REP le textbox correspondant s'affiche
If CB_REP.Value = True Then
'La cellule correspondant à la valeur du bouton REP, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_REP2") = "1"
TB_REP.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton REP est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_REP.Visible = False
TB_REP.BackColor = &H80000005
Range("C_REP2") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ
End If

End Sub
Private Sub CB_SKU_Click()

'Si l'on a cliqué sur le bouton CB_SKU le textbox correspondant s'affiche
If CB_SKU.Value = True Then
'La cellule correspondant à la valeur du bouton SKU, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_SKU2") = "1"
TB_SKU.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Détecte les tractages comme AG et enclenche le bouton AG correspondant et une MAJ des champs AG AC
TB_CAG.Value = True
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton SKU est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_SKU.Visible = False
TB_SKU.BackColor = &H80000005
Range("C_SKU2") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ

TB_CAG.Value = False
Call C_TEXT_MAJ2

End If

End Sub
Private Sub CB_TVX_Click()

'Si l'on a cliqué sur le bouton CB_TVX le textbox correspondant s'affiche
If CB_TVX.Value = True Then
'La cellule correspondant à la valeur du bouton TVX, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_TVX2") = "1"
TB_TVX.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton TVX est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_TVX.Visible = False
TB_TVX.BackColor = &H80000005
Range("C_TVX2") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ
End If

End Sub
Private Sub TB_CAFR_Click()

'Si l'on a cliqué sur le bouton TB_CAFR le textbox correspondant s'affiche
If TB_CAFR.Value = True Then
'La cellule correspondant à la valeur du bouton AFR, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_AFR") = "1"
TB_AFR_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton AFR est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_AFR_C.Visible = False
TB_AFR_C.BackColor = &H80000005
Range("C_AFR") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If

End Sub
Private Sub TB_CAVP_Click()

'Si l'on a cliqué sur le bouton TB_CAVP le textbox correspondant s'affiche
If TB_CAVP.Value = True Then
'La cellule correspondant à la valeur du bouton AVP, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_AVP") = "1"
TB_AVP_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton AVP est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_AVP_C.Visible = False
TB_AVP_C.BackColor = &H80000005
Range("C_AVP") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If


End Sub
Private Sub TB_CGEH_Click()

'Si l'on a cliqué sur le bouton TB_CGEH le textbox correspondant s'affiche
If TB_CGEH.Value = True Then
'La cellule correspondant à la valeur du bouton GEH, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_GEH") = "1"
TB_GEH_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton GEH est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_GEH_C.Visible = False
TB_GEH_C.BackColor = &H80000005
Range("C_GEH") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If

End Sub
Private Sub TB_CMZS_Click()

'Si l'on a cliqué sur le bouton TB_CMZS le textbox correspondant s'affiche
If TB_CMZS.Value = True Then
'La cellule correspondant à la valeur du bouton MZS, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_MZS") = "1"
TB_MZS_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton MZS est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_MZS_C.Visible = False
TB_MZS_C.BackColor = &H80000005
Range("C_MZS") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If

End Sub
Private Sub TB_CAPR_Click()

'Si l'on a cliqué sur le bouton TB_CAPR le textbox correspondant s'affiche
If TB_CAPR.Value = True Then
'La cellule correspondant à la valeur du bouton APR, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_APR") = "1"
TB_APR_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton APR est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_APR_C.Visible = False
TB_APR_C.BackColor = &H80000005
Range("C_APR") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If

End Sub
Private Sub TB_CSEA_Click()

'Si l'on a cliqué sur le bouton TB_CSEA le textbox correspondant s'affiche
If TB_CSEA.Value = True Then
'La cellule correspondant à la valeur du bouton SEA, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_SEA") = "1"
TB_SEA_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton SEA est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_SEA_C.Visible = False
TB_SEA_C.BackColor = &H80000005
Range("C_SEA") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If


End Sub
Private Sub TB_CSFS_Click()

'Si l'on a cliqué sur le bouton TB_CSFS le textbox correspondant s'affiche
If TB_CSFS.Value = True Then
'La cellule correspondant à la valeur du bouton SFS, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("C_SFS") = "1"
TB_SFS_C.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call C_TEXT_MAJ2
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton SFS est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_SFS_C.Visible = False
TB_SFS_C.BackColor = &H80000005
Range("C_SFS") = ""
'Les valeurs des textbox sont mises à jour via la procédure C_TEXT_MAJ
Call C_TEXT_MAJ2
End If

End Sub
Private Sub TB_AC_C_AfterUpdate()

Range("C_MAC").Value = TB_AC_C.Text

If TB_AC_C.Text = "" Then
TB_AC_C.BackColor = &H8080FF
Else
TB_AC_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_AG_C_AfterUpdate()

Range("C_MAG").Value = TB_AG_C.Text

If TB_AG_C.Text = "" Then
TB_AG_C.BackColor = &H8080FF
Else
TB_AG_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_AFR_C_AfterUpdate()

Range("C_MAFR").Value = TB_AFR_C.Text

If TB_AFR_C.Text = "" Then
TB_AFR_C.BackColor = &H8080FF
Else
TB_AFR_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_AVP_C_AfterUpdate()

Range("C_MAVP").Value = TB_AVP_C.Text

If TB_AVP_C.Text = "" Then
TB_AVP_C.BackColor = &H8080FF
Else
TB_AVP_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_GEH_C_AfterUpdate()

Range("C_MGEH").Value = TB_GEH_C.Text

If TB_GEH_C.Text = "" Then
TB_GEH_C.BackColor = &H8080FF
Else
TB_GEH_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_MZS_C_AfterUpdate()

Range("C_MMZS").Value = TB_MZS_C.Text

If TB_MZS_C.Text = "" Then
TB_MZS_C.BackColor = &H8080FF
Else
TB_MZS_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_APR_C_AfterUpdate()

Range("C_MAPR").Value = TB_APR_C.Text

If TB_APR_C.Text = "" Then
TB_APR_C.BackColor = &H8080FF
Else
TB_APR_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_SFS_C_AfterUpdate()

Range("C_MSFS").Value = TB_SFS_C.Text

If TB_SFS_C.Text = "" Then
TB_SFS_C.BackColor = &H8080FF
Else
TB_SFS_C.BackColor = &H80000005
End If

End Sub
Private Sub TB_SEA_C_AfterUpdate()

Range("C_MSEA").Value = TB_SEA_C.Text

If TB_SEA_C.Text = "" Then
TB_SEA_C.BackColor = &H8080FF
Else
TB_SEA_C.BackColor = &H80000005
End If

End Sub
