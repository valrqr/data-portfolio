' Module : FRET.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_DIV_ANN_Click()

TB_ENT.BackColor = &H80000005
CB_ENT.Value = False
TB_ENT.Visible = False
TB_ENT.Text = ""

TB_SOR.BackColor = &H80000005
CB_SOR.Value = False
TB_SOR.Visible = False
TB_SOR.Text = ""

TB_M30.BackColor = &H80000005
CB_M30.Value = False
TB_M30.Visible = False
TB_M30.Text = ""

TB_P30.BackColor = &H80000005
CB_P30.Value = False
TB_P30.Visible = False
TB_P30.Text = ""

TB_1H.BackColor = &H80000005
CB_1H.Value = False
TB_1H.Visible = False
TB_1H.Text = ""

TB_FRET2.BackColor = &H80000005
TB_FRET2.Text = ""

FRET.Hide

Range("FRET_ENT") = ""
Range("FRET_SOR") = ""
Range("FRET_M30") = ""
Range("FRET_P30") = ""
Range("FRET_1H") = ""

End Sub

Private Sub CB_DIV_VAL_Click()

If Range("FRET_CNT") = 1 Then

MsgBox "ERREUR: Le nombre d'occurences saisie ne correspond pas"

ElseIf CB_ENT.Value = True And TB_ENT.Text = "" Then

MsgBox "Le nombre d'Entrée(s) FRET est manquant"
TB_ENT.BackColor = &H8080FF

ElseIf CB_SOR.Value = True And TB_SOR.Text = "" Then

MsgBox "Le nombre de Sortie(s) FRET est manquant"
TB_SOR.BackColor = &H8080FF

ElseIf CB_M30.Value = True And TB_M30.Text = "" Then

MsgBox "Le nombre de Tranche(s) inférieure(s) à 30MIN est manquant"
TB_M30.BackColor = &H8080FF

ElseIf CB_P30.Value = True And TB_P30.Text = "" Then

MsgBox "Le nombre de Tranche(s) supérieure(s) à 30MIN est manquant"
TB_P30.BackColor = &H8080FF

ElseIf CB_1H.Value = True And TB_1H.Text = "" Then

MsgBox "Le nombre de Tranche(s) supérieure(s) à 1H est manquant"
TB_1H.BackColor = &H8080FF

ElseIf TB_FRET2.Text = "" Then

MsgBox "Le(s) numéro(s) de GECO FRET est/sont manquants"
TB_FRET2.BackColor = &H8080FF

ElseIf TB_FRET2.Text <> "" Then

Dim ADD_FRET2, ADD_ENT, ADD_SOR, ADD_M30, ADD_P30, ADD_1H

ADD_FRET2 = Range("ADD_JJ_31").Value
ADD_ENT = Range("ADD_JJ_26").Value
ADD_SOR = Range("ADD_JJ_27").Value
ADD_M30 = Range("ADD_JJ_28").Value
ADD_P30 = Range("ADD_JJ_29").Value
ADD_1H = Range("ADD_JJ_30").Value

Sheets("LOG_M").Range(ADD_FRET2) = TB_FRET2.Text
Sheets("LOG_M").Range(ADD_ENT) = Range("FRET_PENT").Value
Sheets("LOG_M").Range(ADD_SOR) = Range("FRET_PSOR").Value
Sheets("LOG_M").Range(ADD_M30) = Range("FRET_PM30").Value
Sheets("LOG_M").Range(ADD_P30) = Range("FRET_PP30").Value
Sheets("LOG_M").Range(ADD_1H) = Range("FRET_P1H").Value

End If

FRET.Hide

End Sub
Private Sub CB_ENT_Click()

'Si l'on a cliqué sur le bouton CB_ENT le textbox correspondant s'affiche
If CB_ENT.Value = True Then
'La cellule correspondant à la valeur du bouton ENT, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("FRET_ENT") = "1"
TB_ENT.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call FRET_TEXT_MAJ
'Si les conditions déclarées precedemment ne sont pas vraies, alors
Else
'La textbox attenante au bouton ENT est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_ENT.Visible = False
TB_ENT.BackColor = &H80000005
Range("FRET_ENT") = ""
'Les valeurs des textbox sont mises à jour via la procédure FRET_TEXT_MAJ
Call FRET_TEXT_MAJ
End If

End Sub
Private Sub CB_SOR_Click()

'Si l'on a cliqué sur le bouton CB_SOR le textbox correspondant s'affiche
If CB_SOR.Value = True Then
'La cellule correspondant à la valeur du bouton SOR, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("FRET_SOR") = "1"
TB_SOR.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call FRET_TEXT_MAJ
'Si les conditions déclarées precedemmSOR ne sont pas vraies, alors
Else
'La textbox attenante au bouton SOR est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_SOR.Visible = False
TB_SOR.BackColor = &H80000005
Range("FRET_SOR") = ""
'Les valeurs des textbox sont mises à jour via la procédure FRET_TEXT_MAJ
Call FRET_TEXT_MAJ
End If

End Sub
Private Sub CB_M30_Click()

'Si l'on a cliqué sur le bouton CB_M30 le textbox correspondant s'affiche
If CB_M30.Value = True Then
'La cellule correspondant à la valeur du bouton M30, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("FRET_M30") = "1"
TB_M30.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call FRET_TEXT_MAJ
'Si les conditions déclarées precedemmM30 ne sont pas vraies, alors
Else
'La textbox attenante au bouton M30 est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_M30.Visible = False
TB_M30.BackColor = &H80000005
Range("FRET_M30") = ""
'Les valeurs des textbox sont mises à jour via la procédure FRET_TEXT_MAJ
Call FRET_TEXT_MAJ
End If

End Sub
Private Sub CB_P30_Click()

'Si l'on a cliqué sur le bouton CB_P30 le textbox correspondant s'affiche
If CB_P30.Value = True Then
'La cellule correspondant à la valeur du bouton P30, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("FRET_P30") = "1"
TB_P30.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call FRET_TEXT_MAJ
'Si les conditions déclarées precedemmP30 ne sont pas vraies, alors
Else
'La textbox attenante au bouton P30 est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_P30.Visible = False
TB_P30.BackColor = &H80000005
Range("FRET_P30") = ""
'Les valeurs des textbox sont mises à jour via la procédure FRET_TEXT_MAJ
Call FRET_TEXT_MAJ
End If

End Sub
Private Sub CB_1H_Click()

'Si l'on a cliqué sur le bouton CB_1H le textbox correspondant s'affiche
If CB_1H.Value = True Then
'La cellule correspondant à la valeur du bouton 1H, 1 ou 0, est activé en feuille DATA_MAIN et la textbox correspondante est rendue visible
Range("FRET_1H") = "1"
TB_1H.Visible = True
'Les textbox sont mises à jour avec les données dites automatisées, colonne 1 de DATA_MAIN
Call FRET_TEXT_MAJ
'Si les conditions déclarées precedemm1H ne sont pas vraies, alors
Else
'La textbox attenante au bouton 1H est cachée, sa couleur est blanche et la cellule correspondante en DATA_MAIN est vidée.
TB_1H.Visible = False
TB_1H.BackColor = &H80000005
Range("FRET_1H") = ""
'Les valeurs des textbox sont mises à jour via la procédure FRET_TEXT_MAJ
Call FRET_TEXT_MAJ
End If

End Sub

Private Sub TB_PLUS_FRET_Change()

Range("DATA_VAL_MAIN").Value = TB_PLUS_FRET.Text

End Sub

Private Sub TGB_MS_FRET_Click()

If Range("DATA_VAL_MAIN").Value = 1 Then

TGB_MS_FRET.Visible = False

Else

Range("DATA_VAL_MAIN").Value = Range("DATA_VAL_MAIN").Value - 1
TB_PLUS_FRET.Text = Range("DATA_VAL_MAIN").Value

Call FRET_TEXT_MAJ

End If

End Sub

Private Sub TGB_PL_FRET_Click()

Range("DATA_VAL_MAIN").Value = Range("DATA_VAL_MAIN").Value + 1
TB_PLUS_FRET.Text = Range("DATA_VAL_MAIN").Value

Call FRET_TEXT_MAJ

TGB_MS_FRET.Visible = True

End Sub
Private Sub FRET_TEXT_MAJ()

TB_ENT.Text = Range("FRET_AENT")
TB_SOR.Text = Range("FRET_ASOR")
TB_M30.Text = Range("FRET_AM30")
TB_P30.Text = Range("FRET_AP30")
TB_1H.Text = Range("FRET_A1H")

End Sub
Private Sub Userform_Activate()

Range("JOUR_X") = Range("JOUR_J")

TB_PLUS_FRET.Text = 1
Range("DATA_VAL_MAIN").Value = 1
TGB_MS_FRET.Visible = False

LB_DIV_Gn.Caption = Range("GECO_N").Value

CB_ENT.Value = False
TB_ENT.Visible = False
TB_ENT.Text = ""

CB_SOR.Value = False
TB_SOR.Visible = False
TB_SOR.Text = ""

CB_M30.Value = False
TB_M30.Visible = False
TB_M30.Text = ""

CB_P30.Value = False
TB_P30.Visible = False
TB_P30.Text = ""

CB_1H.Value = False
TB_1H.Visible = False
TB_1H.Text = ""

TB_FRET2.Visible = True
TB_FRET2.Text = "'"

Range("FRET_ENT") = ""
Range("FRET_SOR") = ""
Range("FRET_M30") = ""
Range("FRET_P30") = ""
Range("FRET_1H") = ""

Range("FRET_MENT") = ""
Range("FRET_MSOR") = ""
Range("FRET_MM30") = ""
Range("FRET_MP30") = ""
Range("FRET_M1H") = ""

End Sub
Private Sub TB_1H_AfterUpdate()

Range("FRET_M1H").Value = TB_1H.Text

If TB_1H.Text = "" Then
TB_1H.BackColor = &H8080FF
Else
TB_1H.BackColor = &H80000005
End If

End Sub

Private Sub TB_ENT_AfterUpdate()

Range("FRET_MENT").Value = TB_ENT.Text

If TB_ENT.Text = "" Then
TB_ENT.BackColor = &H8080FF
Else
TB_ENT.BackColor = &H80000005
End If

End Sub
Private Sub TB_FRET2_Change()
If TB_FRET2.Text = "" Then
TB_FRET2.BackColor = &H8080FF
Else
TB_FRET2.BackColor = &H80000005
End If
End Sub
Private Sub TB_M30_AfterUpdate()

Range("FRET_MM30").Value = TB_M30.Text

If TB_M30.Text = "" Then
TB_M30.BackColor = &H8080FF
Else
TB_M30.BackColor = &H80000005
End If

End Sub
Private Sub TB_P30_AfterUpdate()

Range("FRET_MP30").Value = TB_P30.Text

If TB_P30.Text = "" Then
TB_P30.BackColor = &H8080FF
Else
TB_P30.BackColor = &H80000005
End If

End Sub
Private Sub TB_SOR_AfterUpdate()

Range("FRET_MSOR").Value = TB_SOR.Text

If TB_SOR.Text = "" Then
TB_SOR.BackColor = &H8080FF
Else
TB_SOR.BackColor = &H80000005
End If

End Sub
