' Module : SURTOU_2.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

Private Sub CB_VAL_ST2_Click()


Sheets("LOG_ST").Range("Q6").End(xlDown).Offset(1, 0) = Range("ST_IPA1").Value
Sheets("LOG_ST").Range("R6").End(xlDown).Offset(1, 0) = Range("ST_PPP1").Value
Sheets("LOG_ST").Range("S6").End(xlDown).Offset(1, 0) = Range("ST_ZD1").Value
Sheets("LOG_ST").Range("T6").End(xlDown).Offset(1, 0) = Range("ST_PARR").Value
Sheets("LOG_ST").Range("U6").End(xlDown).Offset(1, 0) = Range("ST_MIR").Value
Sheets("LOG_ST").Range("V6").End(xlDown).Offset(1, 0) = Range("ST_4001").Value
Sheets("LOG_ST").Range("W6").End(xlDown).Offset(1, 0) = Range("ST_ACO").Value
Sheets("LOG_ST").Range("X6").End(xlDown).Offset(1, 0) = Range("ST_CMP").Value
Sheets("LOG_ST").Range("Y6").End(xlDown).Offset(1, 0) = Range("ST_ECD").Value
Sheets("LOG_ST").Range("Z6").End(xlDown).Offset(1, 0) = Range("ST_DT").Value
Sheets("LOG_ST").Range("AA6").End(xlDown).Offset(1, 0) = Range("ST_GC").Value
Sheets("LOG_ST").Range("AB6").End(xlDown).Offset(1, 0) = Range("ST_DAO").Value
Sheets("LOG_ST").Range("AC6").End(xlDown).Offset(1, 0) = Range("ST_PSI").Value
Sheets("LOG_ST").Range("AD6").End(xlDown).Offset(1, 0) = TB_CBV.Text
Sheets("LOG_ST").Range("AE6").End(xlDown).Offset(1, 0) = Range("ST_LOD").Value
Sheets("LOG_ST").Range("AF6").End(xlDown).Offset(1, 0) = Range("ST_CAT").Value
Sheets("LOG_ST").Range("AG6").End(xlDown).Offset(1, 0) = Range("ST_CPM").Value
Sheets("LOG_ST").Range("AH6").End(xlDown).Offset(1, 0) = TB_CAB.Text
Sheets("LOG_ST").Range("AI6").End(xlDown).Offset(1, 0) = Range("ST_WW").Value
Sheets("LOG_ST").Range("AJ6").End(xlDown).Offset(1, 0) = Range("ST_CTL").Value
Sheets("LOG_ST").Range("AK6").End(xlDown).Offset(1, 0) = Range("ST_CC").Value
Sheets("LOG_ST").Range("AL6").End(xlDown).Offset(1, 0) = Range("ST_VDP").Value
Sheets("LOG_ST").Range("AM6").End(xlDown).Offset(1, 0) = Range("ST_4002").Value
Sheets("LOG_ST").Range("AN6").End(xlDown).Offset(1, 0) = Range("ST_PPP2").Value
Sheets("LOG_ST").Range("AO6").End(xlDown).Offset(1, 0) = Range("ST_ZD2").Value
Sheets("LOG_ST").Range("AP6").End(xlDown).Offset(1, 0) = Range("ST_REPO").Value
Sheets("LOG_ST").Range("AQ6").End(xlDown).Offset(1, 0) = Range("ST_IPA2").Value

SURTOU_2.Hide
SURTOU_1.Hide

ActiveWorkbook.Save

End Sub

Private Sub ST_OK1_Click()

CBO_1.Value = True
CBO_2.Value = True
CBO_3.Value = True
CBO_4.Value = True
CBO_5.Value = True

End Sub

Private Sub ST_OK2_Click()

CBO_6.Value = True
CBO_7.Value = True
CBO_8.Value = True

End Sub

Private Sub ST_OK3_Click()

CBO_9.Value = True
CBO_12.Value = True
CBO_13.Value = True

End Sub

Private Sub ST_OK4_Click()

CBO_19.Value = True
CBO_20.Value = True
CBO_21.Value = True
CBO_22.Value = True

End Sub

Private Sub ST_OK5_Click()

CBO_23.Value = True
CBO_24.Value = True
CBO_25.Value = True
CBO_26.Value = True
CBO_27.Value = True

End Sub

Private Sub TB_CAB_Enter()
MsgBox ("Saisir un nombre entier, comptabilisant le nombre de conteneurs total")
End Sub

Private Sub TB_CBV_Enter()
MsgBox ("Saisir un nombre entier, comptabilisant le nombre de chariots total")
End Sub

Private Sub Userform_Activate()

TB_CAB.Text = ""
TB_CBV.Text = ""

CBO_1.Value = False
CBN_1.Value = False
CBNA_1.Value = False

CBO_2.Value = False
CBN_2.Value = False
CBNA_2.Value = False

CBO_3.Value = False
CBN_3.Value = False
CBNA_3.Value = False

CBO_3.Value = False
CBN_3.Value = False
CBNA_3.Value = False

CBO_4.Value = False
CBN_4.Value = False
CBNA_4.Value = False

CBO_5.Value = False
CBN_5.Value = False
CBNA_5.Value = False

CBO_6.Value = False
CBN_6.Value = False
CBNA_6.Value = False

CBO_7.Value = False
CBN_7.Value = False
CBNA_7.Value = False

CBO_8.Value = False
CBN_8.Value = False
CBNA_8.Value = False

CBO_9.Value = False
CBN_9.Value = False
CBNA_9.Value = False

CBO_10.Value = False
CBN_10.Value = False
CBNA_10.Value = False

CBO_11.Value = False
CBN_11.Value = False
CBNA_11.Value = False

CBO_12.Value = False
CBN_12.Value = False
CBNA_12.Value = False

CBO_13.Value = False
CBN_13.Value = False
CBNA_13.Value = False

CBO_14.Value = False
CBN_14.Value = False
CBNA_14.Value = False

CBO_15.Value = False
CBN_15.Value = False
CBNA_15.Value = False

CBO_16.Value = False
CBN_16.Value = False
CBNA_16.Value = False

CBO_17.Value = False
CBN_17.Value = False
CBNA_17.Value = False

CBO_18.Value = False
CBN_18.Value = False
CBNA_18.Value = False

CBO_19.Value = False
CBN_19.Value = False
CBNA_19.Value = False

CBO_20.Value = False
CBN_20.Value = False
CBNA_20.Value = False

CBO_21.Value = False
CBN_21.Value = False
CBNA_21.Value = False

CBO_22.Value = False
CBN_22.Value = False
CBNA_22.Value = False

CBO_23.Value = False
CBN_23.Value = False
CBNA_23.Value = False

CBO_24.Value = False
CBN_24.Value = False
CBNA_24.Value = False

CBO_25.Value = False
CBN_25.Value = False
CBNA_25.Value = False

CBO_26.Value = False
CBN_26.Value = False
CBNA_26.Value = False

CBO_27.Value = False
CBN_27.Value = False
CBNA_27.Value = False

If Range("ST_PKG_WW").Value <> 2 Then
CBNA_19.Value = True
CBNA_20.Value = True
CBNA_21.Value = True
CBNA_22.Value = True

End If

If Range("ST_PKG_LG").Value <> 1 Then
CBNA_2.Value = True
CBNA_24.Value = True
End If

Range("ST_IPA1") = "X"
Range("ST_PPP1") = "X"
Range("ST_ZD1") = "X"
Range("ST_PARR") = "X"
Range("ST_MIR") = "X"
Range("ST_4001") = "X"
Range("ST_ACO") = "X"
Range("ST_CMP") = "X"
Range("ST_ECD") = "X"
Range("ST_DT") = "X"
Range("ST_GC") = "X"
Range("ST_DAO") = "X"
Range("ST_PSI") = "X"
Range("ST_cbv") = "X"
Range("ST_LOD") = "X"
Range("ST_CAT") = "X"
Range("ST_CPM") = "X"
Range("ST_AKH") = "X"
Range("ST_WW") = "X"
Range("ST_CTL") = "X"
Range("ST_CC") = "X"
Range("ST_VDP") = "X"
Range("ST_4002") = "X"
Range("ST_PPP2") = "X"
Range("ST_ZD2") = "X"
Range("ST_REPO") = "X"
Range("ST_IPA2") = "X"

End Sub

Private Sub CB_ANS2_Click()

Range("ST_IPA1") = "X"
Range("ST_PPP1") = "X"
Range("ST_ZD1") = "X"
Range("ST_PARR") = "X"
Range("ST_MIR") = "X"
Range("ST_4001") = "X"
Range("ST_ACO") = "X"
Range("ST_CMP") = "X"
Range("ST_ECD") = "X"
Range("ST_DT") = "X"
Range("ST_GC") = "X"
Range("ST_DAO") = "X"
Range("ST_PSI") = "X"
Range("ST_cbv") = "X"
Range("ST_LOD") = "X"
Range("ST_CAT") = "X"
Range("ST_CPM") = "X"
Range("ST_AKH") = "X"
Range("ST_WW") = "X"
Range("ST_CTL") = "X"
Range("ST_CC") = "X"
Range("ST_VDP") = "X"
Range("ST_4002") = "X"
Range("ST_PPP2") = "X"
Range("ST_ZD2") = "X"
Range("ST_REPO") = "X"
Range("ST_IPA2") = "X"

SURTOU_2.Hide
SURTOU_1.Hide
End Sub
Private Sub CBO_1_Click()
If CBO_1.Value = True Then
CBN_1.Value = False
CBNA_1.Value = False
Range("ST_IPA1") = 1
Else
Range("ST_IPA1") = "X"
End If
End Sub
Private Sub CBN_1_Click()
If CBN_1.Value = True Then
CBO_1.Value = False
CBNA_1.Value = False
Range("ST_IPA1") = 2
Else
Range("ST_IPA1") = "X"
End If
End Sub
Private Sub CBNA_1_Click()
If CBNA_1.Value = True Then
CBN_1.Value = False
CBO_1.Value = False
Range("ST_IPA1") = 0
Else
Range("ST_IPA1") = "X"
End If
End Sub
Private Sub CBO_2_Click()
If CBO_2.Value = True Then
CBN_2.Value = False
CBNA_2.Value = False
Range("ST_PPP1") = 1
Else
Range("ST_PPP1") = "X"
End If
End Sub
Private Sub CBN_2_Click()
If CBN_2.Value = True Then
CBO_2.Value = False
CBNA_2.Value = False
Range("ST_PPP1") = 2
Else
Range("ST_PPP1") = "X"
End If
End Sub
Private Sub CBNA_2_Click()
If CBNA_2.Value = True Then
CBN_2.Value = False
CBO_2.Value = False
Range("ST_PPP1") = 0
Else
Range("ST_PPP1") = "X"
End If
End Sub
Private Sub CBO_3_Click()
If CBO_3.Value = True Then
CBN_3.Value = False
CBNA_3.Value = False
Range("ST_ZD1") = 1
Else
Range("ST_ZD1") = "X"
End If
End Sub
Private Sub CBN_3_Click()
If CBN_3.Value = True Then
CBO_3.Value = False
CBNA_3.Value = False
Range("ST_ZD1") = 2
Else
Range("ST_ZD1") = "X"
End If
End Sub
Private Sub CBNA_3_Click()
If CBNA_3.Value = True Then
CBN_3.Value = False
CBO_3.Value = False
Range("ST_ZD1") = 0
Else
Range("ST_ZD1") = "X"
End If
End Sub
Private Sub CBO_4_Click()
If CBO_4.Value = True Then
CBN_4.Value = False
CBNA_4.Value = False
Range("ST_PARR") = 1
Else
Range("ST_PARR") = "X"
End If
End Sub
Private Sub CBN_4_Click()
If CBN_4.Value = True Then
CBO_4.Value = False
CBNA_4.Value = False
Range("ST_PARR") = 2
Else
Range("ST_PARR") = "X"
End If
End Sub
Private Sub CBNA_4_Click()
If CBNA_4.Value = True Then
CBN_4.Value = False
CBO_4.Value = False
Range("ST_PARR") = 0
Else
Range("ST_PARR") = "X"
End If
End Sub
Private Sub CBO_5_Click()
If CBO_5.Value = True Then
CBN_5.Value = False
CBNA_5.Value = False
Range("ST_MIR") = 1
Else
Range("ST_MIR") = "X"
End If
End Sub
Private Sub CBN_5_Click()
If CBN_5.Value = True Then
CBO_5.Value = False
CBNA_5.Value = False
Range("ST_MIR") = 2
Else
Range("ST_MIR") = "X"
End If
End Sub
Private Sub CBNA_5_Click()
If CBNA_5.Value = True Then
CBN_5.Value = False
CBO_5.Value = False
Range("ST_MIR") = 0
Else
Range("ST_MIR") = "X"
End If
End Sub
Private Sub CBO_6_Click()
If CBO_6.Value = True Then
If Range("TOU_CT1") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_6.Value = False
CBNA_6.Value = False
Range("ST_4001") = 1
Else
Range("ST_4001") = "X"
End If
End Sub
Private Sub CBN_6_Click()
If CBN_6.Value = True Then
If Range("TOU_CT1") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_6.Value = False
CBNA_6.Value = False
Range("ST_4001") = 2
Else
Range("ST_4001") = "X"
End If
End Sub
Private Sub CBNA_6_Click()
If CBNA_6.Value = True Then
If Range("TOU_CT1") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_6.Value = False
CBN_6.Value = False
Range("ST_4001") = 0
Else
Range("ST_4001") = "X"
End If
End Sub
Private Sub CBO_7_Click()
If CBO_7.Value = True Then
If Range("TOU_CT1") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_7.Value = False
CBNA_7.Value = False
Range("ST_ACO") = 1
Else
Range("ST_ACO") = "X"
End If
End Sub
Private Sub CBN_7_Click()
If CBN_7.Value = True Then
If Range("TOU_CT1") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_7.Value = False
CBNA_7.Value = False
Range("ST_ACO") = 2
Else
Range("ST_ACO") = "X"
End If
End Sub
Private Sub CBNA_7_Click()
If CBNA_7.Value = True Then
If Range("TOU_CT1") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_7.Value = False
CBN_7.Value = False
Range("ST_ACO") = 0
Else
Range("ST_ACO") = "X"
End If
End Sub
Private Sub CBO_8_Click()
If CBO_8.Value = True Then
CBN_8.Value = False
CBNA_8.Value = False
Range("ST_CMP") = 1
Else
Range("ST_CMP") = "X"
End If
End Sub
Private Sub CBN_8_Click()
If CBN_8.Value = True Then
CBO_8.Value = False
CBNA_8.Value = False
Range("ST_CMP") = 2
Else
Range("ST_CMP") = "X"
End If
End Sub
Private Sub CBNA_8_Click()
If CBNA_8.Value = True Then
CBN_8.Value = False
CBO_8.Value = False
Range("ST_CMP") = 0
Else
Range("ST_CMP") = "X"
End If
End Sub
Private Sub CBO_9_Click()
If CBO_9.Value = True Then
If Range("TOU_CT2") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_9.Value = False
CBNA_9.Value = False
Range("ST_ECD") = 1
Else
Range("ST_ECD") = "X"
End If
End Sub
Private Sub CBN_9_Click()
If CBN_9.Value = True Then
If Range("TOU_CT2") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_9.Value = False
CBNA_9.Value = False
Range("ST_ECD") = 2
Else
Range("ST_ECD") = "X"
End If
End Sub
Private Sub CBNA_9_Click()
If CBNA_9.Value = True Then
If Range("TOU_CT2") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_9.Value = False
CBN_9.Value = False
Range("ST_ECD") = 0
CBO_10.Value = False
CBN_10.Value = False
CBNA_10.Value = True
Range("ST_DT") = 0
CBO_11.Value = False
CBN_11.Value = False
CBNA_11.Value = True
Range("ST_GC") = 0
CBO_12.Value = False
CBN_12.Value = False
CBNA_12.Value = True
Range("ST_DAO") = 0
CBO_13.Value = False
CBN_13.Value = False
CBNA_13.Value = True
Range("ST_PSI") = 0
Else
Range("ST_ECD") = "X"
CBO_10.Value = False
CBN_10.Value = False
CBNA_10.Value = False
Range("ST_DT") = "X"
CBO_11.Value = False
CBN_11.Value = False
CBNA_11.Value = False
Range("ST_GC") = "X"
CBO_12.Value = False
CBN_12.Value = False
CBNA_12.Value = False
Range("ST_DAO") = "X"
CBO_13.Value = False
CBN_13.Value = False
CBNA_13.Value = False
Range("ST_PSI") = "X"
End If
End Sub
Private Sub CBO_10_Click()
If CBO_10.Value = True Then
If Range("TOU_CT2") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_10.Value = False
CBNA_10.Value = False
Range("ST_DT") = 1
CBO_11.Value = False
CBNA_11.Value = True
CBN_11.Value = False
Range("ST_GC") = 1
Else
Range("ST_DT") = "X"
CBO_11.Value = False
CBN_11.Value = False
CBNA_11.Value = False
Range("ST_GC") = "X"
End If
End Sub
Private Sub CBN_10_Click()
If CBN_10.Value = True Then
If Range("TOU_CT2") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_10.Value = False
CBNA_10.Value = False
Range("ST_DT") = 2
CBO_11.Value = True
CBN_11.Value = False
CBNA_11.Value = False
Range("ST_GC") = 2
Else
Range("ST_DT") = "X"
CBO_11.Value = False
CBN_11.Value = False
CBNA_11.Value = False
Range("ST_GC") = "X"
End If
End Sub
Private Sub CBNA_10_Click()
If CBNA_10.Value = True Then
If Range("TOU_CT2") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_10.Value = False
CBN_10.Value = False
Range("ST_DT") = 0
Else
Range("ST_DT") = "X"
End If
End Sub
Private Sub CBO_11_Click()
If CBO_11.Value = True Then
CBN_11.Value = False
CBNA_11.Value = False
Range("ST_GC") = 1
Else
Range("ST_GC") = "X"
End If
End Sub
Private Sub CBN_11_Click()
If CBN_11.Value = True Then
CBO_11.Value = False
CBNA_11.Value = False
Range("ST_GC") = 2
Else
Range("ST_GC") = "X"
End If
End Sub
Private Sub CBNA_11_Click()
If CBNA_11.Value = True Then
CBN_11.Value = False
CBO_11.Value = False
Range("ST_GC") = 0
Else
Range("ST_GC") = "X"
End If
End Sub
Private Sub CBO_12_Click()
If CBO_12.Value = True Then
CBN_12.Value = False
CBNA_12.Value = False
Range("ST_DAO") = 1
Else
Range("ST_DAO") = "X"
End If
End Sub
Private Sub CBN_12_Click()
If CBN_12.Value = True Then
CBO_12.Value = False
CBNA_12.Value = False
Range("ST_DAO") = 2
Else
Range("ST_DAO") = "X"
End If
End Sub
Private Sub CBNA_12_Click()
If CBNA_12.Value = True Then
CBN_12.Value = False
CBO_12.Value = False
Range("ST_DAO") = 0
Else
Range("ST_DAO") = "X"
End If
End Sub
Private Sub CBO_13_Click()
If CBO_13.Value = True Then
CBN_13.Value = False
CBNA_13.Value = False
Range("ST_PSI") = 1
Else
Range("ST_PSI") = "X"
End If
End Sub
Private Sub CBN_13_Click()
If CBN_13.Value = True Then
CBO_13.Value = False
CBNA_13.Value = False
Range("ST_PSI") = 2
Else
Range("ST_PSI") = "X"
End If
End Sub
Private Sub CBNA_13_Click()
If CBNA_13.Value = True Then
CBN_13.Value = False
CBO_13.Value = False
Range("ST_PSI") = 0
Else
Range("ST_PSI") = "X"
End If
End Sub
Private Sub CBO_14_Click()

If CBO_14.Value = True Then

If Range("TOU_CT3") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If

TB_CBV.Visible = True
CBN_14.Value = False
CBNA_14.Value = False
Range("ST_CBV") = 1

Else
Range("ST_CBV") = "X"
TB_CBV.Visible = False

End If

End Sub
Private Sub CBN_14_Click()

If CBN_14.Value = True Then

If Range("TOU_CT3") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If

TB_CBV.Visible = False
CBO_14.Value = False
CBNA_14.Value = False
Range("ST_CBV") = 2

Else

Range("ST_CBV") = "X"

End If

End Sub
Private Sub CBNA_14_Click()
If CBNA_14.Value = True Then

If Range("TOU_CT3") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If

TB_CBV.Visible = False
CBN_14.Value = False
CBO_14.Value = False
Range("ST_CBV") = "X"

Else

Range("ST_CBV") = "X"

End If

End Sub
Private Sub CBO_15_Click()
If CBO_15.Value = True Then
If Range("TOU_CT3") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_15.Value = False
CBNA_15.Value = False
Range("ST_LOD") = 0
CBO_18.Value = True
Range("ST_AKH") = 1
Else
Range("ST_LOD") = "X"
CBO_18.Value = False
Range("ST_AKH") = "X"
End If
End Sub
Private Sub CBN_15_Click()
If CBN_15.Value = True Then
If Range("TOU_CT3") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_15.Value = False
CBNA_15.Value = False
Range("ST_LOD") = 2
Else
Range("ST_LOD") = "X"
End If
End Sub
Private Sub CBNA_15_Click()
If CBNA_15.Value = True Then
If Range("TOU_CT3") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBNA_18.Value = True
CBN_15.Value = False
CBO_15.Value = False
Range("ST_LOD") = 0
Else
Range("ST_LOD") = "X"
End If
End Sub
Private Sub CBO_16_Click()
If CBO_16.Value = True Then
CBN_16.Value = False
CBNA_16.Value = False
Range("ST_CAT") = 1
Else
Range("ST_CAT") = "X"
End If
End Sub
Private Sub CBN_16_Click()
If CBN_16.Value = True Then
CBO_16.Value = False
CBNA_16.Value = False
Range("ST_CAT") = 2
Else
Range("ST_CAT") = "X"
End If
End Sub
Private Sub CBNA_16_Click()
If CBNA_16.Value = True Then
CBN_16.Value = False
CBO_16.Value = False
Range("ST_CAT") = 0
Else
Range("ST_CAT") = "X"
End If
End Sub
Private Sub CBO_17_Click()
If CBO_17.Value = True Then
CBN_17.Value = False
CBNA_17.Value = False
Range("ST_CPM") = 1
Else
Range("ST_CPM") = "X"
End If
End Sub
Private Sub CBN_17_Click()
If CBN_17.Value = True Then
CBO_17.Value = False
CBNA_17.Value = False
Range("ST_CPM") = 2
Else
Range("ST_CPM") = "X"
End If
End Sub
Private Sub CBNA_17_Click()
If CBNA_17.Value = True Then
CBN_17.Value = False
CBO_17.Value = False
Range("ST_CPM") = 0
Else
Range("ST_CPM") = "X"
End If
End Sub
Private Sub CBO_18_Click()
If CBO_18.Value = True Then
TB_CAB.Visible = True
CBN_18.Value = False
CBNA_18.Value = False
Range("ST_AKH") = 1
Else
Range("ST_AKH") = "X"
TB_CAB.Visible = False
End If
End Sub
Private Sub CBN_18_Click()
If CBN_18.Value = True Then
TB_CAB.Visible = False
CBO_18.Value = False
CBNA_18.Value = False
Range("ST_AKH") = 2
Else
Range("ST_AKH") = "X"
End If
End Sub
Private Sub CBNA_18_Click()
If CBNA_18.Value = True Then
TB_CAB.Visible = False
CBN_18.Value = False
CBO_18.Value = False
Range("ST_AKH") = "X"
Else
Range("ST_AKH") = "X"
End If
End Sub
Private Sub CBO_19_Click()
If CBO_19.Value = True Then
If Range("TOU_CT4") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_19.Value = False
CBNA_19.Value = False
Range("ST_WW") = 0
Else
Range("ST_WW") = "X"
End If
End Sub
Private Sub CBN_19_Click()
If CBN_19.Value = True Then
If Range("TOU_CT4") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_19.Value = False
CBNA_19.Value = False
Range("ST_WW") = 2
Else
Range("ST_WW") = "X"
End If
End Sub
Private Sub CBNA_19_Click()
If CBNA_19.Value = True Then
If Range("TOU_CT4") = 1 And Range("ST_PKG_WW") = 2 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBNA_20.Value = True
Range("ST_CTL") = 0
CBNA_21.Value = True
Range("ST_CC") = 0
CBNA_22.Value = True
Range("ST_VDP") = 0
CBN_19.Value = False
CBO_19.Value = False
Range("ST_WW") = 0
Else
Range("ST_WW") = "X"
CBNA_20.Value = False
Range("ST_CTL") = "X"
CBNA_21.Value = False
Range("ST_CC") = "X"
CBNA_22.Value = False
Range("ST_VDP") = "X"
End If
End Sub
Private Sub CBO_20_Click()
If CBO_20.Value = True Then
If Range("TOU_CT4") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_20.Value = False
CBNA_20.Value = False
Range("ST_CTL") = 0
Else
Range("ST_CTL") = "X"
End If
End Sub
Private Sub CBN_20_Click()
If CBN_20.Value = True Then
If Range("TOU_CT4") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_20.Value = False
CBNA_20.Value = False
Range("ST_CTL") = 2
Else
Range("ST_CTL") = "X"
End If
End Sub
Private Sub CBNA_20_Click()
If CBNA_20.Value = True And Range("ST_PKG_WW") = 2 Then
If Range("TOU_CT4") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
Range("ST_CTL") = 0
Else
Range("ST_CTL") = "X"
End If
End Sub
Private Sub CBO_21_Click()
If CBO_21.Value = True Then
CBN_21.Value = False
CBNA_21.Value = False
Range("ST_CC") = 0
Else
Range("ST_CC") = "X"
End If
End Sub
Private Sub CBN_21_Click()
If CBN_21.Value = True Then
CBO_21.Value = False
CBNA_21.Value = False
Range("ST_CC") = 2
Else
Range("ST_CC") = "X"
End If
End Sub
Private Sub CBNA_21_Click()
If CBNA_21.Value = True Then
Range("ST_CC") = 0
Else
Range("ST_CC") = "X"
End If
End Sub
Private Sub CBO_22_Click()
If CBO_22.Value = True Then
CBN_22.Value = False
CBNA_22.Value = False
Range("ST_VDP") = 0
Else
Range("ST_VDP") = "X"
End If
End Sub
Private Sub CBN_22_Click()
If CBN_22.Value = True Then
CBO_22.Value = False
CBNA_22.Value = False
Range("ST_VDP") = 2
Else
Range("ST_VDP") = "X"
End If
End Sub
Private Sub CBNA_22_Click()
If CBNA_22.Value = True Then
Range("ST_VDP") = 0
Else
Range("ST_VDP") = "X"
End If
End Sub
Private Sub CBO_23_Click()
If CBO_23.Value = True Then
If Range("TOU_CT5") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_23.Value = False
CBNA_23.Value = False
Range("ST_4002") = 1
Else
Range("ST_4002") = "X"
End If
End Sub
Private Sub CBN_23_Click()
If CBN_23.Value = True Then
If Range("TOU_CT5") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_23.Value = False
CBNA_23.Value = False
Range("ST_4002") = 2
Else
Range("ST_4002") = "X"
End If
End Sub
Private Sub CBNA_23_Click()
If CBNA_23.Value = True Then
If Range("TOU_CT5") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_23.Value = False
CBO_23.Value = False
Range("ST_4002") = 0
Else
Range("ST_4002") = "X"
End If
End Sub
Private Sub CBO_24_Click()
If CBO_24.Value = True Then
If Range("TOU_CT5") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_24.Value = False
CBNA_24.Value = False
Range("ST_PPP2") = 1
Else
Range("ST_PPP2") = "X"
End If
End Sub
Private Sub CBN_24_Click()
If CBN_24.Value = True Then
If Range("TOU_CT5") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBO_24.Value = False
CBNA_24.Value = False
Range("ST_PPP2") = 2
Else
Range("ST_PPP2") = "X"
End If
End Sub
Private Sub CBNA_24_Click()
If CBNA_24.Value = True And Range("ST_PKG_LG") <> 1 Then
If Range("TOU_CT5") = 1 Then
MsgBox "ERREUR: L'encart précédent n'est pas entièrement complété."
End If
CBN_24.Value = False
CBO_24.Value = False
Range("ST_PPP2") = 0
Else
Range("ST_PPP2") = "X"
End If
End Sub
Private Sub CBO_25_Click()
If CBO_25.Value = True Then
CBN_25.Value = False
CBNA_25.Value = False
Range("ST_ZD2") = 1
Else
Range("ST_ZD2") = "X"
End If
End Sub
Private Sub CBN_25_Click()
If CBN_25.Value = True Then
CBO_25.Value = False
CBNA_25.Value = False
Range("ST_ZD2") = 2
Else
Range("ST_ZD2") = "X"
End If
End Sub
Private Sub CBNA_25_Click()
If CBNA_25.Value = True Then
CBN_25.Value = False
CBO_25.Value = False
Range("ST_ZD2") = 0
Else
Range("ST_ZD2") = "X"
End If
End Sub
Private Sub CBO_26_Click()
If CBO_26.Value = True Then
CBN_26.Value = False
CBNA_26.Value = False
Range("ST_REPO") = 1
Else
Range("ST_REPO") = "X"
End If
End Sub
Private Sub CBN_26_Click()
If CBN_26.Value = True Then
CBO_26.Value = False
CBNA_26.Value = False
Range("ST_REPO") = 2
Else
Range("ST_REPO") = "X"
End If
End Sub
Private Sub CBNA_26_Click()
If CBNA_26.Value = True Then
CBN_26.Value = False
CBO_26.Value = False
Range("ST_REPO") = 0
Else
Range("ST_REPO") = "X"
End If
End Sub
Private Sub CBO_27_Click()
If CBO_27.Value = True Then
CBN_27.Value = False
CBNA_27.Value = False
Range("ST_IPA2") = 1
Else
Range("ST_IPA2") = "X"
End If
End Sub
Private Sub CBN_27_Click()
If CBN_27.Value = True Then
CBO_27.Value = False
CBNA_27.Value = False
Range("ST_IPA2") = 2
Else
Range("ST_IPA2") = "X"
End If
End Sub
Private Sub CBNA_27_Click()
If CBNA_27.Value = True Then
CBN_27.Value = False
CBO_27.Value = False
Range("ST_IPA2") = 0
Else
Range("ST_IPA2") = "X"
End If
End Sub
