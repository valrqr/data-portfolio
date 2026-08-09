' ============================================
' Code VBA complet — Indicateurs_V2.1.1.xlsm
' 26 modules, extraits via olevba, nettoyés
' ============================================


' ============================================================
' MODULE : ADMIN.frm
' ============================================================

Private Sub Userfrom_Activate()

TB_ADMIN.Value = ""
TB_VRU.Value = ""
TB_MAJ.Value = ""

TB_ADMIN.Text = ""
TB_VRU.Text = ""
TB_MAJ.Text = ""

End Sub

Private Sub CB_ADM_F_Click()

ADMIN.Hide

End Sub

Private Sub CB_ADMIN_Click()
If TB_ADMIN.Text = "MAFALDA" Then

Sheets("SAISIE").Unprotect

Sheets("LOG_M").Visible = True
Sheets("LOG_ST").Visible = True
Sheets("LOG_ANNUEL").Visible = True
Sheets("LOG_ANNUEL_A-1").Visible = True
Sheets("DATA_ST").Visible = True
Sheets("DATA_ST3").Visible = True
Sheets("DATA_MAIN").Visible = True
Sheets("DATA_BASE").Visible = True
Sheets("LOG_OPS").Visible = True
Sheets("SYNTH").Visible = True
Sheets("CR").Visible = True
Sheets("CR_PP").Visible = True
Sheets("CON+ADD").Visible = True

ADMIN.Hide

Else
MsgBox ("Le mot de passe saisie est erroné")
End If
End Sub

Private Sub CB_MAJ_Click()

If TB_MAJ.Text = "B737_CONTENEURISE" Then

MAJ_DB1.Show

ADMIN.Hide

Else
MsgBox ("Le mot de passe saisie est erroné")
End If

End Sub
Private Sub CB_VRU_Click()

If TB_VRU.Text = "MAFALDA" Then

Sheets("SAISIE").Protect

Sheets("LOG_M").Visible = xlSheetVeryHidden
Sheets("LOG_ST").Visible = xlSheetVeryHidden
Sheets("LOG_ANNUEL").Visible = xlSheetVeryHidden
Sheets("LOG_ANNUEL_A-1").Visible = xlSheetVeryHidden
Sheets("DATA_ST").Visible = xlSheetVeryHidden
Sheets("DATA_ST3").Visible = xlSheetVeryHidden
Sheets("DATA_MAIN").Visible = xlSheetVeryHidden
Sheets("DATA_BASE").Visible = xlSheetVeryHidden
Sheets("LOG_OPS").Visible = xlSheetVeryHidden
Sheets("SYNTH").Visible = xlSheetVeryHidden
Sheets("CR").Visible = xlSheetVeryHidden
Sheets("CR_PP").Visible = xlSheetVeryHidden
Sheets("CON+ADD").Visible = xlSheetVeryHidden


ADMIN.Hide

Else

MsgBox ("Le mot de passe saisie est erroné")

End If

End Sub

' ============================================================
' MODULE : ANX.frm
' ============================================================

Private Sub CB_ANX_F_Click()
ANX.Hide
End Sub
Private Sub CB_CSL_Click()
CARBU.Show
End Sub
Private Sub CB_FNA_Click()
FAT.Show
End Sub
Private Sub CB_PERM_Click()
PERMIS.Show
End Sub

' ============================================================
' MODULE : CARBU.frm
' ============================================================

Private Sub CB_ANN_CAL_Click()

TB_CA_LF.BackColor = &H80000005
TB_CA_MF.BackColor = &H80000005
TB_CA_LG.BackColor = &H80000005
TB_CA_MG.BackColor = &H80000005
TB_CA_LS.BackColor = &H80000005
TB_CA_MS.BackColor = &H80000005
TB_CA_AS.BackColor = &H80000005
TB_CA_MAS.BackColor = &H80000005
TB_CA_LA.BackColor = &H80000005
TB_CA_MLA.BackColor = &H80000005

CARBU.Hide

End Sub

Private Sub CB_M_CARB_Change()

Range("MOIS_X") = CB_M_CARB.ListIndex + 1

TB_CA_LF.Text = Range("DMX_152")
TB_CA_MF.Text = Range("DMX_153")
TB_CA_LG.Text = Range("DMX_154")
TB_CA_MG.Text = Range("DMX_155")
TB_CA_LS.Text = Range("DMX_156")
TB_CA_MS.Text = Range("DMX_157")
TB_CA_AS.Text = Range("DMX_158")
TB_CA_MAS.Text = Range("DMX_159")
TB_CA_LA.Text = Range("DMX_160")
TB_CA_MLA.Text = Range("DMX_161")


End Sub
Private Sub CB_VAL_CAL_Click()

Dim ADD_LS, ADD_MS, ADD_LF, ADD_MF, ADD_LG, ADD_MG, ADD_AS, ADD_MAS, ADD_LA, ADD_MLA, ADD_152, ADD_153, ADD_154, ADD_155, ADD_156, ADD_157, ADD_158, ADD_159, ADD_160, ADD_161

ADD_LS = Range("ADD_MX_156").Value
ADD_MS = Range("ADD_MX_157").Value
ADD_LF = Range("ADD_MX_152").Value
ADD_MF = Range("ADD_MX_153").Value
ADD_LG = Range("ADD_MX_154").Value
ADD_MG = Range("ADD_MX_155").Value
ADD_AS = Range("ADD_MX_158").Value
ADD_MAS = Range("ADD_MX_159").Value
ADD_LA = Range("ADD_MX_160").Value
ADD_MLA = Range("ADD_MX_161").Value

ADD_156 = Range("ADD_MCR_156").Value
ADD_157 = Range("ADD_MCR_157").Value
ADD_152 = Range("ADD_MCR_152").Value
ADD_153 = Range("ADD_MCR_153").Value
ADD_154 = Range("ADD_MCR_154").Value
ADD_155 = Range("ADD_MCR_155").Value
ADD_158 = Range("ADD_MCR_158").Value
ADD_159 = Range("ADD_MCR_159").Value
ADD_160 = Range("ADD_MCR_160").Value
ADD_161 = Range("ADD_MCR_161").Value


If TB_CA_LF.Text = "" Or TB_CA_MF.Text = "" Or TB_CA_LG.Text = "" Or TB_CA_MG.Text = "" Or TB_CA_LS.Text = "" Or TB_CA_MS.Text = "" Or TB_CA_AS.Text = "" Or TB_CA_MAS.Text = "" Or TB_CA_LA.Text = "" Or TB_CA_MLA.Text = "" Then

MsgBox "Un ou plusieurs champs sont vides, complêtez pour valider"

Else

Sheets("LOG_ANNUEL").Range(ADD_LS) = TB_CA_LS.Value
Sheets("LOG_ANNUEL").Range(ADD_MS) = TB_CA_MS.Value
Sheets("LOG_ANNUEL").Range(ADD_LF) = TB_CA_LF.Value
Sheets("LOG_ANNUEL").Range(ADD_MF) = TB_CA_MF.Value
Sheets("LOG_ANNUEL").Range(ADD_LG) = TB_CA_LG.Value
Sheets("LOG_ANNUEL").Range(ADD_MG) = TB_CA_MG.Value
Sheets("LOG_ANNUEL").Range(ADD_AS) = TB_CA_AS.Value
Sheets("LOG_ANNUEL").Range(ADD_MAS) = TB_CA_MAS.Value
Sheets("LOG_ANNUEL").Range(ADD_LA) = TB_CA_LA.Value
Sheets("LOG_ANNUEL").Range(ADD_MLA) = TB_CA_MLA.Value

Sheets("LOG_M (2)").Range(ADD_156) = TB_CA_LS.Value
Sheets("LOG_M (2)").Range(ADD_157) = TB_CA_MS.Value
Sheets("LOG_M (2)").Range(ADD_152) = TB_CA_LF.Value
Sheets("LOG_M (2)").Range(ADD_153) = TB_CA_MF.Value
Sheets("LOG_M (2)").Range(ADD_154) = TB_CA_LG.Value
Sheets("LOG_M (2)").Range(ADD_155) = TB_CA_MG.Value
Sheets("LOG_M (2)").Range(ADD_158) = TB_CA_AS.Value
Sheets("LOG_M (2)").Range(ADD_159) = TB_CA_MAS.Value
Sheets("LOG_M (2)").Range(ADD_160) = TB_CA_LA.Value
Sheets("LOG_M (2)").Range(ADD_161) = TB_CA_MLA.Value


CARBU.Hide

End If

Range("ETS_CAR") = 1

ActiveWorkbook.Save

End Sub


Private Sub CommandButton1_Click()
Range("ESSAI_F244") = TB_CA_LF.Value
End Sub

Private Sub TB_CA_LF_Change()
If TB_CA_LF.Text = "" Then
TB_CA_LF.BackColor = &H8080FF
Else
TB_CA_LF.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MF_Change()
If TB_CA_MF.Text = "" Then
TB_CA_MF.BackColor = &H8080FF
Else
TB_CA_MF.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_LG_Change()
If TB_CA_LG.Text = "" Then
TB_CA_LG.BackColor = &H8080FF
Else
TB_CA_LG.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MG_Change()
If TB_CA_MG.Text = "" Then
TB_CA_MG.BackColor = &H8080FF
Else
TB_CA_MG.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_LS_Change()
If TB_CA_LS.Text = "" Then
TB_CA_LS.BackColor = &H8080FF
Else
TB_CA_LS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MS_Change()
If TB_CA_MS.Text = "" Then
TB_CA_MS.BackColor = &H8080FF
Else
TB_CA_MS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_AS_Change()
If TB_CA_AS.Text = "" Then
TB_CA_AS.BackColor = &H8080FF
Else
TB_CA_AS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MAS_Change()
If TB_CA_MAS.Text = "" Then
TB_CA_MAS.BackColor = &H8080FF
Else
TB_CA_MAS.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_LA_Change()
If TB_CA_LA.Text = "" Then
TB_CA_LA.BackColor = &H8080FF
Else
TB_CA_LA.BackColor = &H80000005
End If
End Sub

Private Sub TB_CA_MLA_Change()
If TB_CA_MLA.Text = "" Then
TB_CA_MLA.BackColor = &H8080FF
Else
TB_CA_MLA.BackColor = &H80000005
End If
End Sub

' ============================================================
' MODULE : CONSUL_MAIN.frm
' ============================================================

Private Sub CB_C_CLOSE_Click()

CONSUL_MAIN.Hide

ActiveWorkbook.Save

End Sub
Private Sub CB_C_ST_Click()

If Range("TOTM_ST").Value <> 0 Then

CONSUL_ST.Show

Else

MsgBox ("Aucune surveillance n'ayant été saisie, aucune donnée n'est actuellement disponible en consultation")

End If

End Sub
Private Sub CB_CRARC_Click()
CR.Show
End Sub

Private Sub CB_MODIF_C_Click()

If TB_ADM.Value = True Then

Dim AMX_114, AMX_115, AMX_116, AMX_117, AMX_118, AMX_119, AMX_124, AMX_126, AMX_127, AMX_128, AMX_129, AMX_130, AMX_131, AMX_132, AMX_137, AMX_195, AMX_196, AMX_197, AMX_198, AMX_199, AMX_200, AMX_201, AMX_206, AMX_FA1, AMX_FA2, AMX_FA3, AMX_FA4, AMX_FA5, AMX_I1, AMX_I3, AMX_I4, AMX_I5, AMX_I6, AMX_I7, AMX_I8, AMX_I9, AMX_I10, AMX_I11, AMX_EP1, AMX_EP2, AMX_EP3, AMX_EP4, AMX_D1, AMX_D2, AMX_D3, AMX_S1, AMX_IT1, AMX_IT2, AMX_IT3, AMX_IT4, AMX_FD1, AMX_FD2, AMX_FD3, AMX_F2, AMX_F3, AMX_F5, AMX_F6, AMX_F7, AMX_C2, AMX_C3, AMX_C4, AMX_C5, AMX_C6, AMX_C7, AMX_C8, AMX_C9, AMX_C10, AMX_C11, AMX_C12, AMX_DIV3, AMX_DIV4, AMX_DIV5, AMX_DIV6, AMX_DIV7, AMX_DIV8, AMX_DIV9, AMX_DIV10, AMX_DIV11, AMX_DIV12, AMX_DIV13, AMX_DIV14, AMX_MA1, AMX_MA2, AMX_MA3, AMX_MA4, AMX_MA5, AMX_MA6, AMX_MA7

AMX_113 = Range("ADD_MX_113").Value
AMX_114 = Range("ADD_MX_114").Value
AMX_115 = Range("ADD_MX_115").Value
AMX_116 = Range("ADD_MX_116").Value
AMX_117 = Range("ADD_MX_117").Value
AMX_118 = Range("ADD_MX_118").Value
AMX_119 = Range("ADD_MX_119").Value
AMX_124 = Range("ADD_MX_124").Value
AMX_126 = Range("ADD_MX_126").Value
AMX_127 = Range("ADD_MX_127").Value
AMX_128 = Range("ADD_MX_128").Value
AMX_129 = Range("ADD_MX_129").Value
AMX_130 = Range("ADD_MX_130").Value
AMX_131 = Range("ADD_MX_131").Value
AMX_132 = Range("ADD_MX_132").Value
AMX_137 = Range("ADD_MX_137").Value
AMX_195 = Range("ADD_MX_195").Value
AMX_196 = Range("ADD_MX_196").Value
AMX_197 = Range("ADD_MX_197").Value
AMX_198 = Range("ADD_MX_198").Value
AMX_199 = Range("ADD_MX_199").Value
AMX_200 = Range("ADD_MX_200").Value
AMX_201 = Range("ADD_MX_201").Value
AMX_206 = Range("ADD_MX_206").Value

AMX_I1 = Range("ADD_MX_2").Value
AMX_I3 = Range("ADD_MX_3").Value
AMX_I4 = Range("ADD_MX_4").Value
AMX_I5 = Range("ADD_MX_5").Value
AMX_I6 = Range("ADD_MX_6").Value
AMX_I7 = Range("ADD_MX_7").Value
AMX_I8 = Range("ADD_MX_8").Value
AMX_I9 = Range("ADD_MX_9").Value
AMX_I10 = Range("ADD_MX_10").Value
AMX_I11 = Range("ADD_MX_11").Value

AMX_EP1 = Range("ADD_MX_16").Value
AMX_EP2 = Range("ADD_MX_17").Value
AMX_EP3 = Range("ADD_MX_18").Value
AMX_EP4 = Range("ADD_MX_19").Value

AMX_D1 = Range("ADD_MX_24").Value
AMX_D2 = Range("ADD_MX_25").Value
AMX_D3 = Range("ADD_MX_111").Value

AMX_S1 = Range("ADD_MX_20").Value

AMX_IT1 = Range("ADD_MX_12").Value
AMX_IT2 = Range("ADD_MX_13").Value
AMX_IT3 = Range("ADD_MX_14").Value
AMX_IT4 = Range("ADD_MX_15").Value

AMX_FD1 = Range("ADD_MX_21").Value
AMX_FD2 = Range("ADD_MX_22").Value
AMX_FD3 = Range("ADD_MX_23").Value

AMX_F2 = Range("ADD_MX_26").Value
AMX_F3 = Range("ADD_MX_27").Value
AMX_F5 = Range("ADD_MX_28").Value
AMX_F6 = Range("ADD_MX_29").Value
AMX_F7 = Range("ADD_MX_30").Value

AMX_C2 = Range("ADD_MX_33").Value
AMX_C3 = Range("ADD_MX_34").Value
AMX_C4 = Range("ADD_MX_35").Value
AMX_C5 = Range("ADD_MX_36").Value
AMX_C6 = Range("ADD_MX_37").Value
AMX_C7 = Range("ADD_MX_38").Value
AMX_C8 = Range("ADD_MX_39").Value
AMX_C9 = Range("ADD_MX_32").Value
AMX_C10 = Range("ADD_MX_144").Value
AMX_C11 = Range("ADD_MX_42").Value
AMX_C12 = Range("ADD_MX_41").Value

AMX_DIV3 = Range("ADD_MX_43").Value
AMX_DIV4 = Range("ADD_MX_44").Value
AMX_DIV5 = Range("ADD_MX_45").Value
AMX_DIV6 = Range("ADD_MX_46").Value
AMX_DIV7 = Range("ADD_MX_47").Value
AMX_DIV8 = Range("ADD_MX_48").Value
AMX_DIV9 = Range("ADD_MX_49").Value
AMX_DIV10 = Range("ADD_MX_50").Value
AMX_DIV11 = Range("ADD_MX_51").Value
AMX_DIV12 = Range("ADD_MX_52").Value
AMX_DIV13 = Range("ADD_MX_53").Value
AMX_DIV14 = Range("ADD_MX_54").Value

AMX_MA1 = Range("ADD_MX_145").Value
AMX_MA2 = Range("ADD_MX_146").Value
AMX_MA3 = Range("ADD_MX_147").Value
AMX_MA4 = Range("ADD_MX_148").Value
AMX_MA5 = Range("ADD_MX_149").Value
AMX_MA6 = Range("ADD_MX_150").Value
AMX_MA7 = Range("ADD_MX_151").Value

AMX_FA1 = Range("ADD_MX_108").Value
AMX_FA2 = Range("ADD_MX_112").Value
AMX_FA3 = Range("ADD_MX_109").Value
AMX_FA4 = Range("ADD_MX_110").Value
AMX_FA5 = Range("ADD_MX_125").Value

Sheets("LOG_ANNUEL").Range(AMX_113) = TB_FH_1.Text
Sheets("LOG_ANNUEL").Range(AMX_115) = TB_FH_2.Text
Sheets("LOG_ANNUEL").Range(AMX_114) = TB_FH_4.Text
Sheets("LOG_ANNUEL").Range(AMX_117) = TB_FH_3.Text
Sheets("LOG_ANNUEL").Range(AMX_118) = TB_FH_5.Text
Sheets("LOG_ANNUEL").Range(AMX_116) = TB_FH_6.Text
Sheets("LOG_ANNUEL").Range(AMX_119) = TB_FH_7.Text
Sheets("LOG_ANNUEL").Range(AMX_124) = TB_FH_8.Text

Sheets("LOG_ANNUEL").Range(AMX_126) = TB_FHP_1.Text
Sheets("LOG_ANNUEL").Range(AMX_128) = TB_FHP_2.Text
Sheets("LOG_ANNUEL").Range(AMX_127) = TB_FHP_4.Text
Sheets("LOG_ANNUEL").Range(AMX_130) = TB_FHP_3.Text
Sheets("LOG_ANNUEL").Range(AMX_131) = TB_FHP_5.Text
Sheets("LOG_ANNUEL").Range(AMX_129) = TB_FHP_6.Text
Sheets("LOG_ANNUEL").Range(AMX_132) = TB_FHP_7.Text
Sheets("LOG_ANNUEL").Range(AMX_137) = TB_FHP_8.Text

Sheets("LOG_ANNUEL").Range(AMX_195) = TB_FHM_1.Text
Sheets("LOG_ANNUEL").Range(AMX_197) = TB_FHM_2.Text
Sheets("LOG_ANNUEL").Range(AMX_196) = TB_FHM_4.Text
Sheets("LOG_ANNUEL").Range(AMX_199) = TB_FHM_3.Text
Sheets("LOG_ANNUEL").Range(AMX_200) = TB_FHM_5.Text
Sheets("LOG_ANNUEL").Range(AMX_198) = TB_FHM_6.Text
Sheets("LOG_ANNUEL").Range(AMX_201) = TB_FHM_7.Text
Sheets("LOG_ANNUEL").Range(AMX_206) = TB_FHM_8.Text

Sheets("LOG_ANNUEL").Range(AMX_I1) = TB_I_1.Text
Sheets("LOG_ANNUEL").Range(AMX_I3) = TB_I_3.Text
Sheets("LOG_ANNUEL").Range(AMX_I4) = TB_I_4.Text
Sheets("LOG_ANNUEL").Range(AMX_I5) = TB_I_5.Text
Sheets("LOG_ANNUEL").Range(AMX_I6) = TB_I_6.Text
Sheets("LOG_ANNUEL").Range(AMX_I7) = TB_I_7.Text
Sheets("LOG_ANNUEL").Range(AMX_I8) = TB_I_8.Text
Sheets("LOG_ANNUEL").Range(AMX_I9) = TB_I_9.Text
Sheets("LOG_ANNUEL").Range(AMX_I10) = TB_I_10.Text
Sheets("LOG_ANNUEL").Range(AMX_I11) = TB_I_11.Text

Sheets("LOG_ANNUEL").Range(AMX_EP1) = TB_EP_1.Text
Sheets("LOG_ANNUEL").Range(AMX_EP2) = TB_EP_2.Text
Sheets("LOG_ANNUEL").Range(AMX_EP3) = TB_EP_3.Text
Sheets("LOG_ANNUEL").Range(AMX_EP4) = TB_EP_4.Text

Sheets("LOG_ANNUEL").Range(AMX_D1) = TB_D_1.Text
Sheets("LOG_ANNUEL").Range(AMX_D2) = TB_D_2.Text
Sheets("LOG_ANNUEL").Range(AMX_D3) = TB_FA_6.Text
Sheets("LOG_ANNUEL").Range(AMX_D4) = TB_D_4.Text

Sheets("LOG_ANNUEL").Range(AMX_S1) = TB_S_1.Text

Sheets("LOG_ANNUEL").Range(AMX_IT1) = TB_IT_1.Text
Sheets("LOG_ANNUEL").Range(AMX_IT2) = TB_IT_2.Text
Sheets("LOG_ANNUEL").Range(AMX_IT3) = TB_IT_3.Text
Sheets("LOG_ANNUEL").Range(AMX_IT4) = LB_IT_4.Text

Sheets("LOG_ANNUEL").Range(AMX_FD1) = TB_FD_1.Text
Sheets("LOG_ANNUEL").Range(AMX_FD2) = TB_FD_2.Text
Sheets("LOG_ANNUEL").Range(AMX_FD3) = TB_FD_3.Text

Sheets("LOG_ANNUEL").Range(AMX_F1) = TB_F_1.Text
Sheets("LOG_ANNUEL").Range(AMX_F2) = TB_F_2.Text
Sheets("LOG_ANNUEL").Range(AMX_F3) = TB_F_3.Text
Sheets("LOG_ANNUEL").Range(AMX_F5) = TB_F_5.Text
Sheets("LOG_ANNUEL").Range(AMX_F6) = TB_F_6.Text
Sheets("LOG_ANNUEL").Range(AMX_F7) = TB_F_7.Text
Sheets("LOG_ANNUEL").Range(AMX_F8) = TB_F_8.Text

Sheets("LOG_ANNUEL").Range(AMX_C2) = TB_C_2.Text
Sheets("LOG_ANNUEL").Range(AMX_C3) = TB_C_3.Text
Sheets("LOG_ANNUEL").Range(AMX_C4) = TB_C_4.Text
Sheets("LOG_ANNUEL").Range(AMX_C5) = TB_C_5.Text
Sheets("LOG_ANNUEL").Range(AMX_C6) = TB_C_6.Text
Sheets("LOG_ANNUEL").Range(AMX_C7) = TB_C_7.Text
Sheets("LOG_ANNUEL").Range(AMX_C8) = TB_C_8.Text
Sheets("LOG_ANNUEL").Range(AMX_C9) = TB_C_9.Text
Sheets("LOG_ANNUEL").Range(AMX_C10) = TB_C_10.Text
Sheets("LOG_ANNUEL").Range(AMX_C11) = TB_C_11.Text
Sheets("LOG_ANNUEL").Range(AMX_C12) = TB_C_12.Text

Sheets("LOG_ANNUEL").Range(AMX_DIV3) = TB_DIV_3.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV4) = TB_DIV_4.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV5) = TB_DIV_5.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV6) = TB_DIV_6.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV7) = TB_DIV_7.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV8) = TB_DIV_8.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV9) = TB_DIV_9.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV10) = TB_DIV_10.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV11) = TB_DIV_11.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV12) = TB_DIV_12.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV13) = TB_DIV_13.Text
Sheets("LOG_ANNUEL").Range(AMX_DIV14) = TB_DIV_14.Text

Sheets("LOG_ANNUEL").Range(AMX_MA1) = TB_MA_1.Text
Sheets("LOG_ANNUEL").Range(AMX_MA2) = TB_MA_2.Text
Sheets("LOG_ANNUEL").Range(AMX_MA3) = TB_MA_3.Text
Sheets("LOG_ANNUEL").Range(AMX_MA4) = TB_MA_4.Text
Sheets("LOG_ANNUEL").Range(AMX_MA5) = TB_MA_5.Text
Sheets("LOG_ANNUEL").Range(AMX_MA6) = TB_MA_6.Text
Sheets("LOG_ANNUEL").Range(AMX_MA7) = TB_MA_7.Text

Sheets("LOG_ANNUEL").Range(AMX_FA1) = TB_FA_1.Text
Sheets("LOG_ANNUEL").Range(AMX_FA2) = TB_FA_2.Text
Sheets("LOG_ANNUEL").Range(AMX_FA3) = TB_FA_3.Text
Sheets("LOG_ANNUEL").Range(AMX_FA4) = TB_FA_4.Text
Sheets("LOG_ANNUEL").Range(AMX_FA5) = TB_FA_5.Text

MsgBox ("Saisie enregistrée!")

Call CHECK_ALL_77

Else

Dim AJX_114, AJX_115, AJX_116, AJX_117, AJX_118, AJX_119, AJX_124, AJX_126, AJX_127, AJX_128, AJX_129, AJX_130, AJX_131, AJX_132, AJX_137, AJX_195, AJX_196, AJX_197, AJX_198, AJX_199, AJX_200, AJX_201, AJX_206, AJX_FA1, AJX_FA2, AJX_FA3, AJX_FA4, AJX_FA5, AJX_I1, AJX_I3, AJX_I4, AJX_I5, AJX_I6, AJX_I7, AJX_I8, AJX_I9, AJX_I10, AJX_I11, AJX_EP1, AJX_EP2, AJX_EP3, AJX_EP4, AJX_D1, AJX_D2, AJX_D3, AJX_S1, AJX_IT1, AJX_IT2, AJX_IT3, AJX_IT4, AJX_FD1, AJX_FD2, AJX_FD3, AJX_F2, AJX_F3, AJX_F5, AJX_F6, AJX_F7, AJX_F8, AJX_C2, AJX_C3, AJX_C4, AJX_C5, AJX_C6, AJX_C7, AJX_C8, AJX_C9, AJX_C10, AJX_C11, AJX_C12, AJX_DIV3, AJX_DIV4, AJX_DIV5, AJX_DIV6, AJX_DIV7, AJX_DIV8, AJX_DIV9, AJX_DIV10, AJX_DIV11, AJX_DIV12, AJX_DIV13, AJX_DIV14, AJX_DIV15, AJX_MA1, AJX_MA2, AJX_MA3, AJX_MA4, AJX_MA5, AJX_MA6, AJX_MA7

AJX_113 = Range("ADD_JX_113").Value
AJX_114 = Range("ADD_JX_114").Value
AJX_115 = Range("ADD_JX_115").Value
AJX_116 = Range("ADD_JX_116").Value
AJX_117 = Range("ADD_JX_117").Value
AJX_118 = Range("ADD_JX_118").Value
AJX_119 = Range("ADD_JX_119").Value
AJX_124 = Range("ADD_JX_124").Value
AJX_126 = Range("ADD_JX_126").Value
AJX_127 = Range("ADD_JX_127").Value
AJX_128 = Range("ADD_JX_128").Value
AJX_129 = Range("ADD_JX_129").Value
AJX_130 = Range("ADD_JX_130").Value
AJX_131 = Range("ADD_JX_131").Value
AJX_132 = Range("ADD_JX_132").Value
AJX_137 = Range("ADD_JX_137").Value
AJX_195 = Range("ADD_JX_195").Value
AJX_196 = Range("ADD_JX_196").Value
AJX_197 = Range("ADD_JX_197").Value
AJX_198 = Range("ADD_JX_198").Value
AJX_199 = Range("ADD_JX_199").Value
AJX_200 = Range("ADD_JX_200").Value
AJX_201 = Range("ADD_JX_201").Value
AJX_206 = Range("ADD_JX_206").Value

AJX_I1 = Range("ADD_JX_2").Value
AJX_I3 = Range("ADD_JX_3").Value
AJX_I4 = Range("ADD_JX_4").Value
AJX_I5 = Range("ADD_JX_5").Value
AJX_I6 = Range("ADD_JX_6").Value
AJX_I7 = Range("ADD_JX_7").Value
AJX_I8 = Range("ADD_JX_8").Value
AJX_I9 = Range("ADD_JX_9").Value
AJX_I10 = Range("ADD_JX_10").Value
AJX_I11 = Range("ADD_JX_11").Value

AJX_EP1 = Range("ADD_JX_16").Value
AJX_EP2 = Range("ADD_JX_17").Value
AJX_EP3 = Range("ADD_JX_18").Value
AJX_EP4 = Range("ADD_JX_19").Value

AJX_D1 = Range("ADD_JX_24").Value
AJX_D2 = Range("ADD_JX_25").Value
AJX_D3 = Range("ADD_JX_111").Value

AJX_S1 = Range("ADD_JX_20").Value

AJX_IT1 = Range("ADD_JX_12").Value
AJX_IT2 = Range("ADD_JX_13").Value
AJX_IT3 = Range("ADD_JX_14").Value
AJX_IT4 = Range("ADD_JX_15").Value

AJX_FD1 = Range("ADD_JX_21").Value
AJX_FD2 = Range("ADD_JX_22").Value
AJX_FD3 = Range("ADD_JX_23").Value

AJX_F2 = Range("ADD_JX_26").Value
AJX_F3 = Range("ADD_JX_27").Value
AJX_F5 = Range("ADD_JX_28").Value
AJX_F6 = Range("ADD_JX_29").Value
AJX_F7 = Range("ADD_JX_30").Value
AJX_F8 = Range("ADD_JX_31").Value

AJX_C2 = Range("ADD_JX_33").Value
AJX_C3 = Range("ADD_JX_34").Value
AJX_C4 = Range("ADD_JX_35").Value
AJX_C5 = Range("ADD_JX_36").Value
AJX_C6 = Range("ADD_JX_37").Value
AJX_C7 = Range("ADD_JX_38").Value
AJX_C8 = Range("ADD_JX_39").Value
AJX_C9 = Range("ADD_JX_32").Value
AJX_C10 = Range("ADD_JX_144").Value
AJX_C11 = Range("ADD_JX_42").Value
AJX_C12 = Range("ADD_JX_41").Value

AJX_DIV3 = Range("ADD_JX_43").Value
AJX_DIV4 = Range("ADD_JX_44").Value
AJX_DIV5 = Range("ADD_JX_45").Value
AJX_DIV6 = Range("ADD_JX_46").Value
AJX_DIV7 = Range("ADD_JX_47").Value
AJX_DIV8 = Range("ADD_JX_48").Value
AJX_DIV9 = Range("ADD_JX_49").Value
AJX_DIV10 = Range("ADD_JX_50").Value
AJX_DIV11 = Range("ADD_JX_51").Value
AJX_DIV12 = Range("ADD_JX_52").Value
AJX_DIV13 = Range("ADD_JX_53").Value
AJX_DIV14 = Range("ADD_JX_54").Value

AJX_MA1 = Range("ADD_JX_145").Value
AJX_MA2 = Range("ADD_JX_146").Value
AJX_MA3 = Range("ADD_JX_147").Value
AJX_MA4 = Range("ADD_JX_148").Value
AJX_MA5 = Range("ADD_JX_149").Value
AJX_MA6 = Range("ADD_JX_150").Value
AJX_MA7 = Range("ADD_JX_151").Value

AJX_FA1 = Range("ADD_JX_108").Value
AJX_FA2 = Range("ADD_JX_112").Value
AJX_FA3 = Range("ADD_JX_109").Value
AJX_FA4 = Range("ADD_JX_110").Value
AJX_FA5 = Range("ADD_JX_125").Value


Sheets("LOG_M").Range(AJX_113) = TB_FH_1.Text
Sheets("LOG_M").Range(AJX_115) = TB_FH_2.Text
Sheets("LOG_M").Range(AJX_114) = TB_FH_4.Text
Sheets("LOG_M").Range(AJX_117) = TB_FH_3.Text
Sheets("LOG_M").Range(AJX_118) = TB_FH_5.Text
Sheets("LOG_M").Range(AJX_116) = TB_FH_6.Text
Sheets("LOG_M").Range(AJX_119) = TB_FH_7.Text
Sheets("LOG_M").Range(AJX_124) = TB_FH_8.Text

Sheets("LOG_M").Range(AJX_126) = TB_FHP_1.Text
Sheets("LOG_M").Range(AJX_128) = TB_FHP_2.Text
Sheets("LOG_M").Range(AJX_127) = TB_FHP_4.Text
Sheets("LOG_M").Range(AJX_130) = TB_FHP_3.Text
Sheets("LOG_M").Range(AJX_131) = TB_FHP_5.Text
Sheets("LOG_M").Range(AJX_129) = TB_FHP_6.Text
Sheets("LOG_M").Range(AJX_132) = TB_FHP_7.Text
Sheets("LOG_M").Range(AJX_137) = TB_FHP_8.Text

Sheets("LOG_M").Range(AJX_195) = TB_FHM_1.Text
Sheets("LOG_M").Range(AJX_197) = TB_FHM_2.Text
Sheets("LOG_M").Range(AJX_196) = TB_FHM_4.Text
Sheets("LOG_M").Range(AJX_199) = TB_FHM_3.Text
Sheets("LOG_M").Range(AJX_200) = TB_FHM_5.Text
Sheets("LOG_M").Range(AJX_198) = TB_FHM_6.Text
Sheets("LOG_M").Range(AJX_201) = TB_FHM_7.Text
Sheets("LOG_M").Range(AJX_206) = TB_FHM_8.Text

Sheets("LOG_M").Range(AJX_I1) = TB_I_1.Text
Sheets("LOG_M").Range(AJX_I3) = TB_I_3.Text
Sheets("LOG_M").Range(AJX_I4) = TB_I_4.Text
Sheets("LOG_M").Range(AJX_I5) = TB_I_5.Text
Sheets("LOG_M").Range(AJX_I6) = TB_I_6.Text
Sheets("LOG_M").Range(AJX_I7) = TB_I_7.Text
Sheets("LOG_M").Range(AJX_I8) = TB_I_8.Text
Sheets("LOG_M").Range(AJX_I9) = TB_I_9.Text
Sheets("LOG_M").Range(AJX_I10) = TB_I_10.Text
Sheets("LOG_M").Range(AJX_I11) = TB_I_11.Text

Sheets("LOG_M").Range(AJX_EP1) = TB_EP_1.Text
Sheets("LOG_M").Range(AJX_EP2) = TB_EP_2.Text
Sheets("LOG_M").Range(AJX_EP3) = TB_EP_3.Text
Sheets("LOG_M").Range(AJX_EP4) = TB_EP_4.Text

Sheets("LOG_M").Range(AJX_D1) = TB_D_1.Text
Sheets("LOG_M").Range(AJX_D2) = TB_D_2.Text
Sheets("LOG_M").Range(AJX_D3) = TB_FA_6.Text
Sheets("LOG_M").Range(AJX_D4) = TB_D_4.Text

Sheets("LOG_M").Range(AJX_S1) = TB_S_1.Text

Sheets("LOG_M").Range(AJX_IT1) = TB_IT_1.Text
Sheets("LOG_M").Range(AJX_IT2) = TB_IT_2.Text
Sheets("LOG_M").Range(AJX_IT3) = TB_IT_3.Text
Sheets("LOG_M").Range(AJX_IT4) = LB_IT_4.Text

Sheets("LOG_M").Range(AJX_FD1) = TB_FD_1.Text
Sheets("LOG_M").Range(AJX_FD2) = TB_FD_2.Text
Sheets("LOG_M").Range(AJX_FD3) = TB_FD_3.Text

Sheets("LOG_M").Range(AJX_F1) = TB_F_1.Text
Sheets("LOG_M").Range(AJX_F2) = TB_F_2.Text
Sheets("LOG_M").Range(AJX_F3) = TB_F_3.Text
Sheets("LOG_M").Range(AJX_F5) = TB_F_5.Text
Sheets("LOG_M").Range(AJX_F6) = TB_F_6.Text
Sheets("LOG_M").Range(AJX_F7) = TB_F_7.Text
Sheets("LOG_M").Range(AJX_F8) = TB_F_8.Text

Sheets("LOG_M").Range(AJX_C2) = TB_C_2.Text
Sheets("LOG_M").Range(AJX_C3) = TB_C_3.Text
Sheets("LOG_M").Range(AJX_C4) = TB_C_4.Text
Sheets("LOG_M").Range(AJX_C5) = TB_C_5.Text
Sheets("LOG_M").Range(AJX_C6) = TB_C_6.Text
Sheets("LOG_M").Range(AJX_C7) = TB_C_7.Text
Sheets("LOG_M").Range(AJX_C8) = TB_C_8.Text
Sheets("LOG_M").Range(AJX_C9) = TB_C_9.Text
Sheets("LOG_M").Range(AJX_C10) = TB_C_10.Text
Sheets("LOG_M").Range(AJX_C11) = TB_C_11.Text
Sheets("LOG_M").Range(AJX_C12) = TB_C_12.Text

Sheets("LOG_M").Range(AJX_DIV3) = TB_DIV_3.Text
Sheets("LOG_M").Range(AJX_DIV4) = TB_DIV_4.Text
Sheets("LOG_M").Range(AJX_DIV5) = TB_DIV_5.Text
Sheets("LOG_M").Range(AJX_DIV6) = TB_DIV_6.Text
Sheets("LOG_M").Range(AJX_DIV7) = TB_DIV_7.Text
Sheets("LOG_M").Range(AJX_DIV8) = TB_DIV_8.Text
Sheets("LOG_M").Range(AJX_DIV9) = TB_DIV_9.Text
Sheets("LOG_M").Range(AJX_DIV10) = TB_DIV_10.Text
Sheets("LOG_M").Range(AJX_DIV11) = TB_DIV_11.Text
Sheets("LOG_M").Range(AJX_DIV12) = TB_DIV_12.Text
Sheets("LOG_M").Range(AJX_DIV13) = TB_DIV_13.Text
Sheets("LOG_M").Range(AJX_DIV14) = TB_DIV_14.Text

Sheets("LOG_M").Range(AJX_MA1) = TB_MA_1.Text
Sheets("LOG_M").Range(AJX_MA2) = TB_MA_2.Text
Sheets("LOG_M").Range(AJX_MA3) = TB_MA_3.Text
Sheets("LOG_M").Range(AJX_MA4) = TB_MA_4.Text
Sheets("LOG_M").Range(AJX_MA5) = TB_MA_5.Text
Sheets("LOG_M").Range(AJX_MA6) = TB_MA_6.Text
Sheets("LOG_M").Range(AJX_MA7) = TB_MA_7.Text

Sheets("LOG_M").Range(AJX_FA1) = TB_FA_1.Text
Sheets("LOG_M").Range(AJX_FA2) = TB_FA_2.Text
Sheets("LOG_M").Range(AJX_FA3) = TB_FA_3.Text
Sheets("LOG_M").Range(AJX_FA4) = TB_FA_4.Text
Sheets("LOG_M").Range(AJX_FA5) = TB_FA_5.Text

MsgBox ("Saisie enregistrée!")

Call CHECK_ALL_77

TB_ATC.Value = False

End If

ActiveWorkbook.Save

End Sub
Private Sub SpinButton1_Change()


With SpinButton1
.Max = Range("JOUR_LAST").Value
.Min = 1
.SmallChange = 1
End With

LB_JX.Caption = SpinButton1.Value
Range("JOUR_X").Value = SpinButton1.Value
Call CHECK_ALL_77

End Sub

Private Sub TB_ADM_Click()

If TB_ADM.Value = True Then


TBC_JAN.Visible = True
TBC_FEV.Visible = True
TBC_MAR.Visible = True
TBC_AVR.Visible = True
TBC_MAI.Visible = True
TBC_JUI.Visible = True
TBC_JUL.Visible = True
TBC_AOU.Visible = True
TBC_SEP.Visible = True
TBC_OCT.Visible = True
TBC_NOV.Visible = True
TBC_DEC.Visible = True

LX_MA.Visible = False
LB_JX.Visible = False

Range("JOUR_X").Value = ""

Else

TBC_JAN.Visible = False
TBC_FEV.Visible = False
TBC_MAR.Visible = False
TBC_AVR.Visible = False
TBC_MAI.Visible = False
TBC_JUI.Visible = False
TBC_JUL.Visible = False
TBC_AOU.Visible = False
TBC_SEP.Visible = False
TBC_OCT.Visible = False
TBC_NOV.Visible = False
TBC_DEC.Visible = False

LX_MA.Visible = True
LB_JX.Visible = True

Range("JOUR_X").Value = Range("JOUR_J").Value

End If

End Sub
Private Sub ONGOING_DATA()

Call VISI_ON

TB_I_1.Text = Range("TOT_1")
TB_I_2.Text = Range("TOT_2")
TB_I_3.Text = Range("TOT_3")
TB_I_4.Text = Range("TOT_4")
TB_I_5.Text = Range("TOT_5")
TB_I_6.Text = Range("TOT_6")
TB_I_7.Text = Range("TOT_7")
TB_I_8.Text = Range("TOT_8")
TB_I_9.Text = Range("TOT_9")
TB_I_10.Text = Range("TOT_10")
TB_I_11.Text = Range("TOT_11")

TB_EP_1.Text = Range("TOT_16")
TB_EP_2.Text = Range("TOT_17")
TB_EP_3.Text = Range("TOT_18")
TB_EP_4.Text = Range("TOT_19")

TB_D_1.Text = Range("TOT_24")
TB_D_2.Text = Range("TOT_25")
TB_FA_6.Text = Range("TOT_111")

TB_S_1.Text = Range("TOT_20")

TB_IT_1.Text = Range("TOT_12")
TB_IT_2.Text = Range("TOT_13")
TB_IT_3.Text = Range("TOT_14")
LB_IT_4.Text = "Contenu non totalisable"

TB_FD_1.Text = Range("TOT_21")
TB_FD_2.Text = Range("TOT_22")
TB_FD_3.Text = Range("TOT_23")

TB_F_1.Text = Range("TOT_26").Value + Range("TOT_27").Value
TB_F_2.Text = Range("TOT_26")
TB_F_3.Text = Range("TOT_27")
TB_F_5.Text = Range("TOT_28")
TB_F_6.Text = Range("TOT_29")
TB_F_7.Text = Range("TOT_30")
TB_F_8.Text = "Contenu non totalisable"

TB_C_1.Text = Range("TOT_41").Value + Range("TOT_42").Value
TB_C_2.Text = Range("TOT_33")
TB_C_3.Text = Range("TOT_34")
TB_C_4.Text = Range("TOT_35")
TB_C_5.Text = Range("TOT_36")
TB_C_6.Text = Range("TOT_37")
TB_C_7.Text = Range("TOT_38")
TB_C_8.Text = Range("TOT_39")
TB_C_9.Text = Range("TOT_32")
TB_C_10.Text = Range("TOT_144")
TB_C_11.Text = Range("TOT_42")
TB_C_12.Text = Range("TOT_41")

TB_DIV_1.Text = Range("TOT_43").Value + Range("TOT_44").Value + Range("TOT_45").Value + Range("TOT_46").Value
TB_DIV_2.Text = Range("TOT_47").Value + Range("TOT_48").Value + Range("TOT_49").Value + Range("TOT_50").Value + Range("TOT_51").Value + Range("TOT_52").Value + Range("TOT_53").Value + Range("TOT_54").Value
TB_DIV_3.Text = Range("TOT_43")
TB_DIV_4.Text = Range("TOT_44")
TB_DIV_5.Text = Range("TOT_45")
TB_DIV_6.Text = Range("TOT_46")
TB_DIV_7.Text = Range("TOT_47")
TB_DIV_8.Text = Range("TOT_48")
TB_DIV_9.Text = Range("TOT_49")
TB_DIV_10.Text = Range("TOT_50")
TB_DIV_11.Text = Range("TOT_51")
TB_DIV_12.Text = Range("TOT_52")
TB_DIV_13.Text = Range("TOT_53")
TB_DIV_14.Text = Range("TOT_54")
TB_DIV_15.Text = "Contenu non totalisable"

TB_T_2.Text = Range("TOT_69")
TB_T_3.Text = Range("TOT_82")
TB_T_4.Text = Range("TOT_56")
TB_T_6.Text = Range("TOT_95")

TB_T_9.Text = Range("TOT_183")
TB_T_10.Text = Range("TOT_184")
TB_T_11.Text = Range("TOT_186")
TB_T_12.Text = Range("TOT_185")
TB_T_13.Text = Range("TOT_188")
TB_T_14.Text = Range("TOT_187")
TB_T_15.Text = Range("TOT_189")
TB_T_16.Text = Range("TOT_191")
TB_T_17.Text = Range("TOT_193")
TB_T_18.Text = Range("TOT_192")
TB_T_19.Text = Range("TOT_190")
TB_T_20.Text = Range("TOT_194")

TB_FA_1.Text = Range("TOT_108")
TB_FA_2.Text = Range("TOT_112")
TB_FA_3.Text = Range("TOT_109")
TB_FA_4.Text = Range("TOT_110")
TB_FA_5.Text = Range("TOT_125")

TB_MA_1.Text = Range("TOT_145")
TB_MA_2.Text = Range("TOT_146")
TB_MA_3.Text = Range("TOT_148")
TB_MA_4.Text = Range("TOT_147")
TB_MA_5.Text = Range("TOT_149")
TB_MA_6.Text = Range("TOT_150")
TB_MA_7.Text = Range("TOT_151")

TB_FHP_1.Text = Range("TOT_126")
TB_FHP_2.Text = Range("TOT_128")
TB_FHP_3.Text = Range("TOT_130")
TB_FHP_4.Text = Range("TOT_127")
TB_FHP_5.Text = Range("TOT_131")
TB_FHP_6.Text = Range("TOT_129")
TB_FHP_7.Text = Range("TOT_132")
TB_FHP_8.Text = Range("TOT_137")

TB_FHM_1.Text = Range("TOT_195")
TB_FHM_2.Text = Range("TOT_197")
TB_FHM_3.Text = Range("TOT_199")
TB_FHM_4.Text = Range("TOT_196")
TB_FHM_5.Text = Range("TOT_200")
TB_FHM_6.Text = Range("TOT_198")
TB_FHM_7.Text = Range("TOT_201")
TB_FHM_8.Text = Range("TOT_206")

TB_FH_1.Text = Range("TOT_113")
TB_FH_2.Text = Range("TOT_115")
TB_FH_3.Text = Range("TOT_117")
TB_FH_4.Text = Range("TOT_114")
TB_FH_5.Text = Range("TOT_118")
TB_FH_6.Text = Range("TOT_116")
TB_FH_7.Text = Range("TOT_119")
TB_FH_8.Text = Range("TOT_124")

End Sub

Private Sub TB_CHECK_M_Click()

If TB_CHECK_M.Value = True Then


LB_JX.Caption = Range("JOUR_DATA").Value
LX_MA.Caption = Range("MOIS_N1").Value

Range("ACT_CHECK").Value = 1
Range("JOUR_C").Value = 1
Range("MOD_S_JX").Value = 1
Range("MOIS_C").Value = ""

Call CHECK_ALL_77

Else

Range("ACT_CHECK").Value = 0
LB_JX.Caption = Range("JOUR_DATA").Value
LX_MA.Caption = Range("MOIS_ANNEE").Value

Call CHECK_ALLSP_77

End If

End Sub
Private Sub TB_FA_2_Change()
Range("ADD_FACT_MAIN").Value = 2
End Sub

Private Sub TB_FA_3_Change()
Range("ADD_FACT_MAIN").Value = 3
End Sub

Private Sub TB_FA_5_Change()
Range("ADD_FACT_MAIN").Value = 5
End Sub

Private Sub TB_FA_6_Change()
Range("ADD_FACT_MAIN").Value = 6
End Sub

Private Sub TB_FH_10_Change()
Range("ADD_FACT_MAIN").Value = 9
End Sub

Private Sub TB_FH_11_Change()
Range("ADD_FACT_MAIN").Value = 8
End Sub

Private Sub TB_FH_12_Change()
Range("ADD_FACT_MAIN").Value = 11
End Sub

Private Sub TB_FH_2_Change()
Range("ADD_FACT_MAIN").Value = 3
End Sub

Private Sub TB_FH_3_Change()
Range("ADD_FACT_MAIN").Value = 5
End Sub

Private Sub TB_FH_4_Change()
Range("ADD_FACT_MAIN").Value = 2
End Sub

Private Sub TB_FH_5_Change()
Range("ADD_FACT_MAIN").Value = 6
End Sub

Private Sub TB_FH_6_Change()
Range("ADD_FACT_MAIN").Value = 4
End Sub

Private Sub TB_FH_7_Change()
Range("ADD_FACT_MAIN").Value = 7
End Sub

Private Sub TB_FH_8_Change()
Range("ADD_FACT_MAIN").Value = 12
End Sub

Private Sub TB_FH_9_Change()
Range("ADD_FACT_MAIN").Value = 10
End Sub

Private Sub TB_T_2_Change()

Range("ADD_TRA_MAIN").Value = 2

End Sub
Private Sub TB_T_2_AfterUpdate()

Dim JX_MNR

JX_MNR = Range("ADD_JX_TMAIN").Value
Sheets("LOG_M_OPS").Range(JX_MNR) = TB_T_2.Text

End Sub

Private Sub TB_T_3_Change()

Range("ADD_TRA_MAIN").Value = 3

End Sub
Private Sub TB_T_3_AfterUpdate()
Dim JX_MAE

JX_MAE = Range("ADD_JX_TMAIN").Value
Sheets("LOG_M_OPS").Range(JX_MAE) = TB_T_3.Text

End Sub
Private Sub TB_FHM_1_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 1
End Sub
Private Sub TB_FHM_2_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 2
End Sub
Private Sub TB_FHM_3_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 3
End Sub
Private Sub TB_FHM_4_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 4
End Sub
Private Sub TB_FHM_5_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 5
End Sub
Private Sub TB_FHM_6_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 6
End Sub
Private Sub TB_FHM_7_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 7
End Sub
Private Sub TB_FHM_8_Change()
Range("ADD_FACT_MAIN").Value = 3
Range("ADD_FACT_HDG").Value = 12
End Sub
Private Sub TB_T_4_Change()
Range("ADD_TRA_MAIN").Value = 1
End Sub
Private Sub TB_FHP_1_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 1
End Sub
Private Sub TB_FHP_2_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 2
End Sub
Private Sub TB_FHP_3_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 3
End Sub
Private Sub TB_FHP_4_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 4
End Sub
Private Sub TB_FHP_5_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 5
End Sub
Private Sub TB_FHP_6_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 6
End Sub
Private Sub TB_FHP_7_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 7
End Sub
Private Sub TB_FHP_8_Change()
Range("ADD_FACT_MAIN").Value = 5
Range("ADD_FACT_HDG").Value = 12
End Sub
Private Sub TB_FH_1_AfterUpdate()

Dim JX_FH1

JX_FH1 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH1) = TB_FH_1.Text

End Sub

Private Sub TB_FH_2_AfterUpdate()

Dim JX_FH2

JX_FH2 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH2) = TB_FH_2.Text

End Sub

Private Sub TB_FH_3_AfterUpdate()

Dim JX_FH3

JX_FH3 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH3) = TB_FH_3.Text

End Sub

Private Sub TB_FH_4_AfterUpdate()

Dim JX_FH4

JX_FH4 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH4) = TB_FH_4.Text

End Sub

Private Sub TB_FH_8_AfterUpdate()

Dim JX_FH8

JX_FH8 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH8) = TB_FH_8.Text

End Sub

Private Sub TB_FH_5_AfterUpdate()

Dim JX_FH5

JX_FH5 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH5) = TB_FH_5.Text

End Sub

Private Sub TB_FH_6_AfterUpdate()

Dim JX_FH6

JX_FH6 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH6) = TB_FH_6.Text

End Sub

Private Sub TB_FH_7_AfterUpdate()

Dim JX_FH7

JX_FH7 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FH7) = TB_FH_7.Text

End Sub


Private Sub TB_FHM_1_AfterUpdate()

Dim JX_FHM1

JX_FHM1 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM1) = TB_FHM_1.Text

End Sub

Private Sub TB_FHM_2_AfterUpdate()

Dim JX_FHM2

JX_FHM2 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM2) = TB_FHM_2.Text

End Sub

Private Sub TB_FHM_3_AfterUpdate()

Dim JX_FHM3

JX_FHM3 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM3) = TB_FHM_3.Text

End Sub

Private Sub TB_FHM_4_AfterUpdate()

Dim JX_FHM4

JX_FHM4 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM4) = TB_FHM_4.Text

End Sub

Private Sub TB_FHM_8_AfterUpdate()

Dim JX_FHM8

JX_FHM8 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM8) = TB_FHM_8.Text

End Sub

Private Sub TB_FHM_5_AfterUpdate()

Dim JX_FHM5

JX_FHM5 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM5) = TB_FHM_5.Text

End Sub

Private Sub TB_FHM_6_AfterUpdate()

Dim JX_FHM6

JX_FHM6 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM6) = TB_FHM_6.Text

End Sub

Private Sub TB_FHM_7_AfterUpdate()

Dim JX_FHM7

JX_FHM7 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHM7) = TB_FHM_7.Text

End Sub
Private Sub TB_FHP_1_AfterUpdate()

Dim JX_FHP1

JX_FHP1 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP1) = TB_FHP_1.Text

End Sub

Private Sub TB_FHP_2_AfterUpdate()

Dim JX_FHP2

JX_FHP2 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP2) = TB_FHP_2.Text

End Sub

Private Sub TB_FHP_3_AfterUpdate()

Dim JX_FHP3

JX_FHP3 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP3) = TB_FHP_3.Text

End Sub

Private Sub TB_FHP_4_AfterUpdate()

Dim JX_FHP4

JX_FHP4 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP4) = TB_FHP_4.Text

End Sub

Private Sub TB_FHP_8_AfterUpdate()

Dim JX_FHP8

JX_FHP8 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP8) = TB_FHP_8.Text

End Sub

Private Sub TB_FHP_5_AfterUpdate()

Dim JX_FHP5

JX_FHP5 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP5) = TB_FHP_5.Text

End Sub

Private Sub TB_FHP_6_AfterUpdate()

Dim JX_FHP6

JX_FHP6 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP6) = TB_FHP_6.Text

End Sub

Private Sub TB_FHP_7_AfterUpdate()

Dim JX_FHP7

JX_FHP7 = Range("ADD_JX_FACT").Value
Sheets("LOG_M_OPS").Range(JX_FHP7) = TB_FHP_7.Text

End Sub
Private Sub TB_T_4_AfterUpdate()

Dim JX_CNR

JX_CNR = Range("ADD_JX_TMAIN").Value
Sheets("LOG_M_OPS").Range(JX_CNR) = TB_T_4.Text

End Sub

Private Sub TB_T_6_Change()

Range("ADD_TRA_MAIN").Value = 4

End Sub
Private Sub TB_T_6_AfterUpdate()

Dim JX_MPA

JX_MPA = Range("ADD_JX_TMAIN").Value
Sheets("LOG_M_OPS").Range(JX_MPA) = TB_T_6.Text

End Sub

Private Sub TB_FH_1_Change()
Range("ADD_TRA_MAIN").Value = 2
Range("ADD_TRAF_HDG").Value = 1
End Sub

Private Sub Userform_Activate()

LB_JX.Caption = Range("JOUR_J").Value
LX_MA.Caption = Range("MOIS_ANNEE").Value

Range("JOUR_X").Value = Range("JOUR_J").Value
Range("JOUR_X").Value = Range("JOUR_J").Value
Range("MOIS_X").Value = ""

SpinButton1.Value = Range("JOUR_J").Value
LB_JX.Caption = SpinButton1.Value

Call CHECK_ALL_77

End Sub
Private Sub CHECK_ALL_77()

If Range("CK_1").Value = 1 Then
LB_I_1.Visible = True
TB_I_1.Visible = True
TB_I_1.Text = Range("DJX_1").Value
Else
LB_I_1.Visible = False
TB_I_1.Visible = False
TB_I_1.Text = ""
End If

If Range("CK_2").Value = 1 Then
LB_I_2.Visible = True
TB_I_2.Visible = True
TB_I_2.Text = Range("DJX_2").Value
Else
LB_I_2.Visible = False
TB_I_2.Visible = False
TB_I_2.Text = ""
End If

If Range("CK_3").Value = 1 Then
LB_I_3.Visible = True
TB_I_3.Visible = True
TB_I_3.Text = Range("DJX_3").Value
Else
LB_I_3.Visible = False
TB_I_3.Visible = False
TB_I_3.Text = ""
End If

If Range("CK_4").Value = 1 Then
LB_I_4.Visible = True
TB_I_4.Visible = True
TB_I_4.Text = Range("DJX_4").Value
Else
LB_I_4.Visible = False
TB_I_4.Visible = False
TB_I_4.Text = ""
End If

If Range("CK_5").Value = 1 Then
LB_I_5.Visible = True
TB_I_5.Visible = True
TB_I_5.Text = Range("DJX_5").Value
Else
LB_I_5.Visible = False
TB_I_5.Visible = False
TB_I_5.Text = ""
End If

If Range("CK_6").Value = 1 Then
LB_I_6.Visible = True
TB_I_6.Visible = True
TB_I_6.Text = Range("DJX_6").Value
Else
LB_I_6.Visible = False
TB_I_6.Visible = False
TB_I_6.Text = ""
End If

If Range("CK_7").Value = 1 Then
LB_I_7.Visible = True
TB_I_7.Visible = True
TB_I_7.Text = Range("DJX_7").Value
Else
LB_I_7.Visible = False
TB_I_7.Visible = False
TB_I_7.Text = ""
End If

If Range("CK_8").Value = 1 Then
LB_I_8.Visible = True
TB_I_8.Visible = True
TB_I_8.Text = Range("DJX_8").Value
Else
LB_I_8.Visible = False
TB_I_8.Visible = False
TB_I_8.Text = ""
End If

If Range("CK_9").Value = 1 Then
LB_I_9.Visible = True
TB_I_9.Visible = True
TB_I_9.Text = Range("DJX_9").Value
Else
LB_I_9.Visible = False
TB_I_9.Visible = False
TB_I_9.Text = ""
End If

If Range("CK_10").Value = 1 Then
LB_I_10.Visible = True
TB_I_10.Visible = True
TB_I_10.Text = Range("DJX_10").Value
Else
LB_I_10.Visible = False
TB_I_10.Visible = False
TB_I_10.Text = ""
End If

If Range("CK_11").Value = 1 Then
LB_I_11.Visible = True
TB_I_11.Visible = True
TB_I_11.Text = Range("DJX_11").Value
Else
LB_I_11.Visible = False
TB_I_11.Visible = False
TB_I_11.Text = ""
End If

If Range("CK_12").Value = 1 Then
LB_IT_1.Visible = True
TB_IT_1.Visible = True
TB_IT_1.Text = Range("DJX_12").Value
Else
LB_IT_1.Visible = False
TB_IT_1.Visible = False
TB_IT_1.Text = ""
End If

If Range("CK_13").Value = 1 Then
LB_IT_2.Visible = True
TB_IT_2.Visible = True
TB_IT_2.Text = Range("DJX_13").Value
Else
LB_IT_2.Visible = False
TB_IT_2.Visible = False
TB_IT_2.Text = ""
End If

If Range("CK_14").Value = 1 Then
LB_IT_3.Visible = True
TB_IT_3.Visible = True
TB_IT_3.Text = Range("DJX_14").Value
Else
LB_IT_3.Visible = False
TB_IT_3.Visible = False
TB_IT_3.Text = ""
End If

If Range("CK_15").Value = 1 Then
LB_IT_4.Visible = True
LB_IT_4.Text = Range("DJX_15").Value
Else
LB_IT_4.Visible = False
LB_IT_4.Text = ""
End If

If Range("CK_16").Value = 1 Then
LB_EP_1.Visible = True
TB_EP_1.Visible = True
TB_EP_1.Text = Range("DJX_16").Value
Else
LB_EP_1.Visible = False
TB_EP_1.Visible = False
TB_EP_1.Text = ""
End If

If Range("CK_17").Value = 1 Then
LB_EP_2.Visible = True
TB_EP_2.Visible = True
TB_EP_2.Text = Range("DJX_17").Value
Else
LB_EP_2.Visible = False
TB_EP_2.Visible = False
TB_EP_2.Text = ""
End If

If Range("CK_18").Value = 1 Then
LB_EP_3.Visible = True
TB_EP_3.Visible = True
TB_EP_3.Text = Range("DJX_18").Value
Else
LB_EP_3.Visible = False
TB_EP_3.Visible = False
TB_EP_3.Text = ""
End If

If Range("CK_19").Value = 1 Then
LB_EP_4.Visible = True
TB_EP_4.Visible = True
TB_EP_4.Text = Range("DJX_19").Value
Else
LB_EP_4.Visible = False
TB_EP_4.Visible = False
TB_EP_4.Text = ""
End If

If Range("CK_20").Value = 1 Then
LB_S_1.Visible = True
TB_S_1.Visible = True
TB_S_1.Text = Range("DJX_20").Value
Else
LB_S_1.Visible = False
TB_S_1.Visible = False
TB_S_1.Text = ""
End If

If Range("CK_21").Value = 1 Then
LB_FD_1.Visible = True
TB_FD_1.Visible = True
TB_FD_1.Text = Range("DJX_21").Value
Else
LB_FD_1.Visible = False
TB_FD_1.Visible = False
TB_FD_1.Text = ""
End If

If Range("CK_22").Value = 1 Then
LB_FD_2.Visible = True
TB_FD_2.Visible = True
TB_FD_2.Text = Range("DJX_22").Value
Else
LB_FD_2.Visible = False
TB_FD_2.Visible = False
TB_FD_2.Text = ""
End If

If Range("CK_23").Value = 1 Then
LB_FD_3.Visible = True
TB_FD_3.Visible = True
TB_FD_3.Text = Range("DJX_23").Value
Else
LB_FD_3.Visible = False
TB_FD_3.Visible = False
TB_FD_3.Text = ""
End If

If Range("CK_24").Value = 1 Then
LB_D_1.Visible = True
TB_D_1.Visible = True
TB_D_1.Text = Range("DJX_24").Value
Else
LB_D_1.Visible = False
TB_D_1.Visible = False
TB_D_1.Text = ""
End If

If Range("CK_26").Value = 1 Or Range("CK_27").Value = 1 Then
LB_F_1.Visible = True
LB_F_4.Visible = True
TB_F_1.Visible = True
TB_F_1.Text = Range("DJX_26").Value + Range("DJX_27").Value
Else
LB_F_1.Visible = False
LB_F_4.Visible = False
TB_F_1.Visible = False
TB_F_1.Text = ""
End If

If Range("CK_26").Value = 1 Then
LB_F_2.Visible = True
TB_F_2.Visible = True
TB_F_2.Text = Range("DJX_26").Value
Else
LB_F_2.Visible = False
TB_F_2.Visible = False
TB_F_2.Text = ""
End If

If Range("CK_27").Value = 1 Then
LB_F_3.Visible = True
TB_F_3.Visible = True
TB_F_3.Text = Range("DJX_27").Value
Else
LB_F_3.Visible = False
TB_F_3.Visible = False
TB_F_3.Text = ""
End If

If Range("CK_28").Value = 1 Then
LB_F_5.Visible = True
TB_F_5.Visible = True
TB_F_5.Text = Range("DJX_28").Value
Else
LB_F_5.Visible = False
TB_F_5.Visible = False
TB_F_5.Text = ""
End If

If Range("CK_29").Value = 1 Then
LB_F_6.Visible = True
TB_F_6.Visible = True
TB_F_6.Text = Range("DJX_29").Value
Else
LB_F_6.Visible = False
TB_F_6.Visible = False
TB_F_6.Text = ""
End If

If Range("CK_30").Value = 1 Then
LB_F_7.Visible = True
TB_F_7.Visible = True
TB_F_7.Text = Range("DJX_30").Value
Else
LB_F_7.Visible = False
TB_F_7.Visible = False
TB_F_7.Text = ""
End If

If Range("CK_31").Value = 1 Then
LB_F_8.Visible = True
TB_F_8.Visible = True
TB_F_8.Text = Range("DJX_31").Value
Else
LB_F_8.Visible = False
TB_F_8.Visible = False
TB_F_8.Text = ""
End If


If Range("CK_25").Value = 1 Then
LB_D_2.Visible = True
TB_D_2.Visible = True
TB_D_2.Text = Range("DJX_25").Value
Else
LB_D_2.Visible = False
TB_D_2.Visible = False
TB_D_2.Text = ""
End If

If Range("CK_111").Value = 1 Then
LB_FA_6.Visible = True
TB_FA_6.Visible = True
TB_FA_6.Text = Range("DJX_111").Value
Else
LB_FA_6.Visible = False
TB_FA_6.Visible = False
TB_FA_6.Text = ""
End If

If Range("CK_41").Value = 1 Then
LB_C_12.Visible = True
TB_C_12.Visible = True
TB_C_12.Text = Range("DJX_41").Value
Else
LB_C_12.Visible = False
TB_C_12.Visible = False
TB_C_12.Text = ""
End If

If Range("CK_41").Value = 1 Or Range("CK_42").Value = 1 Then
LB_C_1.Visible = True
TB_C_1.Visible = True
TB_C_1.Text = Range("DJX_41").Value + Range("DJX_42").Value
Else
LB_C_1.Visible = False
TB_C_1.Visible = False
TB_C_1.Text = ""
End If

If Range("CK_144").Value = 1 Then
LB_C_10.Visible = True
TB_C_10.Visible = True
TB_C_10.Text = Range("DJX_144").Value
Else
LB_C_10.Visible = False
TB_C_10.Visible = False
TB_C_10.Text = ""
End If

If Range("CK_32").Value = 1 Then
LB_C_2.Visible = True
TB_C_2.Visible = True
TB_C_2.Text = Range("DJX_32").Value
Else
LB_C_2.Visible = False
TB_C_2.Visible = False
TB_C_2.Text = ""
End If

If Range("CK_34").Value = 1 Then
LB_C_3.Visible = True
TB_C_3.Visible = True
TB_C_3.Text = Range("DJX_34").Value
Else
LB_C_3.Visible = False
TB_C_3.Visible = False
TB_C_3.Text = ""
End If

If Range("CK_35").Value = 1 Then
LB_C_4.Visible = True
TB_C_4.Visible = True
TB_C_4.Text = Range("DJX_35").Value
Else
LB_C_4.Visible = False
TB_C_4.Visible = False
TB_C_4.Text = ""
End If

If Range("CK_36").Value = 1 Then
LB_C_5.Visible = True
TB_C_5.Visible = True
TB_C_5.Text = Range("DJX_36").Value
Else
LB_C_5.Visible = False
TB_C_5.Visible = False
TB_C_5.Text = ""
End If

If Range("CK_37").Value = 1 Then
LB_C_6.Visible = True
TB_C_6.Visible = True
TB_C_6.Text = Range("DJX_37").Value
Else
LB_C_6.Visible = False
TB_C_6.Visible = False
TB_C_6.Text = ""
End If

If Range("CK_38").Value = 1 Then
LB_C_7.Visible = True
TB_C_7.Visible = True
TB_C_7.Text = Range("DJX_38").Value
Else
LB_C_7.Visible = False
TB_C_7.Visible = False
TB_C_7.Text = ""
End If

If Range("CK_39").Value = 1 Then
LB_C_8.Visible = True
TB_C_8.Visible = True
TB_C_8.Text = Range("DJX_39").Value
Else
LB_C_8.Visible = False
TB_C_8.Visible = False
TB_C_8.Text = ""
End If

If Range("CK_42").Value = 1 Then
LB_C_11.Visible = True
TB_C_11.Visible = True
TB_C_11.Text = Range("DJX_42").Value
Else
LB_C_11.Visible = False
TB_C_11.Visible = False
TB_C_11.Text = ""
End If

If Range("CK_32").Value = 1 Then
LB_C_9.Visible = True
TB_C_9.Visible = True
TB_C_9.Text = Range("DJX_32").Value
Else
LB_C_9.Visible = False
TB_C_9.Visible = False
TB_C_9.Text = ""
End If

If Range("CK_56").Value = 1 Then
LB_T_4.Visible = True
TB_T_4.Visible = True
TB_T_4.Text = Range("DJX_56").Value
Else
LB_T_4.Visible = False
TB_T_4.Visible = False
TB_T_4.Text = ""
End If


If Range("CK_69").Value = 1 Then
LB_T_2.Visible = True
TB_T_2.Visible = True
TB_T_2.Text = Range("DJX_69").Value
Else
LB_T_2.Visible = False
TB_T_2.Visible = False
TB_T_2.Text = ""
End If

If Range("CK_82").Value = 1 Then
LB_T_3.Visible = True
TB_T_3.Visible = True
TB_T_3.Text = Range("DJX_82").Value
Else
LB_T_3.Visible = False
TB_T_3.Visible = False
TB_T_3.Text = ""
End If

If Range("CK_95").Value = 1 Then
LB_T_6.Visible = True
TB_T_6.Visible = True
TB_T_6.Text = Range("DJX_95").Value
Else
LB_T_6.Visible = False
TB_T_6.Visible = False
TB_T_6.Text = ""
End If

If Range("CK_183").Value = 1 Then
LB_T_9.Visible = True
TB_T_9.Visible = True
TB_T_9.Text = Range("DJX_183").Value
Else
LB_T_9.Visible = False
TB_T_9.Visible = False
TB_T_9.Text = ""
End If

If Range("CK_184").Value = 1 Then
LB_T_10.Visible = True
TB_T_10.Visible = True
TB_T_10.Text = Range("DJX_184").Value
Else
LB_T_10.Visible = False
TB_T_10.Visible = False
TB_T_10.Text = ""
End If

If Range("CK_186").Value = 1 Then
LB_T_11.Visible = True
TB_T_11.Visible = True
TB_T_11.Text = Range("DJX_186").Value
Else
LB_T_11.Visible = False
TB_T_11.Visible = False
TB_T_11.Text = ""
End If

If Range("CK_185").Value = 1 Then
LB_T_12.Visible = True
TB_T_12.Visible = True
TB_T_12.Text = Range("DJX_185").Value
Else
LB_T_12.Visible = False
TB_T_12.Visible = False
TB_T_12.Text = ""
End If

If Range("CK_187").Value = 1 Then
LB_T_14.Visible = True
TB_T_14.Visible = True
TB_T_14.Text = Range("DJX_187").Value
Else
LB_T_14.Visible = False
TB_T_14.Visible = False
TB_T_14.Text = ""
End If

If Range("CK_188").Value = 1 Then
LB_T_13.Visible = True
TB_T_13.Visible = True
TB_T_13.Text = Range("DJX_188").Value
Else
LB_T_13.Visible = False
TB_T_13.Visible = False
TB_T_13.Text = ""
End If

If Range("CK_189").Value = 1 Then
LB_T_15.Visible = True
TB_T_15.Visible = True
TB_T_15.Text = Range("DJX_189").Value
Else
LB_T_15.Visible = False
TB_T_15.Visible = False
TB_T_15.Text = ""
End If

If Range("CK_191").Value = 1 Then
LB_T_16.Visible = True
TB_T_16.Visible = True
TB_T_16.Text = Range("DJX_191").Value
Else
LB_T_16.Visible = False
TB_T_16.Visible = False
TB_T_16.Text = ""
End If

If Range("CK_193").Value = 1 Then
LB_T_17.Visible = True
TB_T_17.Visible = True
TB_T_17.Text = Range("DJX_193").Value
Else
LB_T_17.Visible = False
TB_T_17.Visible = False
TB_T_17.Text = ""
End If

If Range("CK_192").Value = 1 Then
LB_T_18.Visible = True
TB_T_18.Visible = True
TB_T_18.Text = Range("DJX_192").Value
Else
LB_T_18.Visible = False
TB_T_18.Visible = False
TB_T_18.Text = ""
End If

If Range("CK_190").Value = 1 Then
LB_T_19.Visible = True
TB_T_19.Visible = True
TB_T_19.Text = Range("DJX_190").Value
Else
LB_T_19.Visible = False
TB_T_19.Visible = False
TB_T_19.Text = ""
End If

If Range("CK_194").Value = 1 Then
LB_T_20.Visible = True
TB_T_20.Visible = True
TB_T_20.Text = Range("DJX_194").Value
Else
LB_T_20.Visible = False
TB_T_20.Visible = False
TB_T_20.Text = ""
End If

If Range("CK_43").Value = 1 Or Range("CK_44").Value = 1 Or Range("CK_45").Value = 1 Or Range("CK_46").Value = 1 Then
LB_DIV_1.Visible = True
TB_DIV_1.Visible = True
TB_DIV_1.Text = Range("DJX_43").Value + Range("DJX_44").Value + Range("DJX_45").Value + Range("DJX_46").Value
Else
LB_DIV_1.Visible = False
TB_DIV_1.Visible = False
TB_DIV_1.Text = ""
End If

If Range("CK_47").Value = 1 Or Range("CK_48").Value = 1 Or Range("CK_49").Value = 1 Or Range("CK_50").Value = 1 Or Range("CK_51").Value = 1 Or Range("CK_52").Value = 1 Or Range("CK_53").Value = 1 Or Range("CK_54").Value = 1 Then
LB_DIV_2.Visible = True
TB_DIV_2.Visible = True
TB_DIV_2.Text = Range("DJX_47").Value + Range("DJX_48").Value + Range("DJX_49").Value + Range("DJX_50").Value + Range("DJX_51").Value + Range("DJX_52").Value + Range("DJX_53").Value + Range("DJX_54").Value
Else
LB_DIV_2.Visible = False
TB_DIV_2.Visible = False
TB_DIV_2.Text = ""
End If

If Range("CK_43").Value = 1 Then
LB_DIV_3.Visible = True
TB_DIV_3.Visible = True
TB_DIV_3.Text = Range("DJX_43").Value
Else
LB_DIV_3.Visible = False
TB_DIV_3.Visible = False
TB_DIV_3.Text = ""
End If

If Range("CK_44").Value = 1 Then
LB_DIV_4.Visible = True
TB_DIV_4.Visible = True
TB_DIV_4.Text = Range("DJX_44").Value
Else
LB_DIV_4.Visible = False
TB_DIV_4.Visible = False
TB_DIV_4.Text = ""
End If

If Range("CK_45").Value = 1 Then
LB_DIV_5.Visible = True
TB_DIV_5.Visible = True
TB_DIV_5.Text = Range("DJX_45").Value
Else
LB_DIV_5.Visible = False
TB_DIV_5.Visible = False
TB_DIV_5.Text = ""
End If

If Range("CK_46").Value = 1 Then
LB_DIV_6.Visible = True
TB_DIV_6.Visible = True
TB_DIV_6.Text = Range("DJX_46").Value
Else
LB_DIV_6.Visible = False
TB_DIV_6.Visible = False
TB_DIV_6.Text = ""
End If

If Range("CK_47").Value = 1 Then
LB_DIV_7.Visible = True
TB_DIV_7.Visible = True
TB_DIV_7.Text = Range("DJX_47").Value
Else
LB_DIV_7.Visible = False
TB_DIV_7.Visible = False
TB_DIV_7.Text = ""
End If

If Range("CK_48").Value = 1 Then
LB_DIV_8.Visible = True
TB_DIV_8.Visible = True
TB_DIV_8.Text = Range("DJX_48").Value
Else
LB_DIV_8.Visible = False
TB_DIV_8.Visible = False
TB_DIV_8.Text = ""
End If

If Range("CK_49").Value = 1 Then
LB_DIV_9.Visible = True
TB_DIV_9.Visible = True
TB_DIV_9.Text = Range("DJX_49").Value
Else
LB_DIV_9.Visible = False
TB_DIV_9.Visible = False
TB_DIV_9.Text = ""
End If

If Range("CK_50").Value = 1 Then
LB_DIV_10.Visible = True
TB_DIV_10.Visible = True
TB_DIV_10.Text = Range("DJX_50").Value
Else
LB_DIV_10.Visible = False
TB_DIV_10.Visible = False
TB_DIV_10.Text = ""
End If

If Range("CK_51").Value = 1 Then
LB_DIV_11.Visible = True
TB_DIV_11.Visible = True
TB_DIV_11.Text = Range("DJX_51").Value
Else
LB_DIV_11.Visible = False
TB_DIV_11.Visible = False
TB_DIV_11.Text = ""
End If

If Range("CK_52").Value = 1 Then
LB_DIV_12.Visible = True
TB_DIV_12.Visible = True
TB_DIV_12.Text = Range("DJX_52").Value
Else
LB_DIV_12.Visible = False
TB_DIV_12.Visible = False
TB_DIV_12.Text = ""
End If

If Range("CK_53").Value = 1 Then
LB_DIV_13.Visible = True
TB_DIV_13.Visible = True
TB_DIV_13.Text = Range("DJX_53").Value
Else
LB_DIV_13.Visible = False
TB_DIV_13.Visible = False
TB_DIV_13.Text = ""
End If

If Range("CK_54").Value = 1 Then
LB_DIV_14.Visible = True
TB_DIV_14.Visible = True
TB_DIV_14.Text = Range("DJX_54").Value
Else
LB_DIV_14.Visible = False
TB_DIV_14.Visible = False
TB_DIV_14.Text = ""
End If

If Range("CK_55").Value = 1 Then
LB_DIV_15.Visible = True
TB_DIV_15.Visible = True
TB_DIV_15.Text = Range("DJX_55").Value
Else
LB_DIV_15.Visible = False
TB_DIV_15.Visible = False
TB_DIV_15.Text = ""
End If

If Range("CK_145").Value = 1 Then
LB_MA_1.Visible = True
TB_MA_1.Visible = True
TB_MA_1.Text = Range("DJX_145").Value
Else
LB_MA_1.Visible = False
TB_MA_1.Visible = False
TB_MA_1.Text = ""
End If

If Range("CK_146").Value = 1 Then
LB_MA_2.Visible = True
TB_MA_2.Visible = True
TB_MA_2.Text = Range("DJX_146").Value
Else
LB_MA_2.Visible = False
TB_MA_2.Visible = False
TB_MA_2.Text = ""
End If

If Range("CK_148").Value = 1 Then
LB_MA_3.Visible = True
TB_MA_3.Visible = True
TB_MA_3.Text = Range("DJX_148").Value
Else
LB_MA_3.Visible = False
TB_MA_3.Visible = False
TB_MA_3.Text = ""
End If

If Range("CK_147").Value = 1 Then
LB_MA_4.Visible = True
TB_MA_4.Visible = True
TB_MA_4.Text = Range("DJX_147").Value
Else
LB_MA_4.Visible = False
TB_MA_4.Visible = False
TB_MA_4.Text = ""
End If

If Range("CK_149").Value = 1 Then
LB_MA_5.Visible = True
TB_MA_5.Visible = True
TB_MA_5.Text = Range("DJX_149").Value
LB_MA_6.Visible = True
TB_MA_6.Visible = True
TB_MA_6.Text = Range("DJX_150").Value
TB_MA_7.Visible = True
TB_MA_7.Text = Range("DJX_151").Value
Else
LB_MA_5.Visible = False
TB_MA_5.Visible = False
TB_MA_5.Text = ""
LB_MA_6.Visible = False
TB_MA_6.Visible = False
TB_MA_6.Text = ""
TB_MA_7.Visible = False
TB_MA_7.Text = ""
End If

If Range("CK_108").Value = 1 Then
LB_FA_1.Visible = True
TB_FA_1.Visible = True
TB_FA_1.Text = Range("DJX_108").Value
Else
LB_FA_1.Visible = False
TB_FA_1.Visible = False
TB_FA_1.Text = ""
End If

If Range("CK_112").Value = 1 Then
LB_FA_2.Visible = True
TB_FA_2.Visible = True
TB_FA_2.Text = Range("DJX_112").Value

TB_FH_1.Visible = True
TB_FH_2.Visible = True
TB_FH_3.Visible = True
TB_FH_4.Visible = True
TB_FH_5.Visible = True
TB_FH_6.Visible = True
TB_FH_7.Visible = True
TB_FH_8.Visible = True

TB_FH_1.Text = Range("DJX_113").Value
TB_FH_2.Text = Range("DJX_115").Value
TB_FH_3.Text = Range("DJX_117").Value
TB_FH_4.Text = Range("DJX_114").Value
TB_FH_5.Text = Range("DJX_118").Value
TB_FH_6.Text = Range("DJX_116").Value
TB_FH_7.Text = Range("DJX_119").Value
TB_FH_8.Text = Range("DJX_124").Value

Else

LB_FA_2.Visible = False
TB_FA_2.Visible = False
TB_FA_2.Text = ""

TB_FH_1.Visible = False
TB_FH_2.Visible = False
TB_FH_3.Visible = False
TB_FH_4.Visible = False
TB_FH_5.Visible = False
TB_FH_6.Visible = False
TB_FH_7.Visible = False
TB_FH_8.Visible = False

End If


If Range("CK_109").Value = 1 Then
LB_FA_3.Visible = True
TB_FA_3.Visible = True
TB_FA_3.Text = Range("DJX_109").Value

TB_FHM_1.Visible = True
TB_FHM_2.Visible = True
TB_FHM_3.Visible = True
TB_FHM_4.Visible = True
TB_FHM_5.Visible = True
TB_FHM_6.Visible = True
TB_FHM_7.Visible = True
TB_FHM_8.Visible = True

TB_FHM_1.Text = Range("DJX_195").Value
TB_FHM_2.Text = Range("DJX_197").Value
TB_FHM_3.Text = Range("DJX_199").Value
TB_FHM_4.Text = Range("DJX_196").Value
TB_FHM_5.Text = Range("DJX_200").Value
TB_FHM_6.Text = Range("DJX_198").Value
TB_FHM_7.Text = Range("DJX_201").Value
TB_FHM_8.Text = Range("DJX_206").Value

Else

TB_FHM_1.Visible = False
TB_FHM_2.Visible = False
TB_FHM_3.Visible = False
TB_FHM_4.Visible = False
TB_FHM_5.Visible = False
TB_FHM_6.Visible = False
TB_FHM_7.Visible = False
TB_FHM_8.Visible = False

LB_FA_3.Visible = False
TB_FA_3.Visible = False
TB_FA_3.Text = ""
End If

If Range("CK_110").Value = 1 Then
LB_FA_4.Visible = True
TB_FA_4.Visible = True
TB_FA_4.Text = Range("DJX_110").Value
Else
LB_FA_4.Visible = False
TB_FA_4.Visible = False
TB_FA_4.Text = ""
End If

If Range("CK_125").Value = 1 Then
LB_FA_5.Visible = True
TB_FA_5.Visible = True
TB_FA_5.Text = Range("DJX_125").Value

TB_FHP_1.Visible = True
TB_FHP_2.Visible = True
TB_FHP_3.Visible = True
TB_FHP_4.Visible = True
TB_FHP_5.Visible = True
TB_FHP_6.Visible = True
TB_FHP_7.Visible = True
TB_FHP_8.Visible = True

TB_FHP_1.Text = Range("DJX_126").Value
TB_FHP_2.Text = Range("DJX_128").Value
TB_FHP_3.Text = Range("DJX_130").Value
TB_FHP_4.Text = Range("DJX_127").Value
TB_FHP_5.Text = Range("DJX_131").Value
TB_FHP_6.Text = Range("DJX_129").Value
TB_FHP_7.Text = Range("DJX_132").Value
TB_FHP_8.Text = Range("DJX_137").Value

Else

TB_FHP_1.Visible = False
TB_FHP_2.Visible = False
TB_FHP_3.Visible = False
TB_FHP_4.Visible = False
TB_FHP_5.Visible = False
TB_FHP_6.Visible = False
TB_FHP_7.Visible = False
TB_FHP_8.Visible = False

LB_FA_5.Visible = False
TB_FA_5.Visible = False
TB_FA_5.Text = ""
End If

End Sub
Private Sub TB_ATC_Click()

If TB_ATC.Value = True Then

Call VISI_ON

Else

Call VISI_OFF

End If

End Sub
Private Sub VISI_ON()

LB_I_1.Visible = True
LB_I_2.Visible = True
LB_I_3.Visible = True
LB_I_4.Visible = True
LB_I_5.Visible = True
LB_I_6.Visible = True
LB_I_7.Visible = True
LB_I_8.Visible = True
LB_I_9.Visible = True
LB_I_10.Visible = True
LB_I_11.Visible = True

TB_I_1.Visible = True
TB_I_2.Visible = True
TB_I_3.Visible = True
TB_I_4.Visible = True
TB_I_5.Visible = True
TB_I_6.Visible = True
TB_I_7.Visible = True
TB_I_8.Visible = True
TB_I_9.Visible = True
TB_I_10.Visible = True
TB_I_11.Visible = True

LB_EP_1.Visible = True
LB_EP_2.Visible = True
LB_EP_3.Visible = True
LB_EP_4.Visible = True

TB_EP_1.Visible = True
TB_EP_2.Visible = True
TB_EP_3.Visible = True
TB_EP_4.Visible = True

LB_D_1.Visible = True
LB_D_2.Visible = True
LB_FA_6.Visible = True

TB_D_1.Visible = True
TB_D_2.Visible = True
TB_FA_6.Visible = True

LB_S_1.Visible = True

TB_S_1.Visible = True

LB_IT_1.Visible = True
LB_IT_2.Visible = True
LB_IT_3.Visible = True
LB_IT_4.Visible = True

TB_IT_1.Visible = True
TB_IT_2.Visible = True
TB_IT_3.Visible = True

LB_FD_1.Visible = True
LB_FD_2.Visible = True
LB_FD_3.Visible = True

TB_FD_1.Visible = True
TB_FD_2.Visible = True
TB_FD_3.Visible = True

LB_F_1.Visible = True
LB_F_2.Visible = True
LB_F_3.Visible = True
LB_F_4.Visible = True
LB_F_5.Visible = True
LB_F_6.Visible = True
LB_F_7.Visible = True
LB_F_8.Visible = True

TB_F_1.Visible = True
TB_F_2.Visible = True
TB_F_3.Visible = True
TB_F_5.Visible = True
TB_F_6.Visible = True
TB_F_7.Visible = True
TB_F_8.Visible = True

LB_C_1.Visible = True
LB_C_2.Visible = True
LB_C_3.Visible = True
LB_C_4.Visible = True
LB_C_5.Visible = True
LB_C_6.Visible = True
LB_C_7.Visible = True
LB_C_8.Visible = True
LB_C_9.Visible = True
LB_C_10.Visible = True
LB_C_11.Visible = True
LB_C_12.Visible = True

TB_C_1.Visible = True
TB_C_2.Visible = True
TB_C_3.Visible = True
TB_C_4.Visible = True
TB_C_5.Visible = True
TB_C_6.Visible = True
TB_C_7.Visible = True
TB_C_8.Visible = True
TB_C_9.Visible = True
TB_C_10.Visible = True
TB_C_11.Visible = True
TB_C_12.Visible = True

LB_DIV_1.Visible = True
LB_DIV_2.Visible = True
LB_DIV_3.Visible = True
LB_DIV_4.Visible = True
LB_DIV_5.Visible = True
LB_DIV_6.Visible = True
LB_DIV_7.Visible = True
LB_DIV_8.Visible = True
LB_DIV_9.Visible = True
LB_DIV_10.Visible = True
LB_DIV_11.Visible = True
LB_DIV_12.Visible = True
LB_DIV_13.Visible = True
LB_DIV_14.Visible = True
LB_DIV_15.Visible = True

TB_DIV_1.Visible = True
TB_DIV_2.Visible = True
TB_DIV_3.Visible = True
TB_DIV_4.Visible = True
TB_DIV_5.Visible = True
TB_DIV_6.Visible = True
TB_DIV_7.Visible = True
TB_DIV_8.Visible = True
TB_DIV_9.Visible = True
TB_DIV_10.Visible = True
TB_DIV_11.Visible = True
TB_DIV_12.Visible = True
TB_DIV_13.Visible = True
TB_DIV_14.Visible = True
TB_DIV_15.Visible = True

LB_T_2.Visible = True
LB_T_3.Visible = True
LB_T_4.Visible = True
LB_T_6.Visible = True
LB_T_9.Visible = True
LB_T_10.Visible = True
LB_T_11.Visible = True
LB_T_12.Visible = True
LB_T_13.Visible = True
LB_T_14.Visible = True
LB_T_15.Visible = True
LB_T_16.Visible = True
LB_T_17.Visible = True
LB_T_18.Visible = True
LB_T_19.Visible = True
LB_T_20.Visible = True


TB_T_2.Visible = True
TB_T_3.Visible = True
TB_T_4.Visible = True
TB_T_6.Visible = True
TB_T_9.Visible = True
TB_T_10.Visible = True
TB_T_11.Visible = True
TB_T_12.Visible = True
TB_T_13.Visible = True
TB_T_14.Visible = True
TB_T_15.Visible = True
TB_T_16.Visible = True
TB_T_17.Visible = True
TB_T_18.Visible = True
TB_T_19.Visible = True
TB_T_20.Visible = True

LB_MA_1.Visible = True
LB_MA_2.Visible = True
LB_MA_3.Visible = True
LB_MA_4.Visible = True
LB_MA_5.Visible = True
LB_MA_6.Visible = True

TB_MA_1.Visible = True
TB_MA_2.Visible = True
TB_MA_3.Visible = True
TB_MA_4.Visible = True
TB_MA_5.Visible = True
TB_MA_6.Visible = True
TB_MA_7.Visible = True

LB_FA_1.Visible = True
LB_FA_2.Visible = True
LB_FA_3.Visible = True
LB_FA_4.Visible = True
LB_FA_5.Visible = True

TB_FA_1.Visible = True
TB_FA_2.Visible = True
TB_FA_3.Visible = True
TB_FA_4.Visible = True
TB_FA_5.Visible = True


TB_FH_1.Visible = True
TB_FH_2.Visible = True
TB_FH_3.Visible = True
TB_FH_4.Visible = True
TB_FH_5.Visible = True
TB_FH_6.Visible = True
TB_FH_7.Visible = True
TB_FH_8.Visible = True

TB_FHM_1.Visible = True
TB_FHM_2.Visible = True
TB_FHM_3.Visible = True
TB_FHM_4.Visible = True
TB_FHM_5.Visible = True
TB_FHM_6.Visible = True
TB_FHM_7.Visible = True
TB_FHM_8.Visible = True

TB_FHP_1.Visible = True
TB_FHP_2.Visible = True
TB_FHP_3.Visible = True
TB_FHP_4.Visible = True
TB_FHP_5.Visible = True
TB_FHP_6.Visible = True
TB_FHP_7.Visible = True
TB_FHP_8.Visible = True

End Sub
Private Sub VISI_OFF()

LB_I_1.Visible = False
LB_I_2.Visible = False
LB_I_3.Visible = False
LB_I_4.Visible = False
LB_I_5.Visible = False
LB_I_6.Visible = False
LB_I_7.Visible = False
LB_I_8.Visible = False
LB_I_9.Visible = False
LB_I_10.Visible = False
LB_I_11.Visible = False

TB_I_1.Visible = False
TB_I_2.Visible = False
TB_I_3.Visible = False
TB_I_4.Visible = False
TB_I_5.Visible = False
TB_I_6.Visible = False
TB_I_7.Visible = False
TB_I_8.Visible = False
TB_I_9.Visible = False
TB_I_10.Visible = False
TB_I_11.Visible = False

LB_EP_1.Visible = False
LB_EP_2.Visible = False
LB_EP_3.Visible = False
LB_EP_4.Visible = False

TB_EP_1.Visible = False
TB_EP_2.Visible = False
TB_EP_3.Visible = False
TB_EP_4.Visible = False

LB_D_1.Visible = False
LB_D_2.Visible = False
LB_FA_6.Visible = False
LB_D_4.Visible = False

TB_D_1.Visible = False
TB_D_2.Visible = False
TB_FA_6.Visible = False
TB_D_4.Visible = False

LB_S_1.Visible = False

TB_S_1.Visible = False

LB_IT_1.Visible = False
LB_IT_2.Visible = False
LB_IT_3.Visible = False
LB_IT_4.Visible = False

TB_IT_1.Visible = False
TB_IT_2.Visible = False
TB_IT_3.Visible = False

LB_FD_1.Visible = False
LB_FD_2.Visible = False
LB_FD_3.Visible = False

TB_FD_1.Visible = False
TB_FD_2.Visible = False
TB_FD_3.Visible = False

LB_F_1.Visible = False
LB_F_2.Visible = False
LB_F_3.Visible = False
LB_F_4.Visible = False
LB_F_5.Visible = False
LB_F_6.Visible = False
LB_F_7.Visible = False
LB_F_8.Visible = False

TB_F_1.Visible = False
TB_F_2.Visible = False
TB_F_3.Visible = False
TB_F_5.Visible = False
TB_F_6.Visible = False
TB_F_7.Visible = False
TB_F_8.Visible = False

LB_C_1.Visible = False
LB_C_2.Visible = False
LB_C_3.Visible = False
LB_C_4.Visible = False
LB_C_5.Visible = False
LB_C_6.Visible = False
LB_C_7.Visible = False
LB_C_8.Visible = False
LB_C_9.Visible = False
LB_C_10.Visible = False
LB_C_11.Visible = False
LB_C_12.Visible = False

TB_C_1.Visible = False
TB_C_2.Visible = False
TB_C_3.Visible = False
TB_C_4.Visible = False
TB_C_5.Visible = False
TB_C_6.Visible = False
TB_C_7.Visible = False
TB_C_8.Visible = False
TB_C_9.Visible = False
TB_C_10.Visible = False
TB_C_11.Visible = False
TB_C_12.Visible = False

LB_DIV_1.Visible = False
LB_DIV_2.Visible = False
LB_DIV_3.Visible = False
LB_DIV_4.Visible = False
LB_DIV_5.Visible = False
LB_DIV_6.Visible = False
LB_DIV_7.Visible = False
LB_DIV_8.Visible = False
LB_DIV_9.Visible = False
LB_DIV_10.Visible = False
LB_DIV_11.Visible = False
LB_DIV_12.Visible = False
LB_DIV_13.Visible = False
LB_DIV_14.Visible = False
LB_DIV_15.Visible = False

TB_DIV_1.Visible = False
TB_DIV_2.Visible = False
TB_DIV_3.Visible = False
TB_DIV_4.Visible = False
TB_DIV_5.Visible = False
TB_DIV_6.Visible = False
TB_DIV_7.Visible = False
TB_DIV_8.Visible = False
TB_DIV_9.Visible = False
TB_DIV_10.Visible = False
TB_DIV_11.Visible = False
TB_DIV_12.Visible = False
TB_DIV_13.Visible = False
TB_DIV_14.Visible = False
TB_DIV_15.Visible = False


LB_T_2.Visible = False
LB_T_3.Visible = False
LB_T_4.Visible = False
LB_T_6.Visible = False
LB_T_9.Visible = False
LB_T_10.Visible = False
LB_T_11.Visible = False
LB_T_12.Visible = False
LB_T_13.Visible = False
LB_T_14.Visible = False
LB_T_15.Visible = False
LB_T_16.Visible = False
LB_T_17.Visible = False
LB_T_18.Visible = False
LB_T_19.Visible = False
LB_T_20.Visible = False


TB_T_2.Visible = False
TB_T_3.Visible = False
TB_T_4.Visible = False
TB_T_6.Visible = False
TB_T_9.Visible = False
TB_T_10.Visible = False
TB_T_11.Visible = False
TB_T_12.Visible = False
TB_T_13.Visible = False
TB_T_14.Visible = False
TB_T_15.Visible = False
TB_T_16.Visible = False
TB_T_17.Visible = False
TB_T_18.Visible = False
TB_T_19.Visible = False
TB_T_20.Visible = False

LB_MA_1.Visible = False
LB_MA_2.Visible = False
LB_MA_3.Visible = False
LB_MA_4.Visible = False
LB_MA_5.Visible = False
LB_MA_6.Visible = False

TB_MA_1.Visible = False
TB_MA_2.Visible = False
TB_MA_3.Visible = False
TB_MA_4.Visible = False
TB_MA_5.Visible = False
TB_MA_6.Visible = False
TB_MA_7.Visible = False

LB_FA_1.Visible = False
LB_FA_2.Visible = False
LB_FA_3.Visible = False
LB_FA_4.Visible = False
LB_FA_5.Visible = False

TB_FA_1.Visible = False
TB_FA_2.Visible = False
TB_FA_3.Visible = False
TB_FA_4.Visible = False
TB_FA_5.Visible = False


TB_FH_1.Visible = False
TB_FH_2.Visible = False
TB_FH_3.Visible = False
TB_FH_4.Visible = False
TB_FH_5.Visible = False
TB_FH_6.Visible = False
TB_FH_7.Visible = False
TB_FH_8.Visible = False

TB_FHM_1.Visible = False
TB_FHM_2.Visible = False
TB_FHM_3.Visible = False
TB_FHM_4.Visible = False
TB_FHM_5.Visible = False
TB_FHM_6.Visible = False
TB_FHM_7.Visible = False
TB_FHM_8.Visible = False

TB_FHP_1.Visible = False
TB_FHP_2.Visible = False
TB_FHP_3.Visible = False
TB_FHP_4.Visible = False
TB_FHP_5.Visible = False
TB_FHP_6.Visible = False
TB_FHP_7.Visible = False
TB_FHP_8.Visible = False

End Sub
Private Sub TBC_JAN_Click()

If TBC_JAN.Value = True And Range("MOIS_M").Value = 1 Then

Call ONGOING_DATA

ElseIf TBC_JAN.Value = True Then

Range("MOIS_X").Value = 1

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X").Value = ""

Call VISI_OFF

End If

End Sub
Private Sub TBC_FEV_Click()

If TBC_FEV.Value = True And Range("MOIS_M").Value = 2 Then

Call ONGOING_DATA

ElseIf TBC_FEV.Value = True Then

Range("MOIS_X").Value = 2

Call CON_MEN

TBC_JAN.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X").Value = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_MAR_Click()

If TBC_MAR.Value = True And Range("MOIS_M").Value = 3 Then

Call ONGOING_DATA

ElseIf TBC_MAR.Value = True Then

Range("MOIS_X").Value = 3

Call CON_MEN

TBC_FEV.Value = False
TBC_JAN.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_AVR_Click()

If TBC_AVR.Value = True And Range("MOIS_M").Value = 4 Then

Call ONGOING_DATA

ElseIf TBC_AVR.Value = True Then

Range("MOIS_X").Value = 4

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_JAN.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_MAI_Click()

If TBC_MAI.Value = True And Range("MOIS_M").Value = 5 Then

Call ONGOING_DATA

ElseIf TBC_MAI.Value = True Then

Range("MOIS_X").Value = 5

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_JAN.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_JUI_Click()

If TBC_JUI.Value = True And Range("MOIS_M").Value = 6 Then

Call ONGOING_DATA

ElseIf TBC_JUI.Value = True Then

Range("MOIS_X").Value = 6

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JAN.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_JUL_Click()

If TBC_JUL.Value = True And Range("MOIS_M").Value = 7 Then

Call ONGOING_DATA

ElseIf TBC_JUL.Value = True Then

Range("MOIS_X").Value = 7

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JAN.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_AOU_Click()

If TBC_AOU.Value = True And Range("MOIS_M").Value = 8 Then

Call ONGOING_DATA

ElseIf TBC_AOU.Value = True Then

Range("MOIS_X").Value = 8

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_JAN.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_SEP_Click()

If TBC_SEP.Value = True And Range("MOIS_M").Value = 9 Then

Call ONGOING_DATA

ElseIf TBC_SEP.Value = True Then

Range("MOIS_X").Value = 9

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_JAN.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_OCT_Click()

If TBC_OCT.Value = True And Range("MOIS_M").Value = 10 Then

Call ONGOING_DATA

ElseIf TBC_OCT.Value = True Then

Range("MOIS_X").Value = 10

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_JAN.Value = False
TBC_NOV.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_NOV_Click()

If TBC_NOV.Value = True And Range("MOIS_M").Value = 11 Then

Call ONGOING_DATA

ElseIf TBC_NOV.Value = True Then

Range("MOIS_X").Value = 11

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_JAN.Value = False
TBC_DEC.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub TBC_DEC_Click()

If TBC_DEC.Value = True And Range("MOIS_M").Value = 12 Then

Call ONGOING_DATA

ElseIf TBC_DEC.Value = True Then

Range("MOIS_X").Value = 12

Call CON_MEN

TBC_FEV.Value = False
TBC_MAR.Value = False
TBC_AVR.Value = False
TBC_MAI.Value = False
TBC_JUI.Value = False
TBC_JUL.Value = False
TBC_AOU.Value = False
TBC_SEP.Value = False
TBC_OCT.Value = False
TBC_NOV.Value = False
TBC_JAN.Value = False

Else

Range("MOIS_X") = ""
Call VISI_OFF

End If

End Sub
Private Sub CON_MEN()

If Range("CK_1").Value = 1 Then
LB_I_1.Visible = True
TB_I_1.Visible = True
TB_I_1.Text = Range("DMX_1").Value
Else
LB_I_1.Visible = False
TB_I_1.Visible = False
TB_I_1.Text = ""
End If

If Range("CK_2").Value = 1 Then
LB_I_2.Visible = True
TB_I_2.Visible = True
TB_I_2.Text = Range("DMX_2").Value
Else
LB_I_2.Visible = False
TB_I_2.Visible = False
TB_I_2.Text = ""
End If

If Range("CK_3").Value = 1 Then
LB_I_3.Visible = True
TB_I_3.Visible = True
TB_I_3.Text = Range("DMX_3").Value
Else
LB_I_3.Visible = False
TB_I_3.Visible = False
TB_I_3.Text = ""
End If

If Range("CK_4").Value = 1 Then
LB_I_4.Visible = True
TB_I_4.Visible = True
TB_I_4.Text = Range("DMX_4").Value
Else
LB_I_4.Visible = False
TB_I_4.Visible = False
TB_I_4.Text = ""
End If

If Range("CK_5").Value = 1 Then
LB_I_5.Visible = True
TB_I_5.Visible = True
TB_I_5.Text = Range("DMX_5").Value
Else
LB_I_5.Visible = False
TB_I_5.Visible = False
TB_I_5.Text = ""
End If

If Range("CK_6").Value = 1 Then
LB_I_6.Visible = True
TB_I_6.Visible = True
TB_I_6.Text = Range("DMX_6").Value
Else
LB_I_6.Visible = False
TB_I_6.Visible = False
TB_I_6.Text = ""
End If

If Range("CK_7").Value = 1 Then
LB_I_7.Visible = True
TB_I_7.Visible = True
TB_I_7.Text = Range("DMX_7").Value
Else
LB_I_7.Visible = False
TB_I_7.Visible = False
TB_I_7.Text = ""
End If

If Range("CK_8").Value = 1 Then
LB_I_8.Visible = True
TB_I_8.Visible = True
TB_I_8.Text = Range("DMX_8").Value
Else
LB_I_8.Visible = False
TB_I_8.Visible = False
TB_I_8.Text = ""
End If

If Range("CK_9").Value = 1 Then
LB_I_9.Visible = True
TB_I_9.Visible = True
TB_I_9.Text = Range("DMX_9").Value
Else
LB_I_9.Visible = False
TB_I_9.Visible = False
TB_I_9.Text = ""
End If

If Range("CK_10").Value = 1 Then
LB_I_10.Visible = True
TB_I_10.Visible = True
TB_I_10.Text = Range("DMX_10").Value
Else
LB_I_10.Visible = False
TB_I_10.Visible = False
TB_I_10.Text = ""
End If

If Range("CK_11").Value = 1 Then
LB_I_11.Visible = True
TB_I_11.Visible = True
TB_I_11.Text = Range("DMX_11").Value
Else
LB_I_11.Visible = False
TB_I_11.Visible = False
TB_I_11.Text = ""
End If

If Range("CK_12").Value = 1 Then
LB_IT_1.Visible = True
TB_IT_1.Visible = True
TB_IT_1.Text = Range("DMX_12").Value
Else
LB_IT_1.Visible = False
TB_IT_1.Visible = False
TB_IT_1.Text = ""
End If

If Range("CK_13").Value = 1 Then
LB_IT_2.Visible = True
TB_IT_2.Visible = True
TB_IT_2.Text = Range("DMX_13").Value
Else
LB_IT_2.Visible = False
TB_IT_2.Visible = False
TB_IT_2.Text = ""
End If

If Range("CK_14").Value = 1 Then
LB_IT_3.Visible = True
TB_IT_3.Visible = True
TB_IT_3.Text = Range("DMX_14").Value
Else
LB_IT_3.Visible = False
TB_IT_3.Visible = False
TB_IT_3.Text = ""
End If

If Range("CK_15").Value = 1 Then
LB_IT_4.Visible = True
LB_IT_4.Text = Range("DMX_15").Value
Else
LB_IT_4.Visible = False
LB_IT_4.Text = ""
End If

If Range("CK_16").Value = 1 Then
LB_EP_1.Visible = True
TB_EP_1.Visible = True
TB_EP_1.Text = Range("DMX_16").Value
Else
LB_EP_1.Visible = False
TB_EP_1.Visible = False
TB_EP_1.Text = ""
End If

If Range("CK_17").Value = 1 Then
LB_EP_2.Visible = True
TB_EP_2.Visible = True
TB_EP_2.Text = Range("DMX_17").Value
Else
LB_EP_2.Visible = False
TB_EP_2.Visible = False
TB_EP_2.Text = ""
End If

If Range("CK_18").Value = 1 Then
LB_EP_3.Visible = True
TB_EP_3.Visible = True
TB_EP_3.Text = Range("DMX_18").Value
Else
LB_EP_3.Visible = False
TB_EP_3.Visible = False
TB_EP_3.Text = ""
End If

If Range("CK_19").Value = 1 Then
LB_EP_4.Visible = True
TB_EP_4.Visible = True
TB_EP_4.Text = Range("DMX_19").Value
Else
LB_EP_4.Visible = False
TB_EP_4.Visible = False
TB_EP_4.Text = ""
End If

If Range("CK_20").Value = 1 Then
LB_S_1.Visible = True
TB_S_1.Visible = True
TB_S_1.Text = Range("DMX_20").Value
Else
LB_S_1.Visible = False
TB_S_1.Visible = False
TB_S_1.Text = ""
End If

If Range("CK_21").Value = 1 Then
LB_FD_1.Visible = True
TB_FD_1.Visible = True
TB_FD_1.Text = Range("DMX_21").Value
Else
LB_FD_1.Visible = False
TB_FD_1.Visible = False
TB_FD_1.Text = ""
End If

If Range("CK_22").Value = 1 Then
LB_FD_2.Visible = True
TB_FD_2.Visible = True
TB_FD_2.Text = Range("DMX_22").Value
Else
LB_FD_2.Visible = False
TB_FD_2.Visible = False
TB_FD_2.Text = ""
End If

If Range("CK_23").Value = 1 Then
LB_FD_3.Visible = True
TB_FD_3.Visible = True
TB_FD_3.Text = Range("DMX_23").Value
Else
LB_FD_3.Visible = False
TB_FD_3.Visible = False
TB_FD_3.Text = ""
End If

If Range("CK_24").Value = 1 Then
LB_D_1.Visible = True
TB_D_1.Visible = True
TB_D_1.Text = Range("DMX_24").Value
Else
LB_D_1.Visible = False
TB_D_1.Visible = False
TB_D_1.Text = ""
End If

If Range("CK_26").Value = 1 Or Range("CK_27").Value = 1 Then
LB_F_1.Visible = True
LB_F_4.Visible = True
TB_F_1.Visible = True
TB_F_1.Text = Range("DMX_26").Value + Range("DMX_27").Value
Else
LB_F_1.Visible = False
LB_F_4.Visible = False
TB_F_1.Visible = False
TB_F_1.Text = ""
End If

If Range("CK_26").Value = 1 Then
LB_F_2.Visible = True
TB_F_2.Visible = True
TB_F_2.Text = Range("DMX_26").Value
Else
LB_F_2.Visible = False
TB_F_2.Visible = False
TB_F_2.Text = ""
End If

If Range("CK_27").Value = 1 Then
LB_F_3.Visible = True
TB_F_3.Visible = True
TB_F_3.Text = Range("DMX_27").Value
Else
LB_F_3.Visible = False
TB_F_3.Visible = False
TB_F_3.Text = ""
End If

If Range("CK_28").Value = 1 Then
LB_F_5.Visible = True
TB_F_5.Visible = True
TB_F_5.Text = Range("DMX_28").Value
Else
LB_F_5.Visible = False
TB_F_5.Visible = False
TB_F_5.Text = ""
End If

If Range("CK_29").Value = 1 Then
LB_F_6.Visible = True
TB_F_6.Visible = True
TB_F_6.Text = Range("DMX_29").Value
Else
LB_F_6.Visible = False
TB_F_6.Visible = False
TB_F_6.Text = ""
End If

If Range("CK_30").Value = 1 Then
LB_F_7.Visible = True
TB_F_7.Visible = True
TB_F_7.Text = Range("DMX_30").Value
Else
LB_F_7.Visible = False
TB_F_7.Visible = False
TB_F_7.Text = ""
End If


If Range("CK_25").Value = 1 Then
LB_D_2.Visible = True
TB_D_2.Visible = True
TB_D_2.Text = Range("DMX_25").Value
Else
LB_D_2.Visible = False
TB_D_2.Visible = False
TB_D_2.Text = ""
End If

If Range("CK_111").Value = 1 Then
LB_FA_6.Visible = True
TB_FA_6.Visible = True
TB_FA_6.Text = Range("DMX_111").Value
Else
LB_FA_6.Visible = False
TB_FA_6.Visible = False
TB_FA_6.Text = ""
End If

If Range("CK_41").Value = 1 Then
LB_C_12.Visible = True
TB_C_12.Visible = True
TB_C_12.Text = Range("DMX_41").Value
Else
LB_C_12.Visible = False
TB_C_12.Visible = False
TB_C_12.Text = ""
End If

If Range("CK_41").Value = 1 Or Range("CK_42").Value = 1 Then
LB_C_1.Visible = True
TB_C_1.Visible = True
TB_C_1.Text = Range("DMX_41").Value + Range("DMX_42").Value
Else
LB_C_1.Visible = False
TB_C_1.Visible = False
TB_C_1.Text = ""
End If

If Range("CK_144").Value = 1 Then
LB_C_10.Visible = True
TB_C_10.Visible = True
TB_C_10.Text = Range("DMX_144").Value
Else
LB_C_10.Visible = False
TB_C_10.Visible = False
TB_C_10.Text = ""
End If

If Range("CK_32").Value = 1 Then
LB_C_2.Visible = True
TB_C_2.Visible = True
TB_C_2.Text = Range("DMX_32").Value
Else
LB_C_2.Visible = False
TB_C_2.Visible = False
TB_C_2.Text = ""
End If

If Range("CK_34").Value = 1 Then
LB_C_3.Visible = True
TB_C_3.Visible = True
TB_C_3.Text = Range("DMX_34").Value
Else
LB_C_3.Visible = False
TB_C_3.Visible = False
TB_C_3.Text = ""
End If

If Range("CK_35").Value = 1 Then
LB_C_4.Visible = True
TB_C_4.Visible = True
TB_C_4.Text = Range("DMX_35").Value
Else
LB_C_4.Visible = False
TB_C_4.Visible = False
TB_C_4.Text = ""
End If

If Range("CK_36").Value = 1 Then
LB_C_5.Visible = True
TB_C_5.Visible = True
TB_C_5.Text = Range("DMX_36").Value
Else
LB_C_5.Visible = False
TB_C_5.Visible = False
TB_C_5.Text = ""
End If

If Range("CK_37").Value = 1 Then
LB_C_6.Visible = True
TB_C_6.Visible = True
TB_C_6.Text = Range("DMX_37").Value
Else
LB_C_6.Visible = False
TB_C_6.Visible = False
TB_C_6.Text = ""
End If

If Range("CK_38").Value = 1 Then
LB_C_7.Visible = True
TB_C_7.Visible = True
TB_C_7.Text = Range("DMX_38").Value
Else
LB_C_7.Visible = False
TB_C_7.Visible = False
TB_C_7.Text = ""
End If

If Range("CK_39").Value = 1 Then
LB_C_8.Visible = True
TB_C_8.Visible = True
TB_C_8.Text = Range("DMX_39").Value
Else
LB_C_8.Visible = False
TB_C_8.Visible = False
TB_C_8.Text = ""
End If

If Range("CK_42").Value = 1 Then
LB_C_11.Visible = True
TB_C_11.Visible = True
TB_C_11.Text = Range("DMX_42").Value
Else
LB_C_11.Visible = False
TB_C_11.Visible = False
TB_C_11.Text = ""
End If

If Range("CK_32").Value = 1 Then
LB_C_9.Visible = True
TB_C_9.Visible = True
TB_C_9.Text = Range("DMX_32").Value
Else
LB_C_9.Visible = False
TB_C_9.Visible = False
TB_C_9.Text = ""
End If

If Range("CK_56").Value = 1 Then
LB_T_4.Visible = True
TB_T_4.Visible = True
TB_T_4.Text = Range("DMX_56").Value
Else
LB_T_4.Visible = False
TB_T_4.Visible = False
TB_T_4.Text = ""
End If


If Range("CK_69").Value = 1 Then
LB_T_2.Visible = True
TB_T_2.Visible = True
TB_T_2.Text = Range("DMX_69").Value
Else
LB_T_2.Visible = False
TB_T_2.Visible = False
TB_T_2.Text = ""
End If

If Range("CK_82").Value = 1 Then
LB_T_3.Visible = True
TB_T_3.Visible = True
TB_T_3.Text = Range("DMX_82").Value
Else
LB_T_3.Visible = False
TB_T_3.Visible = False
TB_T_3.Text = ""
End If

If Range("CK_95").Value = 1 Then
LB_T_6.Visible = True
TB_T_6.Visible = True
TB_T_6.Text = Range("DMX_95").Value
Else
LB_T_6.Visible = False
TB_T_6.Visible = False
TB_T_6.Text = ""
End If

If Range("CK_183").Value = 1 Then
LB_T_9.Visible = True
TB_T_9.Visible = True
TB_T_9.Text = Range("DMX_183").Value
Else
LB_T_9.Visible = False
TB_T_9.Visible = False
TB_T_9.Text = ""
End If

If Range("CK_184").Value = 1 Then
LB_T_10.Visible = True
TB_T_10.Visible = True
TB_T_10.Text = Range("DMX_184").Value
Else
LB_T_10.Visible = False
TB_T_10.Visible = False
TB_T_10.Text = ""
End If

If Range("CK_186").Value = 1 Then
LB_T_11.Visible = True
TB_T_11.Visible = True
TB_T_11.Text = Range("DMX_186").Value
Else
LB_T_11.Visible = False
TB_T_11.Visible = False
TB_T_11.Text = ""
End If

If Range("CK_185").Value = 1 Then
LB_T_12.Visible = True
TB_T_12.Visible = True
TB_T_12.Text = Range("DMX_185").Value
Else
LB_T_12.Visible = False
TB_T_12.Visible = False
TB_T_12.Text = ""
End If

If Range("CK_187").Value = 1 Then
LB_T_14.Visible = True
TB_T_14.Visible = True
TB_T_14.Text = Range("DMX_187").Value
Else
LB_T_14.Visible = False
TB_T_14.Visible = False
TB_T_14.Text = ""
End If

If Range("CK_188").Value = 1 Then
LB_T_13.Visible = True
TB_T_13.Visible = True
TB_T_13.Text = Range("DMX_188").Value
Else
LB_T_13.Visible = False
TB_T_13.Visible = False
TB_T_13.Text = ""
End If

If Range("CK_189").Value = 1 Then
LB_T_15.Visible = True
TB_T_15.Visible = True
TB_T_15.Text = Range("DMX_189").Value
Else
LB_T_15.Visible = False
TB_T_15.Visible = False
TB_T_15.Text = ""
End If

If Range("CK_191").Value = 1 Then
LB_T_16.Visible = True
TB_T_16.Visible = True
TB_T_16.Text = Range("DMX_191").Value
Else
LB_T_16.Visible = False
TB_T_16.Visible = False
TB_T_16.Text = ""
End If

If Range("CK_193").Value = 1 Then
LB_T_17.Visible = True
TB_T_17.Visible = True
TB_T_17.Text = Range("DMX_193").Value
Else
LB_T_17.Visible = False
TB_T_17.Visible = False
TB_T_17.Text = ""
End If

If Range("CK_192").Value = 1 Then
LB_T_18.Visible = True
TB_T_18.Visible = True
TB_T_18.Text = Range("DMX_192").Value
Else
LB_T_18.Visible = False
TB_T_18.Visible = False
TB_T_18.Text = ""
End If

If Range("CK_190").Value = 1 Then
LB_T_19.Visible = True
TB_T_19.Visible = True
TB_T_19.Text = Range("DMX_190").Value
Else
LB_T_19.Visible = False
TB_T_19.Visible = False
TB_T_19.Text = ""
End If

If Range("CK_194").Value = 1 Then
LB_T_20.Visible = True
TB_T_20.Visible = True
TB_T_20.Text = Range("DMX_194").Value
Else
LB_T_20.Visible = False
TB_T_20.Visible = False
TB_T_20.Text = ""
End If

If Range("CK_43").Value = 1 Or Range("CK_44").Value = 1 Or Range("CK_45").Value = 1 Or Range("CK_46").Value = 1 Then
LB_DIV_1.Visible = True
TB_DIV_1.Visible = True
TB_DIV_1.Text = Range("DMX_43").Value + Range("DMX_44").Value + Range("DMX_45").Value + Range("DMX_46").Value
Else
LB_DIV_1.Visible = False
TB_DIV_1.Visible = False
TB_DIV_1.Text = ""
End If

If Range("CK_47").Value = 1 Or Range("CK_48").Value = 1 Or Range("CK_49").Value = 1 Or Range("CK_50").Value = 1 Or Range("CK_51").Value = 1 Or Range("CK_52").Value = 1 Or Range("CK_53").Value = 1 Or Range("CK_54").Value = 1 Then
LB_DIV_2.Visible = True
TB_DIV_2.Visible = True
TB_DIV_2.Text = Range("DMX_47").Value + Range("DMX_48").Value + Range("DMX_49").Value + Range("DMX_50").Value + Range("DMX_51").Value + Range("DMX_52").Value + Range("DMX_53").Value + Range("DMX_54").Value
Else
LB_DIV_2.Visible = False
TB_DIV_2.Visible = False
TB_DIV_2.Text = ""
End If

If Range("CK_43").Value = 1 Then
LB_DIV_3.Visible = True
TB_DIV_3.Visible = True
TB_DIV_3.Text = Range("DMX_43").Value
Else
LB_DIV_3.Visible = False
TB_DIV_3.Visible = False
TB_DIV_3.Text = ""
End If

If Range("CK_44").Value = 1 Then
LB_DIV_4.Visible = True
TB_DIV_4.Visible = True
TB_DIV_4.Text = Range("DMX_44").Value
Else
LB_DIV_4.Visible = False
TB_DIV_4.Visible = False
TB_DIV_4.Text = ""
End If

If Range("CK_45").Value = 1 Then
LB_DIV_5.Visible = True
TB_DIV_5.Visible = True
TB_DIV_5.Text = Range("DMX_45").Value
Else
LB_DIV_5.Visible = False
TB_DIV_5.Visible = False
TB_DIV_5.Text = ""
End If

If Range("CK_46").Value = 1 Then
LB_DIV_6.Visible = True
TB_DIV_6.Visible = True
TB_DIV_6.Text = Range("DMX_46").Value
Else
LB_DIV_6.Visible = False
TB_DIV_6.Visible = False
TB_DIV_6.Text = ""
End If

If Range("CK_47").Value = 1 Then
LB_DIV_7.Visible = True
TB_DIV_7.Visible = True
TB_DIV_7.Text = Range("DMX_47").Value
Else
LB_DIV_7.Visible = False
TB_DIV_7.Visible = False
TB_DIV_7.Text = ""
End If

If Range("CK_48").Value = 1 Then
LB_DIV_8.Visible = True
TB_DIV_8.Visible = True
TB_DIV_8.Text = Range("DMX_48").Value
Else
LB_DIV_8.Visible = False
TB_DIV_8.Visible = False
TB_DIV_8.Text = ""
End If

If Range("CK_49").Value = 1 Then
LB_DIV_9.Visible = True
TB_DIV_9.Visible = True
TB_DIV_9.Text = Range("DMX_49").Value
Else
LB_DIV_9.Visible = False
TB_DIV_9.Visible = False
TB_DIV_9.Text = ""
End If

If Range("CK_50").Value = 1 Then
LB_DIV_10.Visible = True
TB_DIV_10.Visible = True
TB_DIV_10.Text = Range("DMX_50").Value
Else
LB_DIV_10.Visible = False
TB_DIV_10.Visible = False
TB_DIV_10.Text = ""
End If

If Range("CK_51").Value = 1 Then
LB_DIV_11.Visible = True
TB_DIV_11.Visible = True
TB_DIV_11.Text = Range("DMX_51").Value
Else
LB_DIV_11.Visible = False
TB_DIV_11.Visible = False
TB_DIV_11.Text = ""
End If

If Range("CK_52").Value = 1 Then
LB_DIV_12.Visible = True
TB_DIV_12.Visible = True
TB_DIV_12.Text = Range("DMX_52").Value
Else
LB_DIV_12.Visible = False
TB_DIV_12.Visible = False
TB_DIV_12.Text = ""
End If

If Range("CK_53").Value = 1 Then
LB_DIV_13.Visible = True
TB_DIV_13.Visible = True
TB_DIV_13.Text = Range("DMX_53").Value
Else
LB_DIV_13.Visible = False
TB_DIV_13.Visible = False
TB_DIV_13.Text = ""
End If

If Range("CK_54").Value = 1 Then
LB_DIV_14.Visible = True
TB_DIV_14.Visible = True
TB_DIV_14.Text = Range("DMX_54").Value
Else
LB_DIV_14.Visible = False
TB_DIV_14.Visible = False
TB_DIV_14.Text = ""
End If

If Range("CK_145").Value = 1 Then
LB_MA_1.Visible = True
TB_MA_1.Visible = True
TB_MA_1.Text = Range("DMX_145").Value
Else
LB_MA_1.Visible = False
TB_MA_1.Visible = False
TB_MA_1.Text = ""
End If

If Range("CK_146").Value = 1 Then
LB_MA_2.Visible = True
TB_MA_2.Visible = True
TB_MA_2.Text = Range("DMX_146").Value
Else
LB_MA_2.Visible = False
TB_MA_2.Visible = False
TB_MA_2.Text = ""
End If

If Range("CK_148").Value = 1 Then
LB_MA_3.Visible = True
TB_MA_3.Visible = True
TB_MA_3.Text = Range("DMX_148").Value
Else
LB_MA_3.Visible = False
TB_MA_3.Visible = False
TB_MA_3.Text = ""
End If

If Range("CK_147").Value = 1 Then
LB_MA_4.Visible = True
TB_MA_4.Visible = True
TB_MA_4.Text = Range("DMX_147").Value
Else
LB_MA_4.Visible = False
TB_MA_4.Visible = False
TB_MA_4.Text = ""
End If

If Range("CK_149").Value = 1 Then
LB_MA_5.Visible = True
TB_MA_5.Visible = True
TB_MA_5.Text = Range("DMX_149").Value
LB_MA_6.Visible = True
TB_MA_6.Visible = True
TB_MA_6.Text = Range("DMX_150").Value
TB_MA_7.Visible = True
TB_MA_7.Text = Range("DMX_151").Value
Else
LB_MA_5.Visible = False
TB_MA_5.Visible = False
TB_MA_5.Text = ""
LB_MA_6.Visible = False
TB_MA_6.Visible = False
TB_MA_6.Text = ""
TB_MA_7.Visible = False
TB_MA_7.Text = ""
End If

If Range("CK_108").Value = 1 Then
LB_FA_1.Visible = True
TB_FA_1.Visible = True
TB_FA_1.Text = Range("DMX_108").Value
Else
LB_FA_1.Visible = False
TB_FA_1.Visible = False
TB_FA_1.Text = ""
End If

If Range("CK_112").Value = 1 Then
LB_FA_2.Visible = True
TB_FA_2.Visible = True
TB_FA_2.Text = Range("DMX_112").Value

TB_FH_1.Visible = True
TB_FH_2.Visible = True
TB_FH_3.Visible = True
TB_FH_4.Visible = True
TB_FH_5.Visible = True
TB_FH_6.Visible = True
TB_FH_7.Visible = True
TB_FH_8.Visible = True

TB_FH_1.Text = Range("DMX_113").Value
TB_FH_2.Text = Range("DMX_115").Value
TB_FH_3.Text = Range("DMX_117").Value
TB_FH_4.Text = Range("DMX_114").Value
TB_FH_5.Text = Range("DMX_118").Value
TB_FH_6.Text = Range("DMX_116").Value
TB_FH_7.Text = Range("DMX_119").Value
TB_FH_8.Text = Range("DMX_124").Value

Else

LB_FA_2.Visible = False
TB_FA_2.Visible = False
TB_FA_2.Text = ""

TB_FH_1.Visible = False
TB_FH_2.Visible = False
TB_FH_3.Visible = False
TB_FH_4.Visible = False
TB_FH_5.Visible = False
TB_FH_6.Visible = False
TB_FH_7.Visible = False
TB_FH_8.Visible = False

End If


If Range("CK_109").Value = 1 Then
LB_FA_3.Visible = True
TB_FA_3.Visible = True
TB_FA_3.Text = Range("DMX_109").Value

TB_FHM_1.Visible = True
TB_FHM_2.Visible = True
TB_FHM_3.Visible = True
TB_FHM_4.Visible = True
TB_FHM_5.Visible = True
TB_FHM_6.Visible = True
TB_FHM_7.Visible = True
TB_FHM_8.Visible = True

TB_FHM_1.Text = Range("DMX_195").Value
TB_FHM_2.Text = Range("DMX_197").Value
TB_FHM_3.Text = Range("DMX_199").Value
TB_FHM_4.Text = Range("DMX_196").Value
TB_FHM_5.Text = Range("DMX_200").Value
TB_FHM_6.Text = Range("DMX_198").Value
TB_FHM_7.Text = Range("DMX_201").Value
TB_FHM_8.Text = Range("DMX_206").Value

Else

TB_FHM_1.Visible = False
TB_FHM_2.Visible = False
TB_FHM_3.Visible = False
TB_FHM_4.Visible = False
TB_FHM_5.Visible = False
TB_FHM_6.Visible = False
TB_FHM_7.Visible = False
TB_FHM_8.Visible = False

LB_FA_3.Visible = False
TB_FA_3.Visible = False
TB_FA_3.Text = ""
End If

If Range("CK_110").Value = 1 Then
LB_FA_4.Visible = True
TB_FA_4.Visible = True
TB_FA_4.Text = Range("DMX_110").Value
Else
LB_FA_4.Visible = False
TB_FA_4.Visible = False
TB_FA_4.Text = ""
End If

If Range("CK_125").Value = 1 Then
LB_FA_5.Visible = True
TB_FA_5.Visible = True
TB_FA_5.Text = Range("DMX_125").Value

TB_FHP_1.Visible = True
TB_FHP_2.Visible = True
TB_FHP_3.Visible = True
TB_FHP_4.Visible = True
TB_FHP_5.Visible = True
TB_FHP_6.Visible = True
TB_FHP_7.Visible = True
TB_FHP_8.Visible = True

TB_FHP_1.Text = Range("DMX_126").Value
TB_FHP_2.Text = Range("DMX_128").Value
TB_FHP_3.Text = Range("DMX_130").Value
TB_FHP_4.Text = Range("DMX_127").Value
TB_FHP_5.Text = Range("DMX_131").Value
TB_FHP_6.Text = Range("DMX_129").Value
TB_FHP_7.Text = Range("DMX_132").Value
TB_FHP_8.Text = Range("DMX_137").Value

Else

TB_FHP_1.Visible = False
TB_FHP_2.Visible = False
TB_FHP_3.Visible = False
TB_FHP_4.Visible = False
TB_FHP_5.Visible = False
TB_FHP_6.Visible = False
TB_FHP_7.Visible = False
TB_FHP_8.Visible = False

LB_FA_5.Visible = False
TB_FA_5.Visible = False
TB_FA_5.Text = ""
End If
End Sub

' ============================================================
' MODULE : CONVOY.frm
' ============================================================

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

' ============================================================
' MODULE : COP.frm
' ============================================================

Private Sub CB_MJCOP_CNX_Click()
COP.Hide
End Sub

Private Sub CB_MJCOP_VAL_Click()

Range("COP_1") = TB_COP_1.Value
Range("COP_2") = TB_COP_2.Value
Range("COP_3") = TB_COP_3.Value
Range("COP_4") = TB_COP_4.Value
Range("COP_5") = TB_COP_5.Value
Range("COP_6") = TB_COP_6.Value
Range("COP_7") = TB_COP_7.Value
Range("COP_8") = TB_COP_8.Value
Range("COP_9") = TB_COP_9.Value
Range("COP_10") = TB_COP_10.Value
Range("COP_11") = TB_COP_11.Value
Range("COP_12") = TB_COP_12.Value
Range("COP_13") = TB_COP_13.Value
Range("COP_14") = TB_COP_14.Value
Range("COP_15") = TB_COP_15.Value

Sheets("DATA_BASE").Visible = True

Range("H3:H17").Select
    ActiveWorkbook.Worksheets("DATA_BASE").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DATA_BASE").Sort.SortFields.Add Key:=Range("H3"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DATA_BASE").Sort
        .SetRange Range("H3:H17")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Sheets("DATA_BASE").Visible = xlSheetVeryHidden

COP.Hide

End Sub

Private Sub Userform_Activate()

TB_COP_1.Value = Range("COP_1").Text
TB_COP_2.Value = Range("COP_2").Text
TB_COP_3.Value = Range("COP_3").Text
TB_COP_4.Value = Range("COP_4").Text
TB_COP_5.Value = Range("COP_5").Text
TB_COP_6.Value = Range("COP_6").Text
TB_COP_7.Value = Range("COP_7").Text
TB_COP_8.Value = Range("COP_8").Text
TB_COP_9.Value = Range("COP_9").Text
TB_COP_10.Value = Range("COP_10").Text
TB_COP_11.Value = Range("COP_11").Text
TB_COP_12.Value = Range("COP_12").Text
TB_COP_13.Value = Range("COP_13").Text
TB_COP_14.Value = Range("COP_14").Text
TB_COP_15.Value = Range("COP_15").Text

End Sub

' ============================================================
' MODULE : CR.frm
' ============================================================

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

' ============================================================
' MODULE : DEV.frm
' ============================================================

Private Sub Userform_Activate()

Range("JOUR_X") = Range("JOUR_J")

If Range("CK_55") = 1 Then

TB_DV_SGS.Text = Range("DJX_55")

End If

Range("DV_CS1").Value = "0"
Range("DV_CS2").Value = "0"
Range("DV_CS3").Value = "0"
Range("DV_CS4").Value = "0"
Range("DV_CS5").Value = "0"
Range("DV_CS6").Value = "0"
Range("DV_CS7").Value = "0"
Range("DV_CS8").Value = "0"
Range("DV_CS9").Value = "0"
Range("DV_CS10").Value = "0"
Range("DV_CS11").Value = "0"
Range("DV_CS12").Value = "0"

End Sub
Private Sub CB_DEV_VAL_Click()

If TB_PI.Visible = True And TB_DV_SGS.Text = "" Then

MsgBox "Le numéro de(s) SGS est manquant, compléter au format 000,000,000..."
TB_DV_SGS.BackColor = &H8080FF

Else

Call DEV_VAL_JJ

End If

DEV.Hide

End Sub
Private Sub Saisie_DV_ANN_Click()

TB_PI.Text = ""
TB_DV_SGS.Text = ""

DEV.Hide

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

End Sub
Private Sub TB_DV_GEC_Change()

If TB_DV_GEC.Text = "" Then
TB_DV_GEC.BackColor = &H8080FF
Else
TB_DV_GEC.BackColor = &H80000005
End If

End Sub
Private Sub TB_DV_SGS_Change()

If TB_DV_SGS.Text = "" Then
TB_DV_SGS.BackColor = &H8080FF
Else
TB_DV_SGS.BackColor = &H80000005
End If

End Sub
Private Sub TB_DV1_Click()

LB_DVC1.Visible = True
LB_DVS.Visible = True
TB_DVS.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS1").Value = Range("DV_CS1").Value + 1

LB_DVC1.Caption = Range("DV_CS1").Value


TB_DVS.Text = Range("TOT_DV_CV").Value


End Sub
Private Sub TB_DV2_Click()

LB_DVC2.Visible = True
LB_DVS.Visible = True
TB_DVS.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS2").Value = Range("DV_CS2").Value + 1

LB_DVC2.Caption = Range("DV_CS2").Value


TB_DVS.Text = Range("TOT_DV_CV").Value
End Sub
Private Sub TB_DV3_Click()

LB_DVC3.Visible = True
LB_DVS.Visible = True
TB_DVS.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS3").Value = Range("DV_CS3").Value + 1

LB_DVC3.Caption = Range("DV_CS3").Value


TB_DVS.Text = Range("TOT_DV_CV").Value

End Sub
Private Sub TB_DV4_Click()

LB_DVC4.Visible = True
LB_DVS.Visible = True
TB_DVS.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True
Range("DV_CS4").Value = Range("DV_CS4").Value + 1

LB_DVC4.Caption = Range("DV_CS4").Value


TB_DVS.Text = Range("TOT_DV_CV").Value

End Sub
Private Sub TB_DV5_Click()

LB_DVC5.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS5").Value = Range("DV_CS5").Value + 1

LB_DVC5.Caption = Range("DV_CS5").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub TB_DV6_Click()

LB_DVC6.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS6").Value = Range("DV_CS6").Value + 1

LB_DVC6.Caption = Range("DV_CS6").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub TB_DV7_Click()

LB_DVC7.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS7").Value = Range("DV_CS7").Value + 1

LB_DVC7.Caption = Range("DV_CS7").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub TB_DV8_Click()

LB_DVC8.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS8").Value = Range("DV_CS8").Value + 1

LB_DVC8.Caption = Range("DV_CS8").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub TB_DV9_Click()

LB_DVC9.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS9").Value = Range("DV_CS9").Value + 1

LB_DVC9.Caption = Range("DV_CS9").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub TB_DV10_Click()

LB_DVC10.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS10").Value = Range("DV_CS10").Value + 1

LB_DVC10.Caption = Range("DV_CS10").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub TB_DV11_Click()

LB_DVC11.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS11").Value = Range("DV_CS11").Value + 1

LB_DVC11.Caption = Range("DV_CS11").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub TB_DV12_Click()

LB_DVC12.Visible = True
LB_PI.Visible = True
TB_PI.Visible = True
LB_DV_SGS.Visible = True
TB_DV_SGS.Visible = True

Range("DV_CS12").Value = Range("DV_CS12").Value + 1

LB_DVC12.Caption = Range("DV_CS12").Value


TB_PI.Text = Range("TOT_DV_PI").Value

End Sub
Private Sub DEV_VAL_JJ()

Dim ADD_DSGS, ADD_C1, ADD_C2, ADD_C3, ADD_C4, ADD_C5, ADD_C6, ADD_C7, ADD_C8, ADD_C9, ADD_C10, ADD_C11, ADD_C12, ADD_INF

ADD_DSGS = Range("ADD_JJ_55").Value
ADD_INF = Range("ADD_JJ_13").Value

ADD_C1 = Range("ADD_JJ_43").Value
ADD_C2 = Range("ADD_JJ_44").Value
ADD_C3 = Range("ADD_JJ_45").Value
ADD_C4 = Range("ADD_JJ_46").Value
ADD_C5 = Range("ADD_JJ_47").Value
ADD_C6 = Range("ADD_JJ_48").Value
ADD_C7 = Range("ADD_JJ_49").Value
ADD_C8 = Range("ADD_JJ_50").Value
ADD_C9 = Range("ADD_JJ_51").Value
ADD_C10 = Range("ADD_JJ_52").Value
ADD_C11 = Range("ADD_JJ_53").Value
ADD_C12 = Range("ADD_JJ_54").Value

Sheets("LOG_M").Range(ADD_DSGS) = TB_DV_SGS.Text

Sheets("LOG_M").Range(ADD_INF) = Range("DV_PINF").Value

Sheets("LOG_M").Range(ADD_C1) = Range("DV_PC1").Value
Sheets("LOG_M").Range(ADD_C2) = Range("DV_PC2").Value
Sheets("LOG_M").Range(ADD_C3) = Range("DV_PC3").Value
Sheets("LOG_M").Range(ADD_C4) = Range("DV_PC4").Value
Sheets("LOG_M").Range(ADD_C5) = Range("DV_PC5").Value
Sheets("LOG_M").Range(ADD_C6) = Range("DV_PC6").Value
Sheets("LOG_M").Range(ADD_C7) = Range("DV_PC7").Value
Sheets("LOG_M").Range(ADD_C8) = Range("DV_PC8").Value
Sheets("LOG_M").Range(ADD_C9) = Range("DV_PC9").Value
Sheets("LOG_M").Range(ADD_C10) = Range("DV_PC10").Value
Sheets("LOG_M").Range(ADD_C11) = Range("DV_PC11").Value
Sheets("LOG_M").Range(ADD_C12) = Range("DV_PC12").Value


End Sub

' ============================================================
' MODULE : DIV.frm
' ============================================================

Private Sub BDD_PLUS_Click()

TB_BDD.Value = TB_BDD.Value + 1

End Sub


Private Sub CB_DIV_AN2_Click()

TB_MDG.BackColor = &H80000005
CB_MDG.Value = False
TB_MDG.Visible = False
TB_MDG.Text = ""

TB_BDD.BackColor = &H80000005
CB_BDD.Value = False
TB_BDD.Visible = False
TB_BDD.Text = ""

DIV.Hide

End Sub

Private Sub CB_DIV_VAL2_Click()

If CB_MDG.Value = True And TB_MDG.Text = "" Then

MsgBox "Le nombre de Mesure de Glissance est manquant"
TB_MDG.BackColor = &H8080FF

ElseIf CB_BDD.Value = True And TB_BDD.Text = "" Then

MsgBox "Le nombre de Balise de Détresse est manquant"
TB_BDD.BackColor = &H8080FF

Else

Dim ADD_MDG, ADD_BDD

ADD_MDG = Range("ADD_JJ_25").Value
ADD_BDD = Range("ADD_JJ_24").Value


Sheets("LOG_M").Range(ADD_MDG) = TB_MDG.Text
Sheets("LOG_M").Range(ADD_BDD) = TB_BDD.Text

End If

DIV.Hide

End Sub

Private Sub CB_MDG_Click()

If CB_MDG.Value = True Then

TB_MDG.Visible = True
MDG_PLUS.Visible = True
TB_MDG.Text = "1"

Else

TB_MDG.Text = ""
TB_MDG.Visible = False
MDG_PLUS.Visible = False
TB_MDG.BackColor = &H80000005

End If
End Sub
Private Sub CB_BDD_Click()

If CB_BDD.Value = True Then

TB_BDD.Visible = True
BDD_PLUS.Visible = True
TB_BDD.Text = "1"

Else

TB_BDD.Text = ""
TB_BDD.Visible = False
BDD_PLUS.Visible = False
TB_BDD.BackColor = &H80000005

End If
End Sub

Private Sub MDG_PLUS_Click()

TB_MDG.Value = TB_MDG.Value + 1

End Sub


Private Sub TB_BDD_Change()
If TB_BDD.Text = "" Then
TB_BDD.BackColor = &H8080FF
Else
TB_BDD.BackColor = &H80000005
End If
End Sub

Private Sub TB_MDG_Change()
If TB_MDG.Text = "" Then
TB_MDG.BackColor = &H8080FF
Else
TB_MDG.BackColor = &H80000005
End If
End Sub

' ============================================================
' MODULE : ETATS.frm
' ============================================================

Private Sub DPU_PLUS_Click()
TB_DPU.Value = TB_DPU.Value + 1
End Sub

Private Sub ETA_PLUS_Click()
TB_ETA2.Value = TB_ETA2.Value + 1
End Sub

Private Sub ETC_PLUS_Click()
TB_ETC2.Value = TB_ETC2.Value + 1
End Sub

Private Sub ETV_PLUS_Click()
TB_ETV2.Value = TB_ETV2.Value + 1
End Sub

Private Sub Userform_Activate()

TB_ETV2.BackColor = &H80000005
CB_ETV2.Value = False
TB_ETV2.Visible = False
TB_ETV2.Text = ""

TB_ETA2.BackColor = &H80000005
CB_ETA2.Value = False
TB_ETA2.Visible = False
TB_ETA2.Text = ""

TB_ETC2.BackColor = &H80000005
CB_ETC2.Value = False
TB_ETC2.Visible = False
TB_ETC2.Text = ""

TB_DPU.BackColor = &H80000005
CB_DPU.Value = False
TB_DPU.Visible = False
TB_DPU.Text = ""

End Sub

Private Sub CB_ET_CNX_Click()

TB_ETV2.BackColor = &H80000005
CB_ETV2.Value = False
TB_ETV2.Visible = False
TB_ETV2.Text = ""

TB_ETA2.BackColor = &H80000005
CB_ETA2.Value = False
TB_ETA2.Visible = False
TB_ETA2.Text = ""

TB_ETC2.BackColor = &H80000005
CB_ETC2.Value = False
TB_ETC2.Visible = False
TB_ETC2.Text = ""

TB_DPU.BackColor = &H80000005
CB_DPU.Value = False
TB_DPU.Visible = False
TB_DPU.Text = ""

ETATS.Hide

End Sub
Private Sub CB_ET_VAL_Click()

If CB_ETV2.Value = True And TB_ETV2.Text = "" Then

MsgBox "Le nombre d'Etat de Veille est manquant"
TB_POR.BackColor = &H8080FF

ElseIf CB_ETA2.Value = True And TB_ETA2.Text = "" Then

MsgBox "Le nombre d'Etat d'Alerte est manquant"
TB_VTF.BackColor = &H8080FF

ElseIf CB_ETC2.Value = True And TB_ETC2.Text = "" Then

MsgBox "Le nombre d'Etat d'Accident est manquant"
TB_NEI.BackColor = &H8080FF

ElseIf CB_DPU.Value = True And TB_DPU.Text = "" Then

MsgBox "Le nombre de Déclenchement Plan d'Urgence est manquant"
TB_AUT3.BackColor = &H8080FF

Else

Dim ADD_ETV, ADD_ETA, ADD_ETC, ADD_DPU

ADD_ETV = Range("ADD_JJ_16").Value
ADD_ETA = Range("ADD_JJ_17").Value
ADD_ETC = Range("ADD_JJ_18").Value
ADD_DPU = Range("ADD_JJ_19").Value

Sheets("LOG_M").Range(ADD_ETV) = TB_ETV2.Text
Sheets("LOG_M").Range(ADD_ETA) = TB_ETA2.Text
Sheets("LOG_M").Range(ADD_ETC) = TB_ETC2.Text
Sheets("LOG_M").Range(ADD_DPU) = TB_DPU.Text



End If

ETATS.Hide

End Sub

Private Sub CB_ETV2_Click()

If CB_ETV2.Value = True Then

TB_ETV2.Visible = True
ETV_PLUS.Visible = True
TB_ETV2.Text = "1"

Else

TB_ETV2.Text = ""
TB_ETV2.Visible = False
ETV_PLUS.Visible = False
TB_ETV2.BackColor = &H80000005

End If

End Sub
Private Sub CB_ETA2_Click()

If CB_ETA2.Value = True Then

TB_ETA2.Visible = True
ETA_PLUS.Visible = True
TB_ETA2.Text = "1"

Else

TB_ETA2.Text = ""
TB_ETA2.Visible = False
ETA_PLUS.Visible = False
TB_ETA2.BackColor = &H80000005

End If

End Sub
Private Sub CB_ETC2_Click()

If CB_ETC2.Value = True Then

TB_ETC2.Visible = True
ETC_PLUS.Visible = True
TB_ETC2.Text = "1"

Else

TB_ETC2.Text = ""
TB_ETC2.Visible = False
ETC_PLUS.Visible = False
TB_ETC2.BackColor = &H80000005

End If

End Sub
Private Sub CB_DPU_Click()

If CB_DPU.Value = True Then

TB_DPU.Visible = True
DPU_PLUS.Visible = True
TB_DPU.Text = "1"

Else

TB_DPU.Text = ""
TB_DPU.Visible = False
DPU_PLUS.Visible = False
TB_DPU.BackColor = &H80000005

End If

End Sub
Private Sub TB_ETV2_Change()
If TB_ETV2.Text = "" Then
TB_ETV2.BackColor = &H8080FF
Else
TB_ETV2.BackColor = &H80000005
End If
End Sub
Private Sub TB_ETA2_Change()
If TB_ETA2.Text = "" Then
TB_ETA2.BackColor = &H8080FF
Else
TB_ETA2.BackColor = &H80000005
End If
End Sub
Private Sub TB_ETC2_Change()
If TB_ETC2.Text = "" Then
TB_ETC2.BackColor = &H8080FF
Else
TB_ETC2.BackColor = &H80000005
End If
End Sub
Private Sub TB_DPU_Change()
If TB_DPU.Text = "" Then
TB_DPU.BackColor = &H8080FF
Else
TB_DPU.BackColor = &H80000005
End If
End Sub

' ============================================================
' MODULE : FAT.frm
' ============================================================

Private Sub CB_ANN_FNA_Click()
FAT.Hide
End Sub
Private Sub CB_VAL_FNA_Click()

Dim ADD_FAT, ADD_NOT, ADD_MAIP, ADD_SAIP, ADD_MRR, ADD_GRU, ADD_138, ADD_139, ADD_140, ADD_141, ADD_142, ADD_143

ADD_FAT = Range("ADD_MX_138").Value
ADD_NOT = Range("ADD_MX_139").Value
ADD_MAIP = Range("ADD_MX_140").Value
ADD_SAIP = Range("ADD_MX_141").Value
ADD_MRR = Range("ADD_MX_142").Value
ADD_GRU = Range("ADD_MX_143").Value

ADD_138 = Range("ADD_MCR_138").Value
ADD_139 = Range("ADD_MCR_139").Value
ADD_140 = Range("ADD_MCR_140").Value
ADD_141 = Range("ADD_MCR_141").Value
ADD_142 = Range("ADD_MCR_142").Value
ADD_143 = Range("ADD_MCR_143").Value

Sheets("LOG_ANNUEL").Range(ADD_FAT) = TB_FAT.Text
Sheets("LOG_ANNUEL").Range(ADD_NOT) = TB_NOT.Text
Sheets("LOG_ANNUEL").Range(ADD_MAIP) = TB_MAIP.Text
Sheets("LOG_ANNUEL").Range(ADD_SAIP) = TB_SAIP.Text
Sheets("LOG_ANNUEL").Range(ADD_MRR) = TB_MRR.Text
Sheets("LOG_ANNUEL").Range(ADD_GRU) = TB_GRU.Text

Sheets("LOG_M (2)").Range(ADD_138) = TB_FAT.Text
Sheets("LOG_M (2)").Range(ADD_139) = TB_NOT.Text
Sheets("LOG_M (2)").Range(ADD_140) = TB_MAIP.Text
Sheets("LOG_M (2)").Range(ADD_141) = TB_SAIP.Text
Sheets("LOG_M (2)").Range(ADD_142) = TB_MRR.Text
Sheets("LOG_M (2)").Range(ADD_143) = TB_GRU.Text


FAT.Hide

Range("MLG_FAT") = TB_FAT.Text
Range("MLG_NOT") = TB_NOT.Text
Range("MLG_FATA") = TB_FATA.Text
Range("MLG_AIP1") = TB_SAIP.Text
Range("MLG_AIP2") = TB_MAIP.Text
Range("MLG_MRR") = TB_MRR.Text
Range("MLG_GRU") = TB_GRU.Text


Range("CR_FATT") = TB_FATT.Text
Range("CR_NOTA") = TB_NOT.Text
Range("CR_MRRT") = TB_MRRT.Text
Range("CR_GRUA") = TB_GRU.Text
Range("CR_GRUT") = TB_GRUT.Text

If TB_FATA.Text = "" Then

Range("CR_FATA") = 0
Else

Range("CR_FATA") = TB_FATA.Text

End If

If TB_MRR.Text = "" Then

Range("CR_MRRA") = 0

Else

Range("CR_MRRA") = TB_MRR.Text

End If

If TB_GRU.Text = "" Then

Range("CR_GRUA") = 0

Else

Range("CR_GRUA") = TB_GRU.Text

End If


Range("ETS_TVX") = 1

ActiveWorkbook.Save

End Sub
Private Sub MOIS_FNA_Change()

Range("MOIS_X") = MOIS_FNA.ListIndex + 1

TB_FAT.Text = Range("DMX_138")
TB_NOT.Text = Range("DMX_139")
TB_MAIP.Text = Range("DMX_140")
TB_SAIP.Text = Range("DMX_141")
TB_MRR.Text = Range("DMX_142")
TB_GRU.Text = Range("DMX_143")

End Sub

' ============================================================
' MODULE : FRET.frm
' ============================================================

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

' ============================================================
' MODULE : HDG.frm
' ============================================================

Private Sub CB_CNX_Click()
HDG.Hide
End Sub
Private Sub CB_HDG_Change()

Range("HDG_LI1") = CB_HDG.ListIndex + 1

TB_HDG1.Value = Range("HDG_R1").Text
TB_HDG2.Value = Range("HDG_R2").Text

End Sub
Private Sub CB_VAL_HDG_Click()

Dim ADD_TRI, ADD_NAM2

ADD_TRI = Range("HDG_TRI").Value
ADD_NAM2 = Range("HDG_NAM").Value

Sheets("DATA_BASE").Range(ADD_TRI) = TB_HDG1.Text
Sheets("DATA_BASE").Range(ADD_NAM2) = TB_HDG2.Text

HDG.Hide

End Sub

' ============================================================
' MODULE : INTER.frm
' ============================================================

Private Sub AUT2_PLUS_Click()

TB_AUT2.Value = TB_AUT2.Value + 1

End Sub

Private Sub BAL_PLUS_Click()

TB_BAL.Value = TB_BAL.Value + 1

End Sub

Private Sub INF_PLUS_Click()

TB_INF.Value = TB_INF.Value + 1

End Sub
Private Sub Userform_Activate()

TB_INF.BackColor = &H80000005
CB_INF.Value = False
TB_INF.Visible = False
TB_INF.Text = ""

TB_BAL.BackColor = &H80000005
CB_BAL.Value = False
TB_BAL.Visible = False
TB_BAL.Text = ""

TB_AUT2.BackColor = &H80000005
CB_AUT2.Value = False
TB_AUT2.Visible = False
TB_AUT2.Text = ""

TB_APRE.BackColor = &H80000005
TB_APRE.Visible = False
TB_APRE.Text = ""

End Sub

Private Sub CB_INF_Click()

If CB_INF.Value = True Then

MsgBox ("Rappel!: Si un ou plusieurs déversement(s) ont eu lieu, les interventions INFRA correspondantes sont déjà comptabilisées, inutile de les re-saisir!")

TB_INF.Visible = True
INF_PLUS.Visible = True
TB_INF.Text = "1"

Else

TB_INF.Text = ""
TB_INF.Visible = False
INF_PLUS.Visible = False
TB_INF.BackColor = &H80000005

End If

End Sub
Private Sub CB_BAL_Click()

If CB_BAL.Value = True Then

TB_BAL.Visible = True
BAL_PLUS.Visible = True
TB_BAL.Text = "1"

Else

TB_BAL.Text = ""
TB_BAL.Visible = False
BAL_PLUS.Visible = False
TB_BAL.BackColor = &H80000005

End If
End Sub
Private Sub CB_AUT2_Click()

If CB_AUT2.Value = True Then

TB_AUT2.Visible = True
AUT2_PLUS.Visible = True
TB_AUT2.Text = "1"
TB_APRE.Visible = True
LB_APRE.Visible = True
TB_APRE.Text = ""

Else

TB_AUT2.Text = ""
TB_AUT2.Visible = False
AUT2_PLUS.Visible = False
TB_AUT2.BackColor = &H80000005
TB_APRE.Visible = False
LB_APRE.Visible = False

End If
End Sub
Private Sub CB_INT_CNX_Click()

TB_INF.BackColor = &H80000005
CB_INF.Value = False
TB_INF.Visible = False
TB_INF.Text = ""

TB_BAL.BackColor = &H80000005
CB_BAL.Value = False
TB_BAL.Visible = False
TB_BAL.Text = ""

TB_AUT2.BackColor = &H80000005
CB_AUT2.Value = False
TB_AUT2.Visible = False
TB_AUT2.Text = ""

TB_APRE.BackColor = &H80000005
TB_APRE.Visible = False
TB_APRE.Text = ""

INTER.Hide

End Sub
Private Sub CB_INT_VAL_Click()

If CB_INF.Value = True And TB_INF.Text = "" Then

MsgBox "Le nombre d'intervention INFRA est manquant"
TB_INF.BackColor = &H8080FF

ElseIf CB_BAL.Value = True And TB_BAL.Text = "" Then

MsgBox "Le nombre d'intervention BAL est manquant"
TB_BAL.BackColor = &H8080FF

ElseIf CB_AUT2.Value = True And TB_AUT2.Text = "" Then

MsgBox "Le nombre d'intervention AUTRE est manquant"
TB_AUT2.BackColor = &H8080FF

ElseIf CB_AUT2.Value = True And TB_APRE.Text = "" Then

MsgBox "Préciser le service concerné par AUTRE pour poursuivre."
TB_APRE.BackColor = &H8080FF

Else

Dim ADD_INF, ADD_BAL, ADD_AUT2, ADD_APRE

ADD_INF = Range("ADD_JJ_13").Value
ADD_BAL = Range("ADD_JJ_12").Value
ADD_AUT2 = Range("ADD_JJ_14").Value
ADD_APRE = Range("ADD_JJ_15").Value

Sheets("LOG_M").Range(ADD_INF) = TB_INF.Text
Sheets("LOG_M").Range(ADD_BAL) = TB_BAL.Text
Sheets("LOG_M").Range(ADD_AUT2) = TB_AUT2.Text
Sheets("LOG_M").Range(ADD_APRE) = TB_APRE.Text


End If

INTER.Hide

End Sub
Private Sub TB_APRE_Change()
If TB_APRE.Text = "" Then
TB_APRE.BackColor = &H8080FF
Else
TB_APRE.BackColor = &H80000005
End If
End Sub
Private Sub TB_AUT2_Change()

If TB_AUT2.Value <> "" Then
TB_APRE.Visible = True
LB_APRE.Visible = True
Else
TB_APRE.Visible = False
LB_APRE.Visible = False
End If

If TB_AUT2.Text = "" Then
TB_AUT2.BackColor = &H8080FF
Else
TB_AUT2.BackColor = &H80000005
End If

End Sub
Private Sub TB_BAL_Change()
If TB_BAL.Text = "" Then
TB_BAL.BackColor = &H8080FF
Else
TB_BAL.BackColor = &H80000005
End If
End Sub

Private Sub TB_INF_Change()
If TB_INF.Text = "" Then
TB_INF.BackColor = &H8080FF
Else
TB_INF.BackColor = &H80000005
End If
End Sub

' ============================================================
' MODULE : I_NP.frm
' ============================================================

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

' ============================================================
' MODULE : LEG.frm
' ============================================================

Private Sub CB_LEG_CNX_Click()
LEG.Hide
End Sub

' ============================================================
' MODULE : MAA_1.frm
' ============================================================

Private Sub AUT3_PLUS_Click()
TB_AUT3.Value = TB_AUT3.Value + 1
End Sub

Private Sub CDM_PLUS_Click()
TB_CDMR.Value = TB_CDMR.Value + 1
End Sub

Private Sub NEI_PLUS_Click()
TB_NEI.Value = TB_NEI.Value + 1
End Sub

Private Sub POR_PLUS_Click()
TB_POR.Value = TB_POR.Value + 1
End Sub

Private Sub Userform_Activate()

CB_POR.Value = False
TB_POR.Visible = False
TB_POR.Text = ""

CB_VTF.Value = False
TB_VTF.Visible = False
TB_VTF.Text = ""

CB_NEI.Value = False
TB_NEI.Visible = False
TB_NEI.Text = ""

CB_AUT3.Value = False
TB_AUT3.Visible = False
TB_AUT3.Text = ""

CB_CDM.Value = False
TB_CDMR.Visible = False
TB_CDMR.Text = ""

LB_CDM.Visible = False
TB_ICDM.Visible = False
TB_ICDM.Text = ""
TB_TICDM.Visible = False
TB_TICDM.Text = ""

End Sub

Private Sub CB_MAA_CNX_Click()

TB_POR.BackColor = &H80000005
CB_POR.Value = False
TB_POR.Visible = False
TB_POR.Text = ""

TB_VTF.BackColor = &H80000005
CB_VTF.Value = False
TB_VTF.Visible = False
TB_VTF.Text = ""

TB_NEI.BackColor = &H80000005
CB_NEI.Value = False
TB_NEI.Visible = False
TB_NEI.Text = ""

TB_AUT3.BackColor = &H80000005
CB_AUT3.Value = False
TB_AUT3.Visible = False
TB_AUT3.Text = ""

TB_CDMR.BackColor = &H80000005
CB_CDM.Value = False
TB_CDMR.Visible = False
TB_CDMR.Text = ""

LB_CDM.Visible = False
TB_ICDM.BackColor = &H80000005
TB_ICDM.Visible = False
TB_ICDM.Text = ""

TB_TICDM.BackColor = &H80000005
TB_TICDM.Visible = False
TB_TICDM.Text = ""

MAA_1.Hide

End Sub

Private Sub CB_MAA_VAL_Click()


If CB_POR.Value = True And TB_POR.Text = "" Then

MsgBox "Le nombre de MAA Pluie/Orage est manquant"
TB_POR.BackColor = &H8080FF

ElseIf CB_VTF.Value = True And TB_VTF.Text = "" Then

MsgBox "Le nombre de MAA Vent Fort est manquant"
TB_VTF.BackColor = &H8080FF

ElseIf CB_NEI.Value = True And TB_NEI.Text = "" Then

MsgBox "Le nombre de MAA Neige est manquant"
TB_NEI.BackColor = &H8080FF

ElseIf CB_AUT3.Value = True And TB_AUT3.Text = "" Then

MsgBox "Le nombre de MAA Autre est manquant"
TB_AUT3.BackColor = &H8080FF

ElseIf CB_CDM.Value = True And TB_CDMR.Text = "" Then

MsgBox "Le nombre de MAA Coup de Mer est manquant"
TB_CDM.BackColor = &H8080FF

ElseIf CB_CDM.Value = True And TB_ICDM.Text = "" Then

MsgBox "Vous devez saisir le nombre d'inspection(s) Coup de Mer pour poursuivre"
TB_ICDM.BackColor = &H8080FF

ElseIf CB_CDM.Value = True And TB_TICDM.Text = "" Then

MsgBox "Vous devez saisir le nombre d'inspection(s) Coup de Mer pour poursuivre"
TB_TICDM.BackColor = &H8080FF


Else

Dim ADD_PO, ADD_VF, ADD_NE, ADD_CM, ADD_IM, ADD_AU, ADD_TCDM

ADD_PO = Range("ADD_JJ_145").Value
ADD_VF = Range("ADD_JJ_146").Value
ADD_NE = Range("ADD_JJ_147").Value
ADD_CM = Range("ADD_JJ_149").Value
ADD_IM = Range("ADD_JJ_150").Value
ADD_TCDM = Range("ADD_JJ_151").Value
ADD_AU = Range("ADD_JJ_148").Value

Sheets("LOG_M").Range(ADD_PO) = TB_POR.Text
Sheets("LOG_M").Range(ADD_VF) = TB_VTF.Text
Sheets("LOG_M").Range(ADD_NE) = TB_NEI.Text
Sheets("LOG_M").Range(ADD_CM) = TB_CDMR.Text
Sheets("LOG_M").Range(ADD_IM) = TB_ICDM.Text
Sheets("LOG_M").Range(ADD_TCDM) = TB_TICDM.Text
Sheets("LOG_M").Range(ADD_AU) = TB_AUT3.Text

End If

MAA_1.Hide

End Sub
Private Sub CB_POR_Click()

If CB_POR.Value = True Then

TB_POR.Visible = True
POR_PLUS.Visible = True
TB_POR.Text = "1"

Else

TB_POR.Text = ""
TB_POR.Visible = False
POR_PLUS.Visible = False
TB_POR.BackColor = &H80000005

End If
End Sub
Private Sub CB_VTF_Click()

If CB_VTF.Value = True Then

TB_VTF.Visible = True
VTF_PLUS.Visible = True
TB_VTF.Text = "1"

Else

TB_VTF.Text = ""
TB_VTF.Visible = False
VTF_PLUS.Visible = False
TB_VTF.BackColor = &H80000005

End If
End Sub
Private Sub CB_NEI_Click()

If CB_NEI.Value = True Then

TB_NEI.Visible = True
NEI_PLUS.Visible = True
TB_NEI.Text = "1"

Else

TB_NEI.Text = ""
TB_NEI.Visible = False
NEI_PLUS.Visible = False
TB_NEI.BackColor = &H80000005

End If
End Sub
Private Sub CB_AUT3_Click()

If CB_AUT3.Value = True Then

TB_AUT3.Visible = True
AUT3_PLUS.Visible = True
TB_AUT3.Text = "1"

Else

TB_AUT3.Text = ""
TB_AUT3.Visible = False
AUT3_PLUS.Visible = False
TB_AUT3.BackColor = &H80000005

End If
End Sub
Private Sub CB_CDM_Click()

If CB_CDM.Value = True Then

TB_CDMR.Visible = True
TB_ICDM.Visible = True
TB_TICDM.Visible = True
LB_CDM.Visible = True
CDM_PLUS.Visible = True
TB_CDMR.Text = "1"

Else

TB_CDMR.Text = ""
TB_CDMR.Visible = False
TB_CDMR.BackColor = &H80000005
TB_ICDM.Visible = False
TB_TICDM.Visible = False
LB_CDM.Visible = False
CDM_PLUS.Visible = False

End If
End Sub
Private Sub TB_CDM_Change()
If TB_CDMR.Value <> "" Then
TB_ICDM.Visible = True
TB_TICDM.Visible = True
LB_CDM.Visible = True
Else
TB_ICDM.Visible = False
TB_TICDM.Visible = False
LB_CDM.Visible = False
End If
End Sub

Private Sub TB_POR_Change()
If TB_POR.Text = "" Then
TB_POR.BackColor = &H8080FF
Else
TB_POR.BackColor = &H80000005
End If
End Sub
Private Sub TB_VTF_Change()
If TB_VTF.Text = "" Then
TB_VTF.BackColor = &H8080FF
Else
TB_VTF.BackColor = &H80000005
End If
End Sub
Private Sub TB_NEI_Change()
If TB_NEI.Text = "" Then
TB_NEI.BackColor = &H8080FF
Else
TB_NEI.BackColor = &H80000005
End If
End Sub
Private Sub TB_AUT3_Change()
If TB_AUT3.Text = "" Then
TB_AUT3.BackColor = &H8080FF
Else
TB_AUT3.BackColor = &H80000005
End If
End Sub
Private Sub TB_CDMR_Change()
If TB_CDMR.Text = "" Then
TB_CDMR.BackColor = &H8080FF
Else
TB_CDMR.BackColor = &H80000005
End If
End Sub
Private Sub TB_ICDM_Change()
If TB_ICDM.Text = "" Then
TB_ICDM.BackColor = &H8080FF
Else
TB_ICDM.BackColor = &H80000005
End If
End Sub

Private Sub VTF_PLUS_Click()
TB_VTF.Value = TB_VTF.Value + 1
End Sub

' ============================================================
' MODULE : MAJ_DB1.frm
' ============================================================

Private Sub CB_MAJ_CNX_Click()
MAJ_DB1.Hide
End Sub

Private Sub CB_PRIX_Click()
PRIX_PRES.Show
End Sub

Private Sub CommandButton1_Click()
COP.Show
End Sub

' ============================================================
' MODULE : PERMIS.frm
' ============================================================

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

' ============================================================
' MODULE : PRIX_PRES.frm
' ============================================================

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

' ============================================================
' MODULE : SAISIE_J.frm
' ============================================================

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

' ============================================================
' MODULE : SURTOU_1.frm
' ============================================================

Private Sub CB_AUJ_Click()

If CB_AUJ.Value = True Then

CB_ALSDAT.Value = False

Range("ST_ALT_DATE").Value = 0

Else

CB_ALSDAT.Value = True

End If

End Sub

Private Sub CB_COP_ADD_Click()
COP.Show
End Sub

Private Sub CB_COP_Change()
Range("ST_COP").Value = CB_COP.ListIndex + 1
End Sub

Private Sub CommandButton1_Click()
Dim ALT_DATE, ALT_ENT
ALT_ENT = TB_SDAT.Value
ALT_DATE = Format(ALT_ENT, "dd/mm/yyyy")

Sheets("LOG_ST").Range("H6").End(xlDown).Offset(1, 0) = ALT_DATE

End Sub

Private Sub TB_SDAT_afterupdate()
Range("ST_ENT_DATE").Value = TB_SDAT.Value
End Sub

Private Sub Userform_Activate()

CB_AUJ.Value = False
CB_ALSDAT.Value = False

CB_SCIES.Value = ""
CB_ST_APP.Value = ""
CB_COP.Value = ""

TB_SDAT.Text = ""
TB_TRIC.Text = ""
TB_IMM.Text = ""
TB_PKG.Text = ""
TB_TRIA.Text = ""
TB_ACOP.Text = ""
TB_BON.Text = ""
TB_BOFF.Text = ""

Range("ST_LI_CIE").Value = ""
Range("ST_LI_APP").Value = ""
Range("ST_EY_TCIE").Value = ""
Range("ST_EY_TAPP").Value = ""
Range("ST_EY_IMM").Value = ""
Range("ST_ASS_CR").Value = ""
Range("ST_PKG").Value = ""
Range("ST_COP").Value = ""

End Sub
Private Sub CB_ALSDAT_Click()

If CB_ALSDAT.Value = True Then

Range("ST_ALT_DATE").Value = 1

CB_AUJ.Value = False
LB_SDAT.Visible = True
TB_SDAT.Visible = True

Else

Range("ST_ALT_DATE").Value = 0

CB_AUJ.Value = True
LB_SDAT.Visible = False
TB_SDAT.Visible = False


End If
End Sub
Private Sub CB_ANS1_Click()
SURTOU_1.Hide
End Sub
Private Sub CB_ASSC_Click()
If CB_ASSC.Value = True Then
CB_CASS.Visible = True
Else
CB_CASS.Visible = False
Range("ST_ASS_CR").Value = ""
End If
End Sub
Private Sub CB_CASS_Change()

Range("ST_ASS_CR").Value = CB_CASS.ListIndex + 1

Dim ADD_ASS_CR

ADD_ASS_CR = Range("ADD_ASS_CR").Value

Sheets("DATA_ST").Range(ADD_ASS_CR) = Range("ST_HDG_3").Value

End Sub
Private Sub CB_SCIES_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub
Private Sub CB_SCIES_AfterUpdate()
Sheets("DATA_MAIN").Range("ST_LI_CIE").Value = CB_SCIES.ListIndex
Sheets("DATA_ST").Range("ST_LI_CIE3").Value = CB_SCIES.ListIndex
TB_TRIC.Text = Range("ST_MC_TCIE").Value
LB_SASS.Caption = Range("ST_HDG_2").Value

End Sub
Private Sub CB_ST_APP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = 0
End Sub
Private Sub CB_ST_APP_AfterUpdate()
Sheets("DATA_MAIN").Range("ST_LI_APP").Value = CB_ST_APP.ListIndex
TB_TRIA.Text = Range("ST_MC_TAPP").Value
End Sub

Private Sub CB_SURS_Click()

If Range("ST_LI_CIE").Value = "" Then

MsgBox "Le champs Compagnie est manquant, saisir pour poursuivre"

ElseIf Range("ST_COP").Value = "" Then

MsgBox "Le champs COP est manquant, saisir pour poursuivre"

ElseIf CB_AUJ.Value = False And CB_ALSDAT.Value = False Then

MsgBox "Saisir une date pour poursuivre"

ElseIf TB_IMM.Text = "" Then

MsgBox "Le champs Immatriculation est manquant, saisir pour poursuivre"

ElseIf TB_PKG.Text = "" Then

MsgBox "Le champs Parking est manquant, saisir pour poursuivre"

ElseIf Range("ST_LI_APP").Value = "" Then

MsgBox "Le champs Appareil est manquant, saisir pour poursuivre"

ElseIf TB_ACOP.Text = "" Then

MsgBox "Le champs Arrivée COP est manquant, saisir pour poursuivre"

ElseIf TB_BON.Text = "" Then

MsgBox "Le champs ON BT est manquant, saisir pour poursuivre"

ElseIf TB_BOFF.Text = "" Then

MsgBox "Le champs OFF BT est manquant, saisir pour poursuivre"

Else

Sheets("LOG_ST").Range("G6").End(xlDown).Offset(1, 0) = Range("ST_HDG_4").Value
Sheets("LOG_ST").Range("K6").End(xlDown).Offset(1, 0) = Range("ST_RLI_CIE").Value
Sheets("LOG_ST").Range("L6").End(xlDown).Offset(1, 0) = Range("ST_RLI_APP").Value
Sheets("LOG_ST").Range("M6").End(xlDown).Offset(1, 0) = TB_IMM.Text
Sheets("LOG_ST").Range("N6").End(xlDown).Offset(1, 0) = Range("ST_PKG").Value
Sheets("LOG_ST").Range("O6").End(xlDown).Offset(1, 0) = TB_BON.Text
Sheets("LOG_ST").Range("P6").End(xlDown).Offset(1, 0) = TB_BOFF.Text
Sheets("LOG_ST").Range("I6").End(xlDown).Offset(1, 0) = Range("ST_COP").Value
Sheets("LOG_ST").Range("J6").End(xlDown).Offset(1, 0) = TB_ACOP.Text

If Range("ST_ALT_DATE").Value = 1 Then

Sheets("LOG_ST").Range("H6").End(xlDown).Offset(1, 0) = Range("ST_DATE_ALT").Value

ElseIf Range("ST_ALT_DATE").Value = 0 Then

Sheets("LOG_ST").Range("H6").End(xlDown).Offset(1, 0) = Range("DATE_G").Value

End If

SURTOU_2.Show

End If
End Sub
Private Sub TB_PKG_Change()
Range("ST_PKG").Value = TB_PKG.Text
End Sub
Private Sub TB_PKG_AfterUpdate()
If Range("ST_PKG_CT") = 1 Then
MsgBox "Ce parking n'existe pas ou n'est pas encore répertorié, recommencer."
End If
End Sub

Private Sub TB_TRIA_AfterUpdate()

Range("ST_EY_TAPP").Value = TB_TRIA.Text

If Range("ST_EY_APP_CT").Value = "1" Then
MsgBox "Ce tigramme ne correspond à aucun appareil, recommencer."
Else
CB_ST_APP.Text = Range("ST_EY_MC_APP").Value
End If

End Sub
Private Sub TB_TRIC_AfterUpdate()

Range("ST_EY_TCIE").Value = TB_TRIC.Text

If Range("ST_EY_CIE_CT").Value = "1" Then
MsgBox "Ce tigramme ne correspond à aucune compagnie, recommencer."
Else
LB_SASS.Caption = Range("ST_HDG_2").Value
End If

End Sub

' ============================================================
' MODULE : SURTOU_2.frm
' ============================================================

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

' ============================================================
' MODULE : S_ANX.bas
' ============================================================

Sub S_ANX_Cliquer()

ANX.Show

End Sub
Sub S_LEG_Cliquer()

LEG.Show

End Sub
Sub S_ADM_Cliquer()

ADMIN.Show

End Sub
Sub S_CONSUL_C_Cliquer()

CONSUL_MAIN.Show

End Sub
Sub Saisie_START_Cliquer()

If Range("DATA_HEURE").Value = 0 Or Range("DATA_HEURE").Value < 3 Then

MsgBox ("En cas de modification, utilisez CONSULTATION DONNEES")

CONSUL_MAIN.Show

Else

SAISIE_J.Show

End If

End Sub
Sub B_SURTOU_Cliquer()

SURTOU_1.Show

End Sub
Sub Auto_Open()

Call DATA_MAIN_RESET

Range("JOUR_DATA").FormulaR1C1 = "=DAY(TODAY())"

Range("JOUR_X").Value = Range("JOUR_DATA").Value

If Range("DJX_1") <> 3 Then

Range("ACT_MOD") = 0

End If

If Range("JOUR_J") = 1 Then

Sheets("LOG_ST").Range("D141").End(xlDown).Offset(1, 0) = Range("ST_MOY").Value

Sheets("LOG_M").Visible = True
Sheets("LOG_M").Select
Sheets("LOG_M").Copy Before:=Sheets(1)
Sheets("LOG_M (2)").Select
Sheets("SAISIE").Select
Sheets("LOG_M (2)").Visible = xlSheetVeryHidden
Sheets("LOG_M").Visible = xlSheetVeryHidden

End If

If Range("JOUR_J") = 1 And Range("MOIS_M") = 1 Then

Sheets("LOG_ST").Range("D141").End(xlDown).Offset(1, 0) = Range("ST_MOY").Value

End If

If Range("JOUR_J") = 2 And Range("MOIS_M") = 1 And Range("ET_ARSTANN") = 0 Then

Call LOG_ANN_ST

Range("ET_ARSTANN") = 1

End If

If Range("JOUR_J") = 9 And Range("MOIS_M") = 1 And Range("ET_ARCHANN") = 0 Then

Call LOG_ANN

Range("ET_ARCHANN") = 1

End If


If Range("JOUR_J") = 10 Then

Range("ET_ARSTANN") = 0
Range("ET_ARCHANN") = 0

Range("ETS_CAR") = 0
Range("ETS_TVX") = 0
Range("ETS_PER") = 0
Range("ET_ARCH_ST") = 0
Range("ET_ARCH") = 0
Range("ET_MLG").Value = 0

Range("FAITM1") = ""
Range("FAITM2") = ""
Range("FAITM3") = ""
Range("FAITM4") = ""
Range("FAITM5") = ""
Range("FAITM6") = ""
Range("FAITM7") = ""
Range("FAITM8") = ""
Range("CHAMPS_LIB") = ""

Range("CR_FATA") = ""
Range("CR_NOTA") = ""
Range("CR_MRRA") = ""
Range("CR_GRUA") = ""

Range("CR_PERT") = ""
Range("CR_PERM") = ""
Range("CR_PERC") = ""
Range("CR_PERP") = ""
Range("CR_PERH") = ""

Sheets("SAISIE").Unprotect

Sheets("SAISIE").Protect


End If

If Range("ET_MLG").Value = 0 And Range("ETS_CAR").Value = 1 And Range("ETS_TVX").Value = 1 And Range("ETS_PER").Value = 1 Then

Call LOG_MLG

Sheets("SAISIE").Unprotect

Sheets("SAISIE").Protect

Range("ETS_CAR").Value = 0
Range("ETS_TVX").Value = 0
Range("ETS_PER").Value = 0

Call LOG_MENS

Call ARCH

Range("ET_ARCH").Value = 1

MsgBox ("ARCHIVAGE MENSUEL EFFECTUE")

End If


End Sub
Sub LOG_ANN_ST()

Sheets("LOG_ST").Visible = True
Sheets("LOG_ST").Copy

    Range("B2:AQ161").Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Dim STAN_NAME

STAN_NAME = Range("STAN_NAME").Value

ActiveWorkbook.SaveAs Filename:="X:\Anca\#-Groupe\A_DETP_OP\0. EQUIPE\Indicateurs\2025\Logs Données Mensuelles\" & STAN_NAME & ".xlsx"

ActiveWorkbook.Close

'CLEAR TOT ANN ST

Sheets("LOG_ST").Visible = True
Sheets("LOG_ST").Select

    Range("G6:AQ161").Select
    Selection.ClearContents
    
    Range("D141:D152").Select
    Selection.ClearContents
    
Sheets("LOG_ST").Visible = False

Sheets("SAISIE").Select

MsgBox ("LOG ANNUEL ST EDITE!")

End Sub
Sub LOG_ANN()

'TRANSFERT VERS A-1

Sheets("LOG_ANNUEL").Visible = True
Sheets("LOG_ANNUEL_A-1").Visible = True

    Sheets("LOG_ANNUEL").Select
    Range("E8:P75").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E8").Select
    ActiveSheet.Paste
    
    Sheets("LOG_ANNUEL").Select
    Range("E81:P162").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E81").Select
    ActiveSheet.Paste

    Sheets("LOG_ANNUEL").Select
    Range("E165:P222").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E165").Select
    ActiveSheet.Paste
    
    Sheets("LOG_ANNUEL").Select
    Range("E225:P231").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E225").Select
    ActiveSheet.Paste
    
    Sheets("LOG_ANNUEL").Select
    Range("E234:P240").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E234").Select
    ActiveSheet.Paste
    
    Sheets("LOG_ANNUEL").Select
    Range("E244:P254").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E244").Select
    ActiveSheet.Paste
    
    Sheets("LOG_ANNUEL").Select
    Range("E257:P280").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E257").Select
    ActiveSheet.Paste
    
    Sheets("LOG_ANNUEL").Select
    Range("E283:P283").Select
    Selection.Copy
    Sheets("LOG_ANNUEL_A-1").Select
    Range("E283").Select
    ActiveSheet.Paste

Sheets("LOG_ANNUEL").Visible = False
Sheets("LOG_ANNUEL_A-1").Visible = False

'EDIT LOG ANN

Sheets("LOG_ANNUEL").Visible = True
Sheets("LOG_ANNUEL").Copy

    Range("B2:R283").Select
    Selection.Copy
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Dim TAN_NAME

TAN_NAME = Range("TAN_NAME").Value

ActiveWorkbook.SaveAs Filename:="X:\Anca\#-Groupe\A_DETP_OP\0. EQUIPE\Indicateurs\2025\Logs Données Mensuelles\" & TAN_NAME & ".xlsx"

ActiveWorkbook.Close

Windows("Indicateurs_V2.1.1.xlsm").Activate

Sheets("LOG_ANNUEL").Visible = Hide

MsgBox ("LOG ANNUEL EDITE!")

'CLEAR TOT ANN

Sheets("LOG_ANNUEL").Visible = True
Sheets("LOG_ANNUEL").Select

    Range("E8:P37").Select
    Selection.ClearContents

    Range("E40:P65").Select
    Selection.ClearContents
    
    Range("E67:P74").Select
    Selection.ClearContents
    
    Range("E101:P162").Select
    Selection.ClearContents
    
    Range("E165:P165").Select
    Selection.ClearContents
    
    Range("E168:P168").Select
    Selection.ClearContents
    
    Range("E170:P170").Select
    Selection.ClearContents
    
    Range("E179:P222").Select
    Selection.ClearContents
    
    Range("E225:P231").Select
    Selection.ClearContents
    
    Range("E234:P240").Select
    Selection.ClearContents
    
    Range("E244:P254").Select
    Selection.ClearContents
    
    Range("E257:P280").Select
    Selection.ClearContents
    
    Range("E283:P283").Select
    Selection.ClearContents

    
Sheets("LOG_ANNUEL").Visible = False
Sheets("SAISIE").Select

Sheets("SAISIE").Select

End Sub
Sub LOG_MENS()

Sheets("LOG_M_OPS").Visible = True
Sheets("LOG_M_OPS").Copy
        
    Range("C2:AJ276").Select
    Selection.Copy
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False


Dim LOG_NAME

LOG_NAME = Range("LOG_NAME").Value

ActiveWorkbook.SaveAs Filename:="X:\Anca\#-Groupe\A_DETP_OP\0. EQUIPE\Indicateurs\2025\Logs Données Mensuelles\" & LOG_NAME & ".xlsx"

ActiveWorkbook.Close

Sheets("LOG").Visible = Hide

MsgBox ("LOG MENSUEL EDITE!")


End Sub
Private Sub LOG_MLG()

Sheets("LOG_OPS").Visible = True
Sheets("LOG_OPS").Copy

    Range("A1:H120").Select
    Selection.Copy
    Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Dim MLG_NAME

MLG_NAME = Range("MLG_NAME").Value

ActiveWorkbook.SaveAs Filename:="X:\Anca\#-Groupe\A_DETP_OP\0. EQUIPE\Indicateurs\2025\Logs Logistique\" & MLG_NAME & ".xlsx"

ActiveWorkbook.Close

Sheets("LOG_OPS").Visible = Hide

Range("ET_MLG").Value = 1

MsgBox ("LOG LOGISTIQUE EDITE!")

End Sub
Private Sub DATA_MAIN_RESET()

Range("DATA_VAL_MAIN").Value = ""

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

Range("TRA_AFR").Value = ""
Range("TRA_AVP").Value = ""
Range("TRA_MZS").Value = ""
Range("TRA_GEH").Value = ""
Range("TRA_APR").Value = ""
Range("TRA_SEA").Value = ""
Range("TRA_SFS").Value = ""
Range("TRA_AZH").Value = ""
Range("TRA_HLI").Value = ""
Range("TRA_MCM").Value = ""
Range("TRA_MCR").Value = ""
Range("TRA_NI").Value = ""

Range("TRA_MAFR").Value = ""
Range("TRA_MAVP").Value = ""
Range("TRA_MMZS").Value = ""
Range("TRA_MGEH").Value = ""
Range("TRA_MAPR").Value = ""
Range("TRA_MSEA").Value = ""
Range("TRA_MSFS").Value = ""
Range("TRA_MAZH").Value = ""
Range("TRA_MHLI").Value = ""
Range("TRA_MMCM").Value = ""
Range("TRA_MMCR").Value = ""
Range("TRA_MNI").Value = ""

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
Sub ARCH()

Range("MOIS_X").Value = Range("MOIS_M1").Value

Call INP_ARCH 'OK
Call INTER_ARCH 'OK
Call ETAT_ARCH 'OK
Call DIV_ARCH 'OK
Call FRET_ARCH 'OK
Call DVS_ARCH 'OK
Call TRACT_ARCH 'OK
Call DEV_ARCH 'OK
Call FACT_ARCH 'ok

Call CNR_TRA_ARCH 'OK
Call PNR_TRA_ARCH 'ok
Call MNR_TRA_ARCH 'ok
Call MAE_TRA_ARCH 'ok
Call MPA_TRA_ARCH 'ok
Call MSH_TRA_ARCH 'ok
Call PBS_TRA_ARCH 'OK

Call FULL_CLEAR 'OK

ActiveWorkbook.Save

End Sub

Sub CNR_TRA_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MCNR, AMX_CNR1, AMX_CNR2, AMX_CNR3, AMX_CNR4, AMX_CNR5, AMX_CNR6, AMX_CNR7, AMX_CNR8, AMX_CNR9, AMX_CNR10, AMX_CNR11, AMX_CNR12

AMX_MCNR = Range("ADD_MX_56").Value
AMX_CNR1 = Range("ADD_MX_57").Value
AMX_CNR2 = Range("ADD_MX_58").Value
AMX_CNR3 = Range("ADD_MX_59").Value
AMX_CNR4 = Range("ADD_MX_60").Value
AMX_CNR5 = Range("ADD_MX_61").Value
AMX_CNR6 = Range("ADD_MX_62").Value
AMX_CNR7 = Range("ADD_MX_63").Value
AMX_CNR8 = Range("ADD_MX_64").Value
AMX_CNR9 = Range("ADD_MX_65").Value
AMX_CNR10 = Range("ADD_MX_66").Value
AMX_CNR11 = Range("ADD_MX_67").Value
AMX_CNR12 = Range("ADD_MX_68").Value

Sheets("LOG_ANNUEL").Range(AMX_MCNR) = Range("TOT_56").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR1) = Range("TOT_57").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR2) = Range("TOT_58").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR3) = Range("TOT_59").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR4) = Range("TOT_60").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR5) = Range("TOT_61").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR6) = Range("TOT_62").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR7) = Range("TOT_63").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR8) = Range("TOT_64").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR9) = Range("TOT_65").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR10) = Range("TOT_66").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR11) = Range("TOT_67").Text
Sheets("LOG_ANNUEL").Range(AMX_CNR12) = Range("TOT_68").Text

Sheets("SAISIE").Select

End Sub
Sub PBS_TRA_ARCH()


'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MPBS, AMX_PBS1, AMX_PBS2, AMX_PBS3, AMX_PBS4, AMX_PBS5, AMX_PBS6, AMX_PBS7, AMX_PBS8, AMX_PBS9, AMX_PBS10, AMX_PBS11, AMX_PBS12

AMX_MPBS = Range("ADD_MX_112").Value
AMX_PBS1 = Range("ADD_MX_113").Value
AMX_PBS2 = Range("ADD_MX_114").Value
AMX_PBS3 = Range("ADD_MX_115").Value
AMX_PBS4 = Range("ADD_MX_116").Value
AMX_PBS5 = Range("ADD_MX_117").Value
AMX_PBS6 = Range("ADD_MX_118").Value
AMX_PBS7 = Range("ADD_MX_119").Value
AMX_PBS8 = Range("ADD_MX_120").Value
AMX_PBS9 = Range("ADD_MX_121").Value
AMX_PBS10 = Range("ADD_MX_122").Value
AMX_PBS11 = Range("ADD_MX_123").Value
AMX_PBS12 = Range("ADD_MX_124").Value

Range(AMX_MPBS) = Range("TOT_112").Text
Range(AMX_PBS1) = Range("TOT_113").Text
Range(AMX_PBS2) = Range("TOT_114").Text
Range(AMX_PBS3) = Range("TOT_115").Text
Range(AMX_PBS4) = Range("TOT_116").Text
Range(AMX_PBS5) = Range("TOT_117").Text
Range(AMX_PBS6) = Range("TOT_118").Text
Range(AMX_PBS7) = Range("TOT_119").Text
Range(AMX_PBS8) = Range("TOT_120").Text
Range(AMX_PBS9) = Range("TOT_121").Text
Range(AMX_PBS10) = Range("TOT_122").Text
Range(AMX_PBS11) = Range("TOT_123").Text
Range(AMX_PBS12) = Range("TOT_124").Text


End Sub
Sub MAE_TRA_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MMAE, AMX_MAE1, AMX_MAE2, AMX_MAE3, AMX_MAE4, AMX_MAE5, AMX_MAE6, AMX_MAE7, AMX_MAE8, AMX_MAE9, AMX_MAE10, AMX_MAE11, AMX_MAE12

AMX_MMAE = Range("ADD_MX_82").Value
AMX_MAE1 = Range("ADD_MX_83").Value
AMX_MAE2 = Range("ADD_MX_84").Value
AMX_MAE3 = Range("ADD_MX_85").Value
AMX_MAE4 = Range("ADD_MX_86").Value
AMX_MAE5 = Range("ADD_MX_87").Value
AMX_MAE6 = Range("ADD_MX_88").Value
AMX_MAE7 = Range("ADD_MX_89").Value
AMX_MAE8 = Range("ADD_MX_90").Value
AMX_MAE9 = Range("ADD_MX_91").Value
AMX_MAE10 = Range("ADD_MX_92").Value
AMX_MAE11 = Range("ADD_MX_93").Value
AMX_MAE12 = Range("ADD_MX_94").Value

Range(AMX_MMAE) = Range("TOT_82").Text
Range(AMX_MAE1) = Range("TOT_83").Text
Range(AMX_MAE2) = Range("TOT_84").Text
Range(AMX_MAE3) = Range("TOT_85").Text
Range(AMX_MAE4) = Range("TOT_86").Text
Range(AMX_MAE5) = Range("TOT_87").Text
Range(AMX_MAE6) = Range("TOT_88").Text
Range(AMX_MAE7) = Range("TOT_89").Text
Range(AMX_MAE8) = Range("TOT_90").Text
Range(AMX_MAE9) = Range("TOT_91").Text
Range(AMX_MAE10) = Range("TOT_92").Text
Range(AMX_MAE11) = Range("TOT_93").Text
Range(AMX_MAE12) = Range("TOT_94").Text

Sheets("SAISIE").Select

End Sub
Sub MNR_TRA_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MMNR, AMX_MNR1, AMX_MNR2, AMX_MNR3, AMX_MNR4, AMX_MNR5, AMX_MNR6, AMX_MNR7, AMX_MNR8, AMX_MNR9, AMX_MNR10, AMX_MNR11, AMX_MNR12

AMX_MMNR = Range("ADD_MX_69").Value
AMX_MNR1 = Range("ADD_MX_70").Value
AMX_MNR2 = Range("ADD_MX_71").Value
AMX_MNR3 = Range("ADD_MX_72").Value
AMX_MNR4 = Range("ADD_MX_73").Value
AMX_MNR5 = Range("ADD_MX_74").Value
AMX_MNR6 = Range("ADD_MX_75").Value
AMX_MNR7 = Range("ADD_MX_76").Value
AMX_MNR8 = Range("ADD_MX_77").Value
AMX_MNR9 = Range("ADD_MX_78").Value
AMX_MNR10 = Range("ADD_MX_79").Value
AMX_MNR11 = Range("ADD_MX_80").Value
AMX_MNR12 = Range("ADD_MX_81").Value

Range(AMX_MMNR) = Range("TOT_69").Text
Range(AMX_MNR1) = Range("TOT_70").Text
Range(AMX_MNR2) = Range("TOT_71").Text
Range(AMX_MNR3) = Range("TOT_72").Text
Range(AMX_MNR4) = Range("TOT_73").Text
Range(AMX_MNR5) = Range("TOT_74").Text
Range(AMX_MNR6) = Range("TOT_75").Text
Range(AMX_MNR7) = Range("TOT_76").Text
Range(AMX_MNR8) = Range("TOT_77").Text
Range(AMX_MNR9) = Range("TOT_78").Text
Range(AMX_MNR10) = Range("TOT_79").Text
Range(AMX_MNR11) = Range("TOT_80").Text
Range(AMX_MNR12) = Range("TOT_81").Text

Sheets("SAISIE").Select

End Sub
Sub PNR_TRA_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MPNR, AMX_PNR1, AMX_PNR2, AMX_PNR3, AMX_PNR4, AMX_PNR5, AMX_PNR6, AMX_PNR7, AMX_PNR8, AMX_PNR9, AMX_PNR10, AMX_PNR11, AMX_PNR12

AMX_MPNR = Range("ADD_MX_125").Value
AMX_PNR1 = Range("ADD_MX_126").Value
AMX_PNR2 = Range("ADD_MX_127").Value
AMX_PNR3 = Range("ADD_MX_128").Value
AMX_PNR4 = Range("ADD_MX_129").Value
AMX_PNR5 = Range("ADD_MX_130").Value
AMX_PNR6 = Range("ADD_MX_131").Value
AMX_PNR7 = Range("ADD_MX_132").Value
AMX_PNR8 = Range("ADD_MX_133").Value
AMX_PNR9 = Range("ADD_MX_134").Value
AMX_PNR10 = Range("ADD_MX_135").Value
AMX_PNR11 = Range("ADD_MX_136").Value
AMX_PNR12 = Range("ADD_MX_137").Value

Range(AMX_MPNR) = Range("TOT_125").Text
Range(AMX_PNR1) = Range("TOT_126").Text
Range(AMX_PNR2) = Range("TOT_127").Text
Range(AMX_PNR3) = Range("TOT_128").Text
Range(AMX_PNR4) = Range("TOT_129").Text
Range(AMX_PNR5) = Range("TOT_130").Text
Range(AMX_PNR6) = Range("TOT_131").Text
Range(AMX_PNR7) = Range("TOT_132").Text
Range(AMX_PNR8) = Range("TOT_133").Text
Range(AMX_PNR9) = Range("TOT_134").Text
Range(AMX_PNR10) = Range("TOT_135").Text
Range(AMX_PNR11) = Range("TOT_136").Text
Range(AMX_PNR12) = Range("TOT_137").Text

Sheets("SAISIE").Select

End Sub
Sub MPA_TRA_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MMPA, AMX_MPA1, AMX_MPA2, AMX_MPA3, AMX_MPA4, AMX_MPA5, AMX_MPA6, AMX_MPA7, AMX_MPA8, AMX_MPA9, AMX_MPA10, AMX_MPA11, AMX_MPA12

AMX_MMPA = Range("ADD_MX_95").Value
AMX_MPA1 = Range("ADD_MX_96").Value
AMX_MPA2 = Range("ADD_MX_97").Value
AMX_MPA3 = Range("ADD_MX_98").Value
AMX_MPA4 = Range("ADD_MX_99").Value
AMX_MPA5 = Range("ADD_MX_100").Value
AMX_MPA6 = Range("ADD_MX_101").Value
AMX_MPA7 = Range("ADD_MX_102").Value
AMX_MPA8 = Range("ADD_MX_103").Value
AMX_MPA9 = Range("ADD_MX_104").Value
AMX_MPA10 = Range("ADD_MX_105").Value
AMX_MPA11 = Range("ADD_MX_106").Value
AMX_MPA12 = Range("ADD_MX_107").Value

Range(AMX_MMPA) = Range("TOT_95").Text
Range(AMX_MPA1) = Range("TOT_96").Text
Range(AMX_MPA2) = Range("TOT_97").Text
Range(AMX_MPA3) = Range("TOT_98").Text
Range(AMX_MPA4) = Range("TOT_99").Text
Range(AMX_MPA5) = Range("TOT_100").Text
Range(AMX_MPA6) = Range("TOT_101").Text
Range(AMX_MPA7) = Range("TOT_102").Text
Range(AMX_MPA8) = Range("TOT_103").Text
Range(AMX_MPA9) = Range("TOT_104").Text
Range(AMX_MPA10) = Range("TOT_105").Text
Range(AMX_MPA11) = Range("TOT_106").Text
Range(AMX_MPA12) = Range("TOT_107").Text

Sheets("SAISIE").Select

End Sub
Sub MSH_TRA_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MMSH, AMX_MSH1, AMX_MSH2, AMX_MSH3, AMX_MSH4, AMX_MSH5, AMX_MSH6, AMX_MSH7, AMX_MSH8, AMX_MSH9, AMX_MSH10, AMX_MSH11, AMX_MSH12

AMX_MMSH = Range("ADD_MX_109").Value
AMX_MSH1 = Range("ADD_MX_195").Value
AMX_MSH2 = Range("ADD_MX_196").Value
AMX_MSH3 = Range("ADD_MX_197").Value
AMX_MSH4 = Range("ADD_MX_198").Value
AMX_MSH5 = Range("ADD_MX_199").Value
AMX_MSH6 = Range("ADD_MX_200").Value
AMX_MSH7 = Range("ADD_MX_201").Value
AMX_MSH8 = Range("ADD_MX_202").Value
AMX_MSH9 = Range("ADD_MX_203").Value
AMX_MSH10 = Range("ADD_MX_204").Value
AMX_MSH11 = Range("ADD_MX_205").Value
AMX_MSH12 = Range("ADD_MX_206").Value

Range(AMX_MMSH) = Range("TOT_109").Text
Range(AMX_MSH1) = Range("TOT_195").Text
Range(AMX_MSH2) = Range("TOT_196").Text
Range(AMX_MSH3) = Range("TOT_197").Text
Range(AMX_MSH4) = Range("TOT_198").Text
Range(AMX_MSH5) = Range("TOT_199").Text
Range(AMX_MSH6) = Range("TOT_200").Text
Range(AMX_MSH7) = Range("TOT_201").Text
Range(AMX_MSH8) = Range("TOT_202").Text
Range(AMX_MSH9) = Range("TOT_203").Text
Range(AMX_MSH10) = Range("TOT_204").Text
Range(AMX_MSH11) = Range("TOT_205").Text
Range(AMX_MSH12) = Range("TOT_206").Text

Sheets("SAISIE").Select

End Sub

Sub DEV_ARCH()


'TRANSFERT VERS DATA ANNUELLES

Dim AMX_C1, AMX_C2, AMX_C3, AMX_C4, AMX_C5, AMX_C6, AMX_C7, AMX_C8, AMX_C9, AMX_C10, AMX_C11, AMX_C12

AMX_C1 = Range("ADD_MX_43").Value
AMX_C2 = Range("ADD_MX_44").Value
AMX_C3 = Range("ADD_MX_45").Value
AMX_C4 = Range("ADD_MX_46").Value
AMX_C5 = Range("ADD_MX_47").Value
AMX_C6 = Range("ADD_MX_48").Value
AMX_C7 = Range("ADD_MX_49").Value
AMX_C8 = Range("ADD_MX_50").Value
AMX_C9 = Range("ADD_MX_51").Value
AMX_C10 = Range("ADD_MX_52").Value
AMX_C11 = Range("ADD_MX_53").Value
AMX_C12 = Range("ADD_MX_54").Value

Range(AMX_C1) = Range("TOTM_43").Text
Range(AMX_C2) = Range("TOTM_44").Text
Range(AMX_C3) = Range("TOTM_45").Text
Range(AMX_C4) = Range("TOTM_46").Text
Range(AMX_C5) = Range("TOTM_47").Text
Range(AMX_C6) = Range("TOTM_48").Text
Range(AMX_C7) = Range("TOTM_49").Text
Range(AMX_C8) = Range("TOTM_50").Text
Range(AMX_C9) = Range("TOTM_51").Text
Range(AMX_C10) = Range("TOTM_52").Text
Range(AMX_C11) = Range("TOTM_53").Text
Range(AMX_C12) = Range("TOTM_54").Text

Sheets("SAISIE").Select

End Sub
Sub FACT_ARCH()


'TRANSFERT VERS DATA ANNUELLES

Dim AMX_SAT, AMX_CGP

AMX_CGP = Range("ADD_MX_108").Value
AMX_SAT = Range("ADD_MX_110").Value

Range(AMX_CGP) = Range("TOT_108").Text
Range(AMX_SAT) = Range("TOT_110").Text


End Sub
Sub TRACT_ARCH()


'TRANSFERT VERS DATA ANNUELLES

Dim AMX_MDE, AMX_ACF, AMX_CDC, AMX_DK23, AMX_TVX, AMX_SKU, AMX_PAN, AMX_REP, AMX_CGEN, AMX_CCOM, AMX_CVH

AMX_MDE = Range("ADD_MX_32").Value
AMX_ACF = Range("ADD_MX_33").Value
AMX_CDC = Range("ADD_MX_34").Value
AMX_DK23 = Range("ADD_MX_35").Value
AMX_TVX = Range("ADD_MX_36").Value
AMX_SKU = Range("ADD_MX_37").Value
AMX_PAN = Range("ADD_MX_38").Value
AMX_REP = Range("ADD_MX_39").Value
AMX_CGEN = Range("ADD_MX_41").Value
AMX_CCOM = Range("ADD_MX_42").Value
AMX_CVH = Range("ADD_MX_144").Value

Range(AMX_MDE) = Range("TOT_32").Text
Range(AMX_ACF) = Range("TOT_33").Text
Range(AMX_CDC) = Range("TOT_34").Text
Range(AMX_DK23) = Range("TOT_35").Text
Range(AMX_TVX) = Range("TOT_36").Text
Range(AMX_SKU) = Range("TOT_37").Text
Range(AMX_PAN) = Range("TOT_38").Text
Range(AMX_REP) = Range("TOT_39").Text
Range(AMX_CGEN) = Range("TOT_41").Text
Range(AMX_CCOM) = Range("TOT_42").Text
Range(AMX_CVH) = Range("TOT_144").Text

Sheets("SAISIE").Select

End Sub
Sub DVS_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_BDD, AMX_MDG, AMX_EMT

AMX_BDD = Range("ADD_MX_24").Value
AMX_MDG = Range("ADD_MX_25").Value
AMX_EMT = Range("ADD_MX_111").Value

Range(AMX_BDD) = Range("TOT_24").Text
Range(AMX_MDG) = Range("TOT_25").Text
Range(AMX_EMT) = Range("TOT_111").Text

Sheets("SAISIE").Select

End Sub
Sub FRET_ARCH()


'TRANSFERT VERS DATA ANNUELLES

Dim AMX_ENT, AMX_SOR, AMX_M30, AMX_P30, AMX_1H

AMX_ENT = Range("ADD_MX_26").Value
AMX_SOR = Range("ADD_MX_27").Value
AMX_M30 = Range("ADD_MX_28").Value
AMX_P30 = Range("ADD_MX_29").Value
AMX_1H = Range("ADD_MX_30").Value

Range(AMX_ENT) = Range("TOT_26").Text
Range(AMX_SOR) = Range("TOT_27").Text
Range(AMX_M30) = Range("TOT_28").Text
Range(AMX_P30) = Range("TOT_29").Text
Range(AMX_1H) = Range("TOT_30").Text


Sheets("SAISIE").Select

End Sub
Sub DIV_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_DPU, AMX_SGS, AMX_FOD, AMX_FOA, AMX_FPA

AMX_DPU = Range("ADD_MX_19").Value
AMX_SGS = Range("ADD_MX_20").Value
AMX_FOD = Range("ADD_MX_21").Value
AMX_FOA = Range("ADD_MX_22").Value
AMX_FPA = Range("ADD_MX_23").Value

Range(AMX_DPU) = Range("TOT_19").Text
Range(AMX_SGS) = Range("TOT_20").Text
Range(AMX_FOD) = Range("TOT_21").Text
Range(AMX_FOA) = Range("TOT_22").Text
Range(AMX_FPA) = Range("TOT_23").Text

Sheets("SAISIE").Select

End Sub
Sub ETAT_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_ETV, AMX_ETA, AMX_ETC

AMX_ETV = Range("ADD_MX_16").Value
AMX_ETA = Range("ADD_MX_17").Value
AMX_ETC = Range("ADD_MX_18").Value

Range(AMX_ETV) = Range("TOT_16").Text
Range(AMX_ETA) = Range("TOT_17").Text
Range(AMX_ETC) = Range("TOT_18").Text

Sheets("SAISIE").Select

End Sub
Sub INTER_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_BAL, AMX_INF, AMX_AUT2

AMX_BAL = Range("ADD_MX_12").Value
AMX_INF = Range("ADD_MX_13").Value
AMX_AUT2 = Range("ADD_MX_14").Value

Range(AMX_INF) = Range("TOT_13").Text
Range(AMX_BAL) = Range("TOT_12").Text
Range(AMX_AUT2) = Range("TOT_14").Text

Sheets("SAISIE").Select

End Sub
Sub INP_ARCH()

'TRANSFERT VERS DATA ANNUELLES

Dim AMX_IP, AMX_INPN, AMX_ISB, AMX_BA, AMX_AA, AMX_EV, AMX_EA, AMX_EC, AMX_IN, AMX_TV, AMX_DI

AMX_IP = Range("ADD_MX_1").Value
AMX_INPN = Range("ADD_MX_2").Value
AMX_ISB = Range("ADD_MX_3").Value
AMX_BA = Range("ADD_MX_4").Value
AMX_AA = Range("ADD_MX_5").Value
AMX_EV = Range("ADD_MX_6").Value
AMX_EA = Range("ADD_MX_7").Value
AMX_EC = Range("ADD_MX_8").Value
AMX_IN = Range("ADD_MX_9").Value
AMX_TV = Range("ADD_MX_10").Value
AMX_DI = Range("ADD_MX_11").Value

Range(AMX_IP) = Range("TOT_1").Text
Range(AMX_INPN) = Range("TOT_2").Text
Range(AMX_ISB) = Range("TOT_3").Text
Range(AMX_BA) = Range("TOT_4").Text
Range(AMX_AA) = Range("TOT_5").Text
Range(AMX_EV) = Range("TOT_6").Text
Range(AMX_EA) = Range("TOT_7").Text
Range(AMX_EC) = Range("TOT_8").Text
Range(AMX_IN) = Range("TOT_9").Text
Range(AMX_TV) = Range("TOT_10").Text
Range(AMX_DI) = Range("TOT_11").Text


Sheets("SAISIE").Select

End Sub
Private Sub FULL_CLEAR()

Sheets("LOG_M").Select

    Range("E7:AI36").Select
    Selection.ClearContents

    Range("E39:AI64").Select
    Selection.ClearContents

    Range("E66:AI73").Select
    Selection.ClearContents

    Range("E75:AI75").Select
    Selection.ClearContents

    Range("E100:AI161").Select
    Selection.ClearContents

    Range("E164:AI164").Select
    Selection.ClearContents

    Range("E167:AI167").Select
    Selection.ClearContents

    Range("E169:AI169").Select
    Selection.ClearContents

    Range("E171:AI214").Select
    Selection.ClearContents

    Range("E224:AI224").Select
    Selection.ClearContents

    Range("E228:AI234").Select
    Selection.ClearContents
    
Sheets("LOG_M").Visible = False
Sheets("SAISIE").Select

End Sub

' ============================================================
' MODULE : S_FAC.frm
' ============================================================

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

' ============================================================
' MODULE : TRAF_2.frm
' ============================================================

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
