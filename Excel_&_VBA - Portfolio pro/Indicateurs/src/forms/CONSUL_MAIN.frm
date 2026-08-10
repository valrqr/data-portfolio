' Module : CONSUL_MAIN.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

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
