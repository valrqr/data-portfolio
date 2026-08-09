' Module : S_ANX.bas
' Extrait de : Indicateurs_V2.1.1.xlsm

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
