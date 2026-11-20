' Module : STX.frm
' Extrait de : STX_V1.4.xlsm

Private Sub CB_CNX_Click()

STX.Hide

End Sub

Private Sub CB_MOIS_Change()

Sheets("DATA2").Range("DATA_MOIS").Value = CB_MOIS.ListIndex + 1

End Sub

Private Sub CB_VAL_Click()

'Mise en place des databases avec tri par sens, heures et assistants

Sheets("DATA").Range("A:A,E:E,G:G,H:H,I:I,K:K,M:M,O:O,P:P,Q:Q,R:R,S:S,T:T,U:U,V:V,W:W,X:X,Y:Y,Z:Z,AA:AA,AB:AB,AC:AC,AD:AD,AE:AE,AF:AF,AG:AG,AH:AH,AI:AI,AJ:AJ,AL:AL,AM:AM,AN:AN,AO:AO,AP:AP,AQ:AQ,AR:AR,AS:AS,AT:AT,AU:AU,AV:AV").Delete

Sheets("DATA").Range("$A$1:$H$10000").AutoFilter Field:=2, Criteria1:=">=08:00" _
        , Operator:=xlAnd, Criteria2:="<=19:00"

Sheets("DATA").Range("$A$1:$H$10000").AutoFilter Field:=8, Criteria1:="AFR"

Sheets("DATA").Select
Range("$A$1:$H$10000").Select
    Selection.Copy
Sheets("DATA_AFR").Select
Range("DATA_AFR").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.NumberFormat = "0"

Sheets("DATA").Range("$A$1:$H$10000").AutoFilter Field:=8, Criteria1:="AVP"

Sheets("DATA").Select
Range("$A$1:$H$10000").Select
    Selection.Copy
Sheets("DATA_AVP").Select
Range("DATA_AVP").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.NumberFormat = "0"
    
Sheets("DATA").Range("$A$1:$H$10000").AutoFilter Field:=8, Criteria1:="MZS"

Sheets("DATA").Select
Range("$A$1:$H$10000").Select
    Selection.Copy
Sheets("DATA_MZS").Select
Range("DATA_MZS").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.NumberFormat = "0"

Sheets("DATA").Range("$A$1:$H$10000").AutoFilter Field:=8, Criteria1:="=CCM", _
        Operator:=xlOr, Criteria2:="=ALZ"

Sheets("DATA").Select
Range("$A$1:$H$10000").Select
    Selection.Copy
Sheets("DATA_GEH").Select
Range("DATA_GEH").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Selection.NumberFormat = "0"

Do Until Range("CTL_AFR").Value = 1

Call RANDOM_AFR

Loop

Do Until Range("CTL_AVP").Value = 1

Call RANDOM_AVP

Loop

Do Until Range("CTL_GEH").Value = 1

Call RANDOM_GEH

Loop

Do Until Range("CTL_MZS").Value = 1

Call RANDOM_MZS

Loop

STX.Hide

End Sub

Private Sub RANDOM_AVP()

Sheets("MENU").Select

'Random AVP

Dim plage_avp As Range, cel_avp As Range, alea_avp As Double

Dim AVP_MIN, AVP_MAX

AVP_MIN = Range("AVP_MIN").Value
AVP_MAX = Range("AVP_MAX").Value

Set plage_avp = Sheets("MENU").Range("A23:A24") 'modifiable
plage_avp.Value = ""
If plage_avp.Count > AVP_MAX Then Exit Sub
Randomize
For Each cel_avp In plage_avp
1 alea_avp = WorksheetFunction.RandBetween(AVP_MIN, AVP_MAX)
If Application.CountIf(plage_avp, alea_avp) Then GoTo 1 Else cel_avp = alea_avp
Next

End Sub
Private Sub RANDOM_GEH()

Sheets("MENU").Select

'Random GEH

Dim plage_geh As Range, cel_geh As Range, alea_geh As Double

Dim GEH_MIN, GEH_MAX

GEH_MIN = Range("GEH_MIN").Value
GEH_MAX = Range("GEH_MAX").Value

Set plage_geh = Sheets("MENU").Range("A41:A42") 'modifiable
plage_geh.Value = ""
If plage_geh.Count > GEH_MAX Then Exit Sub
Randomize
For Each cel_geh In plage_geh
1 alea_geh = WorksheetFunction.RandBetween(GEH_MIN, GEH_MAX)
If Application.CountIf(plage_geh, alea_geh) Then GoTo 1 Else cel_geh = alea_geh
Next

End Sub
Private Sub RANDOM_MZS()

Sheets("MENU").Select

'Random MZS

Dim plage_mzs As Range, cel_mzs As Range, alea_mzs As Double

Dim MZS_MIN, MZS_MAX

MZS_MIN = Range("MZS_MIN").Value
MZS_MAX = Range("MZS_MAX").Value

Set plage_mzs = Sheets("MENU").Range("A32:A33") 'modifiable
plage_mzs.Value = ""
If plage_mzs.Count > MZS_MAX Then Exit Sub
Randomize
For Each cel_mzs In plage_mzs
1 alea_mzs = WorksheetFunction.RandBetween(MZS_MIN, MZS_MAX)
If Application.CountIf(plage_mzs, alea_mzs) Then GoTo 1 Else cel_mzs = alea_mzs
Next

End Sub
Private Sub RANDOM_AFR()

Sheets("MENU").Select

'Random AFR

Dim plage As Range, cel As Range, alea As Double

Dim AFR_MIN, AFR_MAX

AFR_MIN = Range("AFR_MIN").Value
AFR_MAX = Range("AFR_MAX").Value

Set plage = Sheets("MENU").Range("A14:A15") 'modifiable
plage.Value = ""
If plage.Count > AFR_MAX Then Exit Sub
Randomize
For Each cel In plage
1 alea = WorksheetFunction.RandBetween(AFR_MIN, AFR_MAX)
If Application.CountIf(plage, alea) Then GoTo 1 Else cel = alea
Next

End Sub
