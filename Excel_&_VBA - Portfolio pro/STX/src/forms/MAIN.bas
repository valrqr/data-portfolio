' Module : MAIN.bas
' Extrait de : STX_V1.4.xlsm

Sub MAIN_L()

STX.Show

End Sub

Private Sub RANDOM2_AVP()

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
Private Sub RANDOM2_GEH()

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
Private Sub RANDOM2_MZS()

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
Private Sub RANDOM2_AFR()

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
Sub MAIN_RANDO()

Range("AFR_1").ClearContents
Range("AFR_2").ClearContents

Range("AVP_1").ClearContents
Range("AVP_2").ClearContents

Range("MZS_1").ClearContents
Range("MZS_2").ClearContents

Range("GEH_1").ClearContents
Range("GEH_2").ClearContents

Do Until Range("CTL_AFR").Value = 1

Call RANDOM2_AFR

Loop

Do Until Range("CTL_AVP").Value = 1

Call RANDOM2_AVP

Loop

Do Until Range("CTL_GEH").Value = 1

Call RANDOM2_GEH

Loop

Do Until Range("CTL_MZS").Value = 1

Call RANDOM2_MZS

Loop

End Sub
Sub RAZ()

Sheets("DATA_AFR").Range("$B$1:$I$10000").ClearContents
Sheets("DATA_AVP").Range("$B$1:$I$10000").ClearContents
Sheets("DATA_GEH").Range("$B$1:$I$10000").ClearContents
Sheets("DATA_MZS").Range("$B$1:$I$10000").ClearContents
Sheets("DATA").Range("$A$1:$H$10000").ClearContents
Sheets("DATA").Range("$A$1:$H$10000").ClearContents

Range("AFR_1").ClearContents
Range("AFR_2").ClearContents

Range("AVP_1").ClearContents
Range("AVP_2").ClearContents

Range("MZS_1").ClearContents
Range("MZS_2").ClearContents

Range("GEH_1").ClearContents
Range("GEH_2").ClearContents

End Sub

Sub Test()

    Dim AppWord As Object
    Dim Doc As Object

    'crée une instance de Word
    Set AppWord = CreateObject("Word.Application")

    With AppWord

        .Visible = True

        'ouvre un document
        Set Doc = .Documents.Open("X:\...\Outils Surveillance Touchée Avion\Planning_ST_Modele.docx")

        With Doc

        .Bookmarks("MOIS").Range.Text = Range("EXT_MOIS").Value
        .Bookmarks("ANNEE").Range.Text = Range("DATA_ANN").Value
        
        .Bookmarks("D1AFR").Range.Text = Range("D1AFR").Value
        .Bookmarks("D2AFR").Range.Text = Range("D2AFR").Value
        .Bookmarks("D3AFR").Range.Text = Range("D3AFR").Value

        .Bookmarks("D1AVP").Range.Text = Range("D1AVP").Value
        .Bookmarks("D2AVP").Range.Text = Range("D2AVP").Value
        .Bookmarks("D3AVP").Range.Text = Range("D3AVP").Value
        
        .Bookmarks("D1GEH").Range.Text = Range("D1GEH").Value
        .Bookmarks("D2GEH").Range.Text = Range("D2GEH").Value
        .Bookmarks("D3GEH").Range.Text = Range("D3GEH").Value
        
        .Bookmarks("D1MZS").Range.Text = Range("D1MZS").Value
        .Bookmarks("D2MZS").Range.Text = Range("D2MZS").Value
        .Bookmarks("D3MZS").Range.Text = Range("D3MZS").Value
        
        .Bookmarks("T1AFR").Range.Text = Range("T1AFR").Value
        .Bookmarks("T2AFR").Range.Text = Range("T2AFR").Value
        .Bookmarks("T3AFR").Range.Text = Range("T3AFR").Value
        
        .Bookmarks("T1AVP").Range.Text = Range("T1AVP").Value
        .Bookmarks("T2AVP").Range.Text = Range("T2AVP").Value
        .Bookmarks("T3AVP").Range.Text = Range("T3AVP").Value
        
        .Bookmarks("T1GEH").Range.Text = Range("T1GEH").Value
        .Bookmarks("T2GEH").Range.Text = Range("T2GEH").Value
        .Bookmarks("T3GEH").Range.Text = Range("T3GEH").Value
        
        .Bookmarks("T1MZS").Range.Text = Range("T1MZS").Value
        .Bookmarks("T2MZS").Range.Text = Range("T2MZS").Value
        .Bookmarks("T3MZS").Range.Text = Range("T3MZS").Value
        
        .Bookmarks("H1AFR").Range.Text = Range("H1AFR").Text
        .Bookmarks("H2AFR").Range.Text = Range("H2AFR").Text
        .Bookmarks("H3AFR").Range.Text = Range("H3AFR").Text
        
        .Bookmarks("H1AVP").Range.Text = Range("H1AVP").Text
        .Bookmarks("H2AVP").Range.Text = Range("H2AVP").Text
        .Bookmarks("H3AVP").Range.Text = Range("H3AVP").Text
        
        .Bookmarks("H1GEH").Range.Text = Range("H1GEH").Text
        .Bookmarks("H2GEH").Range.Text = Range("H2GEH").Text
        .Bookmarks("H3GEH").Range.Text = Range("H3GEH").Text
        
        .Bookmarks("H1MZS").Range.Text = Range("H1MZS").Text
        .Bookmarks("H2MZS").Range.Text = Range("H2MZS").Text
        .Bookmarks("H3MZS").Range.Text = Range("H3MZS").Text
        
        

        End With

    End With

End Sub
