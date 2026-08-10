' Module : HDG.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

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
