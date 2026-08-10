' Module : DEV.frm
' Extrait de : Indicateurs_V2.1.1.xlsm

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
