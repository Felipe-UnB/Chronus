VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box2_UPb_Options 
   Caption         =   "Options"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5085
   OleObjectBlob   =   "Box2_UPb_Options.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Box2_UPb_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox3_BlankErrors_Click()
    
    If _
        CheckBox3_BlankErrors.Value = False And _
        CheckBox4_ExtStdErrors.Value = False And _
        CheckBox5_ExtStdRepro.Value = False And _
        CheckBox6_CertExtStd.Value = False _
    Then
    
        MsgBox "You should propagate uncertianties from the blank and the primary standard anaylses to the samples. " & _
            "Otherwise, this will result in samples with unrealistc small uncertainties.", vbOKOnly
    End If
            
End Sub

Private Sub CheckBox4_ExtStdErrors_Click()

    If CheckBox4_ExtStdErrors.Value = True Then
        If CheckBox5_ExtStdRepro.Value = True Then
            CheckBox5_ExtStdRepro.Value = False
        End If
    End If
    
    If _
        CheckBox3_BlankErrors.Value = False And _
        CheckBox4_ExtStdErrors.Value = False And _
        CheckBox5_ExtStdRepro.Value = False And _
        CheckBox6_CertExtStd.Value = False _
    Then
    
        MsgBox "Uncertianties from blank and primary standard anaylses will not be propagated to samples. This will " & _
            "result in samples with unrealistc small uncertainties.", vbOKOnly
    End If

End Sub

Private Sub CheckBox5_ExtStdRepro_Click()

    If CheckBox5_ExtStdRepro.Value = True Then
        If CheckBox4_ExtStdErrors.Value = True Then
            CheckBox4_ExtStdErrors.Value = False
        End If
    End If
    
    If _
        CheckBox3_BlankErrors.Value = False And _
        CheckBox4_ExtStdErrors.Value = False And _
        CheckBox5_ExtStdRepro.Value = False And _
        CheckBox6_CertExtStd.Value = False _
    Then
    
        MsgBox "Uncertianties from blank and primary standard anaylses will not be propagated to samples. This will " & _
            "result in samples with unrealistc small uncertainties.", vbOKOnly
    End If

End Sub

Private Sub CheckBox6_CertExtStd_Click()
    
    If _
        CheckBox3_BlankErrors.Value = False And _
        CheckBox4_ExtStdErrors.Value = False And _
        CheckBox5_ExtStdRepro.Value = False And _
        CheckBox6_CertExtStd.Value = False _
    Then
    
        MsgBox "Uncertianties from blank and primary standard anaylses will not be propagated to samples. This will " & _
            "result in samples with unrealistc small uncertainties.", vbOKOnly
    End If
           
End Sub

Private Sub ComboBox1_ExternalStd_Change()
    
    Dim counter As Integer
    
    Set ChoosenStandard = Box2_UPb_Options.ComboBox1_ExternalStd
           
    For counter = 1 To UBound(UPbStd)
        If UPbStd(counter).StandardName = ChoosenStandard Then
            Box2_UPb_Options.TextBox25 = UPbStd(counter).Mineral
            Box2_UPb_Options.TextBox11_StandardDescription = UPbStd(counter).Description
            Box2_UPb_Options.TextBox12_68Ratio = Val(UPbStd(counter).Ratio68)
            Box2_UPb_Options.TextBox15_68RatioError = Val(UPbStd(counter).Ratio68Error)
            Box2_UPb_Options.TextBox13_75Ratio = Val(UPbStd(counter).Ratio75)
            Box2_UPb_Options.TextBox16_75RatioError = Val(UPbStd(counter).Ratio75Error)
            Box2_UPb_Options.TextBox14_76Ratio = Val(UPbStd(counter).Ratio76)
            Box2_UPb_Options.TextBox17_76RatioError = Val(UPbStd(counter).Ratio76Error)
            Box2_UPb_Options.TextBox23_82Ratio = Val(UPbStd(counter).Ratio82)
            Box2_UPb_Options.TextBox24_82RatioError = Val(UPbStd(counter).Ratio82Error)
            
            If UPbStd(counter).RatioErrors12s = 1 Then
                Box2_UPb_Options.OptionButton1_1sigma = True
            ElseIf UPbStd(counter).RatioErrors12s = 2 Then
                Box2_UPb_Options.OptionButton2_2sigma = True
            End If
            
            If UPbStd(counter).RatioErrorsAbs = True Then
                CheckBox1_Abs = True
            End If
            
            Box2_UPb_Options.TextBox18Uppm = Val(UPbStd(counter).UraniumConc)
            Box2_UPb_Options.TextBox21_UppmError = Val(UPbStd(counter).UraniumConcError)
            Box2_UPb_Options.TextBox19_Thppm = Val(UPbStd(counter).ThoriumConc)
            Box2_UPb_Options.TextBox22_ThppmError = Val(UPbStd(counter).ThoriumConcError)
            
            If UPbStd(counter).ConcErrors12s = 1 Then
                Box2_UPb_Options.OptionButton3 = True
            ElseIf UPbStd(counter).ConcErrors12s = 2 Then
                Box2_UPb_Options.OptionButton4 = True
            End If
                
            If UPbStd(counter).ConcErrorsAbs = True Then
                CheckBox2 = True
            End If

            counter = UBound(UPbStd)
        End If
    Next

End Sub

Private Sub CommandButton1_SaveStandard_Click()

    Dim CellRow As Integer
    Dim UPbNameRng As Range
    Dim OverwriteStd As Boolean
    
    If TW Is Nothing Then
        Call PublicVariables
    End If
    
    OverwriteStd = False
    
    If IsEmpty(Box2_UPb_Options.ComboBox1_ExternalStd) = True Then
        MsgBox "Please, write a name for your new standard. It must be different from any other standard name."
            Box2_UPb_Options.Show
    End If
    
    Set ChoosenStandard = Box2_UPb_Options.ComboBox1_ExternalStd
    
    For Each UPbNameRng In UPbStd_StandardsNames
        
        If ChoosenStandard = UPbNameRng.Value Or ChoosenStandard = UPbNameRng.Text Then 'Using .text and .value properties of range, I'm able to compare ChossenStandard (a name for the standard) to standard names stored in addin, independently, if they are numbers or a strings
            If MsgBox("You can't use the same name for two different standards. Would you like to save these modifications to " & _
                UPbNameRng & "?", vbYesNo) = vbYes Then
                    CellRow = UPbNameRng.Row
                        OverwriteStd = True
            Else
                Exit Sub
            End If
        
        End If
            
    Next
    
    If OverwriteStd = False Then
        With StandardsUPb_TW_Sh
            CellRow = .Range(UPbStd_ColumnStandardName & UPbStd_CHeaderRow + 1).Row
                .Rows(CellRow).Insert Shift:=xlDown
                .Rows(CellRow + 1).EntireRow.Copy
                .Rows(CellRow).PasteSpecial Paste:=xlPasteAllExceptBorders
                .Rows(CellRow).ClearContents
        End With
        
    End If
        
    With StandardsUPb_TW_Sh
        .Range(UPbStd_ColumnStandardName & CellRow) = Box2_UPb_Options.ComboBox1_ExternalStd
        .Range(UPbStd_ColumnMineral & CellRow) = Box2_UPb_Options.TextBox25
        .Range(UPbStd_ColumnDescription & CellRow) = Box2_UPb_Options.TextBox11_StandardDescription
        .Range(UPbStd_ColumnRatio68 & CellRow) = Val(Box2_UPb_Options.TextBox12_68Ratio)
        .Range(UPbStd_ColumnRatio68Error & CellRow) = Val(Box2_UPb_Options.TextBox15_68RatioError)
        .Range(UPbStd_ColumnRatio75 & CellRow) = Val(Box2_UPb_Options.TextBox13_75Ratio)
        .Range(UPbStd_ColumnRatio75Error & CellRow) = Val(Box2_UPb_Options.TextBox16_75RatioError)
        .Range(UPbStd_ColumnRatio76 & CellRow) = Val(Box2_UPb_Options.TextBox14_76Ratio)
        .Range(UPbStd_ColumnRatio76Error & CellRow) = Val(Box2_UPb_Options.TextBox17_76RatioError)
        .Range(UPbStd_ColumnRatio82 & CellRow) = Val(Box2_UPb_Options.TextBox23_82Ratio)
        .Range(UPbStd_ColumnRatio82Error & CellRow) = Val(Box2_UPb_Options.TextBox24_82RatioError)
        .Range(UPbStd_ColumnUraniumConc & CellRow) = Val(Box2_UPb_Options.TextBox18Uppm)
        .Range(UPbStd_ColumnUraniumConcError & CellRow) = Val(Box2_UPb_Options.TextBox21_UppmError)
        .Range(UPbStd_ColumnThoriumConc & CellRow) = Val(Box2_UPb_Options.TextBox19_Thppm)
        .Range(UPbStd_ColumnThoriumConcError & CellRow) = Val(Box2_UPb_Options.TextBox22_ThppmError)

        If Box2_UPb_Options.OptionButton1_1sigma = True Then
            .Range(UPbStd_ColumnRatioErrors12s & CellRow) = 1
        ElseIf Box2_UPb_Options.OptionButton2_2sigma = True Then
            .Range(UPbStd_ColumnRatioErrors12s & CellRow) = 2
        End If

        If CheckBox1_Abs = True Then
            .Range(UPbStd_ColumnRatioErrorsAbs & CellRow) = True
        Else
            .Range(UPbStd_ColumnRatioErrorsAbs & CellRow) = False
        End If
        
        If Box2_UPb_Options.OptionButton3 = True Then
             .Range(UPbStd_ColumnConcErrors12s & CellRow) = 1
        ElseIf Box2_UPb_Options.OptionButton3 = False Then
            .Range(UPbStd_ColumnConcErrors12s & CellRow) = 2
        End If

        If CheckBox2 = True Then
            .Range(UPbStd_ColumnConcErrorsAbs & CellRow) = True
        Else
            .Range(UPbStd_ColumnConcErrorsAbs & CellRow) = False
        End If
        
    End With
            
    Call UserForm_Initialize
    
End Sub

Private Sub CommandButton2_New_Click()

    MsgBox "If you want to change some information from one of the existant standards, " & _
            "select and edit it as you want an then press the save button.", vbOKOnly
    
    Call Clear_Box2UserForm
    
    Box2_UPb_Options.MultiPage1.Value = 1
    
    MsgBox "Please, fill the form with the standard information and then press Save."
    
End Sub

Sub Clear_Box2UserForm()
    
    With Box2_UPb_Options
        .ComboBox1_ExternalStd = ""
        .TextBox25 = ""
        .TextBox11_StandardDescription = ""
        .TextBox12_68Ratio = ""
        .TextBox15_68RatioError = ""
        .TextBox13_75Ratio = ""
        .TextBox16_75RatioError = ""
        .TextBox14_76Ratio = ""
        .TextBox17_76RatioError = ""
        .TextBox23_82Ratio = ""
        .TextBox24_82RatioError = ""
        .OptionButton1_1sigma = False
        .OptionButton2_2sigma = False
        CheckBox1_Abs = False
        .TextBox18Uppm = ""
        .TextBox21_UppmError = ""
        .TextBox19_Thppm = ""
        .TextBox22_ThppmError = ""
        .OptionButton3 = False
        .OptionButton4 = False
        .CheckBox2 = False
    End With

End Sub

Private Sub CommandButton6_Export_Click()
    
    Dim StandardsWB As Workbook
    Dim AlertsDisplay As Boolean
    Dim UpdtDispl As Boolean
    
    If TW Is Nothing Then
        Call PublicVariables
    End If

    Set StandardsWB = Application.Workbooks.Add
    
    UpdtDispl = Application.ScreenUpdating
        Application.ScreenUpdating = True
            StandardsUPb_TW_Sh.Copy Before:=StandardsWB.Worksheets(1)
        Application.ScreenUpdating = UpdtDispl
    
    AlertsDisplay = Application.DisplayAlerts
    Application.DisplayAlerts = False
        StandardsWB.Worksheets(2).Delete
    Application.DisplayAlerts = AlertsDisplay

End Sub

Private Sub CommandButton7_U238U235Default_Click()

    TextBox8_RatioUranium.Value = TW_RatioUranium_UPb

End Sub

Private Sub CommandButton8_mVtoCPS_Click()

    TextBox10_MvtoCPS.Value = TW_mVtoCPS_UPb

End Sub

Private Sub CommandButton9_Hg202Hg202Default_Click()

    TextBox9_NaturalRatioMercury.Value = TW_RatioMercury_UPb

End Sub

Private Sub TextBox10_MvtoCPS_Change()

    If _
        TextBox10_MvtoCPS.Value = "" Or _
        IsNumeric(TextBox10_MvtoCPS.Value) = False Or _
        TextBox10_MvtoCPS.Value <= 0 Then
            
            MsgBox "The constant must be a number bigger than 0."
                TextBox10_MvtoCPS.Value = CurrentmVCPS
                
    End If
    
End Sub

Private Sub TextBox8_RatioUranium_Change()

    If _
        TextBox8_RatioUranium.Value = "" Or _
        IsNumeric(TextBox8_RatioUranium.Value) = False Or _
        TextBox8_RatioUranium.Value <= 0 Then
            
            MsgBox "The constant must be a number bigger than 0."
                TextBox8_RatioUranium.Value = Current238U235U
                
    End If

End Sub

Private Sub TextBox9_NaturalRatioMercury_Change()
        
    If _
        TextBox9_NaturalRatioMercury.Value = "" Or _
        IsNumeric(TextBox9_NaturalRatioMercury.Value) = False Or _
        TextBox9_NaturalRatioMercury.Value <= 0 Then
            
            MsgBox "The constant must be a number bigger than 0."
                TextBox9_NaturalRatioMercury = Current202Hg204Hg
    
    End If

End Sub

Private Sub TextBox8_RatioUranium_Enter()
    
    Current238U235U = TextBox8_RatioUranium.Value
    
End Sub

Private Sub TextBox9_NaturalRatioMercury_Enter()

    Current202Hg204Hg = TextBox9_NaturalRatioMercury

End Sub

Private Sub TextBox10_MvtoCPS_Enter()

    CurrentmVCPS = TextBox10_MvtoCPS.Value

End Sub

Private Sub UserForm_Initialize()
    
    Dim counter As Integer 'Used to add itens to External Standard ComboBox
    Dim a As Integer
    Dim Ctrls
    Dim StdNameCombo As String
    
    If TW Is Nothing Then
        Call PublicVariables
    End If

    Set ErrBlank = Box2_UPb_Options.CheckBox3_BlankErrors
    Set ErrExtStd = Box2_UPb_Options.CheckBox4_ExtStdErrors
    Set ExtStdRepro = Box2_UPb_Options.CheckBox5_ExtStdRepro
    Set ErrExtStdCert = Box2_UPb_Options.CheckBox6_CertExtStd
    
'    ErrBlank = False 'UPDATE
'    ErrExtStd = True 'UPDATE

    Call Load_UPbStandardsTypeList
    
    Call StandardsUPbComboBox
        
'    TextBox8_RatioUranium = TW_RatioUranium_UPb
'    TextBox9_NaturalRatioMercury = TW_RatioMercury_UPb
'    TextBox10_MvtoCPS = TW_mVtoCPS_UPb
    
    Set ChoosenStandard = Box2_UPb_Options.ComboBox1_ExternalStd
    StdNameCombo = ChoosenStandard
                  
    Call Clear_Box2UserForm
    
    'The if structure below just affects the External Standard page.
    For a = 1 To UBound(UPbStd)
        If UPbStd(a).StandardName = StdNameCombo Then
            Box2_UPb_Options.ComboBox1_ExternalStd = StdNameCombo
            Box2_UPb_Options.TextBox25 = UPbStd(a).Mineral
            Box2_UPb_Options.TextBox11_StandardDescription = UPbStd(a).Description
            Box2_UPb_Options.TextBox12_68Ratio = Val(UPbStd(a).Ratio68)
            Box2_UPb_Options.TextBox15_68RatioError = Val(UPbStd(a).Ratio68Error)
            Box2_UPb_Options.TextBox13_75Ratio = Val(UPbStd(a).Ratio75)
            Box2_UPb_Options.TextBox16_75RatioError = Val(UPbStd(a).Ratio75Error)
            Box2_UPb_Options.TextBox14_76Ratio = Val(UPbStd(a).Ratio76)
            Box2_UPb_Options.TextBox17_76RatioError = Val(UPbStd(a).Ratio76Error)
            Box2_UPb_Options.TextBox23_82Ratio = Val(UPbStd(a).Ratio82)
            Box2_UPb_Options.TextBox24_82RatioError = Val(UPbStd(a).Ratio82Error)
            
            If UPbStd(a).RatioErrors12s = 1 Then
                Box2_UPb_Options.OptionButton1_1sigma = True
            ElseIf UPbStd(a).RatioErrors12s = 2 Then
                Box2_UPb_Options.OptionButton2_2sigma = True
            End If
            
            If UPbStd(a).RatioErrorsAbs = True Then
                CheckBox1_Abs = True
            End If
            
            If UPbStd(a).ConcErrorsAbs = True Then
                CheckBox2 = True
            End If
            
            Box2_UPb_Options.TextBox18Uppm = Val(UPbStd(a).UraniumConc)
            Box2_UPb_Options.TextBox21_UppmError = Val(UPbStd(a).UraniumConcError)
            Box2_UPb_Options.TextBox19_Thppm = Val(UPbStd(a).ThoriumConc)
            Box2_UPb_Options.TextBox22_ThppmError = Val(UPbStd(a).ThoriumConcError)
            
            If UPbStd(a).ConcErrors12s = 1 Then
                Box2_UPb_Options.OptionButton3 = True
            ElseIf UPbStd(a).ConcErrors12s = 2 Then
                Box2_UPb_Options.OptionButton4 = True
            End If
                
            If UPbStd(a).ConcErrorsAbs = True Then
                CheckBox2 = True
            End If

            a = UBound(UPbStd)
        End If
    Next
                   
End Sub

Private Sub CommandButton3_Ok_Click()

    'Updated 24-08-2015

    If TW Is Nothing Then
        Call PublicVariables
    End If

    'Page constants
    
        If TW_RatioMercury_UPb <> Val(TextBox9_NaturalRatioMercury.Value) Then
            If MsgBox("The 202Hg/204Hg constant is different than the standard (" & TW_RatioMercury_UPb & _
                "). Would you like to set it as the new standard?", vbYesNo) = vbYes Then
                    
                TW_RatioMercury_UPb = Val(TextBox9_NaturalRatioMercury.Value)
            End If
        End If
        
        If TW_RatioUranium_UPb <> Val(TextBox8_RatioUranium.Value) Then
            If MsgBox("The 238U/235U constant is different than the standard (" & TW_RatioUranium_UPb & _
                "). Would you like to set it as the new standard?", vbYesNo) = vbYes Then
                    
                TW_RatioUranium_UPb = Val(TextBox8_RatioUranium.Value)
            End If
        End If
        
        If TW_mVtoCPS_UPb <> Val(TextBox10_MvtoCPS.Value) Then
            If MsgBox("The mV to CPS constant is different than the standard (" & TW_mVtoCPS_UPb & _
                "). Would you like to set it as the new standard?", vbYesNo) = vbYes Then
                    
                TW_mVtoCPS_UPb = Val(TextBox10_MvtoCPS.Value)
            End If
        End If
        
        'The choosen constants are stored in Start-AND-Options sheet
        RatioMercury_UPb = Val(TextBox9_NaturalRatioMercury.Value)
        RatioUranium_UPb = Val(TextBox8_RatioUranium.Value)
        mVtoCPS_UPb = Val(TextBox10_MvtoCPS.Value)

        'Page error propagation
        ErrBlank_UPb = ErrBlank
        ErrExtStd_UPb = ErrExtStd
        ExtStdRepro_UPb = ExtStdRepro
        ErrExtStdCert_UPb = ErrExtStdCert
    
    Box2_UPb_Options.Hide
            
End Sub

Private Sub CommandButton5_Delete_Click()

    Dim StdFoundTW As Boolean
    Dim counter As Integer
    Dim UPbNameRng As Range
    Dim CellRow As Integer
    
    If MsgBox("Are you sure you would like to delete this standard?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Set ChoosenStandard = Box2_UPb_Options.ComboBox1_ExternalStd
    
    If IsEmpty(ChoosenStandard) = True Then
        MsgBox "The name of the standard is empty."
            Exit Sub
    End If
    
    StdFoundTW = False
    
    For Each UPbNameRng In UPbStd_StandardsNames
        
        CellRow = UPbNameRng.Row
        
        If ChoosenStandard = UPbNameRng Then
            StandardsUPb_TW_Sh.Rows(CellRow).Delete Shift:=xlUp
                StdFoundTW = True
        End If
            
    Next

    If StdFoundTW = False Then
        MsgBox "There is no standard with this name."
            Exit Sub
    End If
    
    Call UserForm_Initialize

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim Response As Integer
    
    If CloseMode = vbFormControlMenu Then
        Box2_UPb_Options.Hide
    End If
    
End Sub
