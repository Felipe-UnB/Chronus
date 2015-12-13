VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box1_Start 
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5625
   OleObjectBlob   =   "Box1_Start.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Box1_Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'from http://stackoverflow.com/questions/11654788/vba-userform-running-twice-when-changing-caption

'Even better, don't use auto-instantiating variables, convenient though they are (and don't use the New keyword
'in a Dim statement either). You can control when your variables are created and destroyed and it's a best
'practice. Something like this in a standard module

'Sub uftst()
'
'    Dim uf As UserForm1
'
'    Set uf = New UserForm1 'you control instantiation here
'
'    'Now you can change properties before you show it
'    uf.Caption = "blech"
'    uf.Show
'
'    Set uf = Nothing 'overkill, but you control destruction here
'
'End Sub
'Note that if the ShowModal property is set to False that the code will continue to execute, so don't
'destroy the variable if running modeless.

Private Sub CommandButton5_Preferences_Click()

    Box2_UPb_Options.Show
    
End Sub

Private Sub CommandButton6_Addresses_Click()

    Call SetAddressess

End Sub

Private Sub CommandButton7_Click()

    If _
        Len(TextBox10_ExternalStandardName.Value) >= 2 And _
        CompareAnalysisNames(TextBox10_ExternalStandardName) <> "ERROR" And _
        Len(TextBox8_BlankName.Value) >= 2 And _
        CompareAnalysisNames(TextBox8_BlankName) <> "ERROR" And _
        Len(TextBox9_SamplesNames.Value) >= 2 And _
        CompareAnalysisNames(TextBox9_SamplesNames) <> "ERROR" _
    Then
        
        TW_BlankName.Value = Box1_Start.TextBox8_BlankName.Value
        TW_SampleName.Value = Box1_Start.TextBox9_SamplesNames.Value
        TW_PrimaryStandardName.Value = Box1_Start.TextBox10_ExternalStandardName
        
        MsgBox "Names set!", vbOKOnly
        
    Else
    
        TW.Save
        
        MsgBox "Names not set. Please, check each field.", vbOKOnly
        
    End If

End Sub

Private Sub TextBox10_ExternalStandardName_Change()

    If Len(TextBox10_ExternalStandardName.Value) >= 2 Then
        If CompareAnalysisNames(TextBox10_ExternalStandardName) = "ERROR" Then
            MsgBox "Please, check the names of the analyses types. The name of type must not contain the name of other."
                TextBox10_ExternalStandardName.SetFocus
        End If
    End If

End Sub

Private Sub TextBox5_InternalStandardName_Change()

    If Len(TextBox5_InternalStandardName.Value) >= 2 Then
        If CompareAnalysisNames(TextBox5_InternalStandardName) = "ERROR" Then
            MsgBox "Please, check the names of the analyses types. The name of type must not contain the name of other."
                TextBox5_InternalStandardName.SetFocus
        End If
    End If

End Sub

Private Sub TextBox8_BlankName_Change()

    If Len(TextBox8_BlankName.Value) >= 2 Then
        If CompareAnalysisNames(TextBox8_BlankName) = "ERROR" Then
            MsgBox "Please, check the names of the analyses types. The name of type must not contain the name of other."
                TextBox8_BlankName.SetFocus
        End If
    End If

End Sub

Private Sub TextBox9_SamplesNames_Change()

    If Len(TextBox9_SamplesNames.Value) >= 2 Then
        If CompareAnalysisNames(TextBox9_SamplesNames) = "ERROR" Then
            MsgBox "Please, check the names of the analyses types. The name of type must not contain the name of other."
                TextBox9_SamplesNames.SetFocus
        End If
    End If

End Sub

Sub SmallNameMessage()
    
    'Added 27/08/2015
    
    'Message that appears if the user writes a non valid for the analyses

    MsgBox "The name of samples, standards and blank analyses must have at least two characters. " & _
        "These names can not be repated in more than one type of analysis.", vbOKOnly

End Sub
Public Sub UserForm_Initialize()
        
    Dim ProblemMsgBox As Variant 'Message box displayed if something seems to be wrong in StandardsUPb sheet.
    Dim c As Variant 'Used to add itens to External Standard ComboBox
    Dim P As Variant

    If mwbk Is Nothing Then
        Call PublicVariables
    End If
    
'    Load Box2_UPb_Options
'    Load Box4_Addresses
    
    'Code to assign values from Box1_Start to the related variables
    Set SampleName = Me.TextBox2
    Set ReductionDate = Me.TextBox4
    Set ReducedBy = Me.TextBox3
    Set FolderPath = Me.TextBox6
    Set ExternalStandard = Me.ComboBox1_ExternalStd
    Set InternalStandardCheck = Me.CheckBox1_InternalStandard
    Set InternalStandardName = Me.TextBox5_InternalStandardName
    Set Spot = Me.OptionButton3_Spot
    Set Raster = Me.OptionButton4_Raster
    Set Detector206MIC = Me.OptionButton1_206MIC
    Set Detector206Faraday = Me.OptionButton2_206Faraday
    Set CheckData = Me.CheckBox2_CheckRawData
    Set BlankName = Me.TextBox8_BlankName
    Set SamplesNames = Me.TextBox9_SamplesNames
    Set ExternalStandardName = Me.TextBox10_ExternalStandardName
    Set SecondaryStandardName = Me.TextBox5_InternalStandardName
    Set RawNumberCycles = Me.TextBox11_HowMany
    Set CycleDuration = Me.TextBox12_CycleDuration
            
    CheckData = True
            
    Call CheckFundamentalParameters
    
    If IIM.count <> 0 Then

        ProblemMsgBox = MsgBox("Are you reducing this data for the first time?", vbYesNo)

            If ProblemMsgBox = vbNo Then
                For Each P In IIM
                    Box3_ProblemsList.ListBox1_ProblemsList.AddItem P
                Next
                    Box3_ProblemsList.Show
                        Call Load_UPbStandardsTypeList
                            Call StandardsUPbComboBox
                                Call PreviousValues
            Else
                Call SelectFolder
                    Call Load_UPbStandardsTypeList
                        Call DefaultValues
                        'Box2_UPb_Options.MultiPage1.Value = 0
                            'Box2_UPb_Options.Show
                                'Call SetAddressess
            End If

    Else

        Call PreviousValues
            Call Load_UPbStandardsTypeList
                Call StandardsUPbComboBox

    End If
        
End Sub

Private Sub CommandButton3_Ok_Click()
    'Code to assign values from Box2_UPb_Options to the related variables
    'UPDATED 29082015 - Code to check analyses name added. Multiple name can be assigned to each type of analysis, but they must have len>2
                        'and one name cant be present in other.
    'UPDATED 02102015 - If the user type a analysis name with lenght smaller than 2, there is an warning message.
    
    Dim MsgAlert As String
    Dim MsgBoxAlert As Variant 'Message box for for many checks done below
    Dim c As Variant 'Variable used in a for each structure
    Dim AddressRawDataFile As Variant 'Array of variables with address in Box2_UPb_Options
    Dim Counter As Integer
    Dim StdName As Integer
    
    'The conditional clauses below are necessary because not all isotopes must have been analyzed
    If Isotope208analyzed = True And Isotope232analyzed = True Then
        AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, RawPb207Range, RawPb208Range, RawTh232Range, RawU238Range, _
        RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, RawPb207HeaderRange, RawPb208HeaderRange, RawTh232HeaderRange, _
        RawU238HeaderRange)
    ElseIf Isotope208analyzed = True And Isotope232analyzed = False Then
        AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, RawPb207Range, RawPb208Range, RawU238Range, _
        RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, RawPb207HeaderRange, RawPb208HeaderRange, _
        RawU238HeaderRange)
    ElseIf Isotope208analyzed = False And Isotope232analyzed = True Then
        AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, RawPb207Range, RawTh232Range, RawU238Range, _
        RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, RawPb207HeaderRange, RawTh232HeaderRange, _
        RawU238HeaderRange)
    ElseIf Isotope208analyzed = False And Isotope232analyzed = False Then
        AddressRawDataFile = Array(RawHg202Range, RawPb204Range, RawPb206Range, RawPb207Range, RawU238Range, _
        RawHg202HeaderRange, RawPb204HeaderRange, RawPb206HeaderRange, RawPb207HeaderRange, _
        RawU238HeaderRange)
    End If
    
    'All of the above variables must not be = ""
    For Each c In AddressRawDataFile
        'on error resume Next
        If c.Value = "" Then
            MsgBoxAlert = MsgBox("There are one or more addresses missing in Start-AND-Options sheet. " & _
            "Please, check it.", vbOKOnly, "Missing Address")
                Call SetAddressess
        End If
    Next
    
    'All the addresses
        
    'If the user checked Internal Standard, he/she must write its name!
    If InternalStandardCheck = True And Len(TextBox5_InternalStandardName.Value) < 2 Then
        MsgBox "What is the name of the internal standard analyzed? It must be 2 characters long or bigger."
            TextBox5_InternalStandardName.SetFocus
                Exit Sub
    End If
            
    'The user must choose an external standard.
    If ExternalStandard = "" Then
        MsgBox "You must choose an external standard!"
            ComboBox1_ExternalStd.SetFocus
                Exit Sub
    End If
            
    'Spot or raster and MIC or Faraday, one option in each pair, must be choosen.
    If Spot = False And Raster = False Then
        MsgBox "You must choose Spot or Raster option!"
            OptionButton3_Spot.SetFocus
                Exit Sub
    End If
    
    If Detector206MIC = False And Detector206Faraday = False Then
        MsgBox "You must choose MIC or Faraday cup for 206 isotope!"
            OptionButton1_206MIC.SetFocus
                Exit Sub
    End If
    
    'The user must indicate the names for blanks, samples and standards
    
    MsgAlert = "You must indicate the names for "
    
            If Len(BlankName) < 2 Then
                MsgBoxAlert = MsgBox(MsgAlert & "blanks! It must be 2 characters long or bigger.", vbOKOnly, "Blank names")
                    TextBox8_BlankName.SetFocus
                        Exit Sub
                ElseIf Len(SamplesNames) < 2 Then
                    MsgBoxAlert = MsgBox(MsgAlert & "samples! It must be 2 characters long or bigger.", vbOKOnly, "Samples names")
                        TextBox9_SamplesNames.SetFocus
                            Exit Sub
                        
                    ElseIf Len(ExternalStandardName) < 2 Then
                        MsgBoxAlert = MsgBox(MsgAlert & "primary standard analyses! It must be 2 characters long or bigger.", vbOKOnly, "Standard analyses names")
                            TextBox10_ExternalStandardName.SetFocus
                                Exit Sub

                        ElseIf CheckBox1_InternalStandard = True And Len(SecondaryStandardName) < 2 Then
                            MsgBoxAlert = MsgBox(MsgAlert & "secondary standard analyses! It must be 2 characters long or bigger.", vbOKOnly, "Standard analyses names")
                                TextBox10_ExternalStandardName.SetFocus
                                    Exit Sub
            End If
            
    If Len(TextBox10_ExternalStandardName.Value) < 2 Then
        Call SmallNameMessage
            TextBox10_ExternalStandardName.SetFocus
                Exit Sub
    End If
    
    If CheckBox1_InternalStandard = True And Len(TextBox5_InternalStandardName.Value) < 2 Then
        Call SmallNameMessage
            TextBox5_InternalStandardName.SetFocus
                Exit Sub
    End If
    
    If Len(TextBox8_BlankName.Value) < 2 Then
        Call SmallNameMessage
            TextBox8_BlankName.SetFocus
                Exit Sub
    End If
    
    If Len(TextBox9_SamplesNames.Value) < 2 Then
        Call SmallNameMessage
            TextBox9_SamplesNames.SetFocus
                Exit Sub
    End If

    'The following 4 if-then blocks were added to check, when the user press ok, all analyses names.
    If CompareAnalysisNames(BlankName) = "ERROR" Then
        Call SmallNameMessage
            TextBox8_BlankName.SetFocus
                Exit Sub
    End If

    If CompareAnalysisNames(SamplesNames) = "ERROR" Then
        Call SmallNameMessage
            SamplesNames.SetFocus
                Exit Sub
    End If
    
    If CompareAnalysisNames(ExternalStandardName) = "ERROR" Then
        Call SmallNameMessage
            ExternalStandardName.SetFocus
                Exit Sub
    End If
           
    If CheckBox1_InternalStandard = True And CompareAnalysisNames(SecondaryStandardName) = False Then
        Call SmallNameMessage
            SecondaryStandardName.SetFocus
                Exit Sub
    End If
    
    '-----------------------------------------------------------------------------------------------------------------

    If RawNumberCycles = "" Then
        MsgBox "You must write the number of cycles per analyis."
            Box1_Start.TextBox11_HowMany.SetFocus
                Exit Sub
    End If
    
    If CycleDuration = "" Then
        MsgBox "You must indicate the cycle duration." & vbNewLine & _
        "Please, be careful, it must be inserted as ss.ms (00 to 59 . 000 to 999)."
            TextBox12_CycleDuration.SetFocus
                Exit Sub
    End If
    
'    Copying values to Workbook UPb
    
    SampleName_UPb = SampleName
    ReductionDate_UPb = ReductionDate
    ReducedBy_UPb = ReducedBy
    FolderPath_UPb = FolderPath
    ExternalStandard_UPb = ExternalStandard
    InternalStandardCheck_UPb = InternalStandardCheck
    RawNumberCycles_UPb = RawNumberCycles
    CycleDuration_UPb = CycleDuration
    
    If InternalStandardCheck = True Then
        InternalStandardCheck_UPb = InternalStandardCheck
        InternalStandard_UPb = InternalStandardName
        Else
        InternalStandard_UPb = "False"
    End If

    If Spot = True Then
            SpotRaster_UPb = "Spot"
        Else
            SpotRaster_UPb = "Raster"
    End If

    If Detector206MIC = True Then
            Detector206_UPb = "MIC"
        Else
            Detector206_UPb = "Faraday Cup"
    End If

        CheckData_UPb = CheckData
        
        BlankName_UPb = BlankName
        
        SamplesNames_UPb = SamplesNames
        
        ExternalStandardName_UPb = ExternalStandardName
            
    'The 6 lines below are necessary to identify the number of the external standard in UpbStd
    StdName = 0
        
        For Counter = LBound(UPbStd) To UBound(UPbStd)
           If UPbStd(Counter).StandardName = ExternalStandard_UPb Then
               StdName = Counter
                   Counter = UBound(UPbStd)
           End If
        Next
        
        If StdName = 0 Then
            MsgBox "Please, check the name of the external standard."
                Box1_Start.ComboBox1_ExternalStd.SetFocus
                    Exit Sub
        End If
        
        StandardName_UPb = UPbStd(StdName).StandardName
        Mineral_UPb = UPbStd(StdName).Mineral
        Description_UPb = UPbStd(StdName).Description
        Ratio68_UPb = UPbStd(StdName).Ratio68
        Ratio68Error_UPb = UPbStd(StdName).Ratio68Error
        Ratio75_UPb = UPbStd(StdName).Ratio75
        Ratio75Error_UPb = UPbStd(StdName).Ratio75Error
        Ratio76_UPb = UPbStd(StdName).Ratio76
        Ratio76Error_UPb = UPbStd(StdName).Ratio76Error
        RatioErrors12s_UPb = UPbStd(StdName).RatioErrors12s
        
        If UPbStd(StdName).RatioErrorsAbs = True Then
            RatioErrorsAbs_UPb = "abs"
        Else
            RatioErrorsAbs_UPb = "%"
        End If
        
        UraniumConc_UPb = UPbStd(StdName).UraniumConc
        UraniumConcError_UPb = UPbStd(StdName).UraniumConcError
        ThoriumConc_UPb = UPbStd(StdName).ThoriumConc
        ThoriumConcError_UPb = UPbStd(StdName).ThoriumConcError
        ConcErrors12s_UPb = UPbStd(StdName).ConcErrors12s
        
        If UPbStd(StdName).ConcErrorsAbs = True Then
            ConcErrorsAbs_UPb = "abs"
        Else
            ConcErrorsAbs_UPb = "%"
        End If
        
        'Storing constants ans options from the preferences menu
                'The choosen constants are stored in Start-AND-Options sheet
        RatioMercury_UPb = Box2_UPb_Options.TextBox9_NaturalRatioMercury.Value
        RatioUranium_UPb = Box2_UPb_Options.TextBox8_RatioUranium.Value
        mVtoCPS_UPb = Box2_UPb_Options.TextBox10_MvtoCPS.Value

        'Page error propagation
        ErrBlank_UPb = Box2_UPb_Options.CheckBox3_BlankErrors
        ErrExtStd_UPb = Box2_UPb_Options.CheckBox4_ExtStdErrors
        ExtStdRepro_UPb = Box2_UPb_Options.CheckBox5_ExtStdRepro
        ErrExtStdCert_UPb = Box2_UPb_Options.CheckBox6_CertExtStd
            
    Box1_Start.Hide
        Box2_UPb_Options.Hide
            Box3_ProblemsList.Hide
                Box4_Addresses.Hide
    
    Call FormatMainSh
    
    ScreenUpd = Application.ScreenUpdating
                            
        If ScreenUpd = False Then Application.DisplayAlerts = True
        
        If MsgBox("Would you like to start the reduction process?", vbYesNo) = vbYes Then
            Application.ScreenUpdating = ScreenUpd
                Call FullDataReduction
        Else
            Application.ScreenUpdating = ScreenUpd
                Call UnloadAll
                    Application.GoTo SamList_Sh.Range("A1")
        End If

    mwbk.Save
        Application.RecentFiles.Add (mwbk.Name)

End Sub

Private Sub CheckBox1_InternalStandard_Click()
    'If IntStdAnalysed is checked there are some actions, if not, there are different actions.
    
    If CheckBox1_InternalStandard.Value = True Then 'Checked
        
            TextBox5_InternalStandardName.Enabled = True 'Enables the textbox.
            TextBox5_InternalStandardName.Value = "" 'Cleans the textbox.
            TextBox5_InternalStandardName.BackColor = vbBlack 'Change the textbox color to white.
            TextBox5_InternalStandardName.ForeColor = vbWhite
            TextBox5_InternalStandardName.SetFocus
            TextBox5_InternalStandardName.Value = InternalStandard_UPb.Value
            
        Else
        
            TextBox5_InternalStandardName.Enabled = False 'Not checked
            TextBox5_InternalStandardName.Value = "" 'Cleans the textbox.
            TextBox5_InternalStandardName.BackColor = &H8000000F 'Changes the textbox color to grey.
    
    End If
End Sub

Private Sub TextBox6_Enter()
    
    Call SelectFolder
    TextBox6.Value = FolderPath
        
End Sub

'Private Sub TextBox6_Change()
'    'This is called when the path of the folder with the data filesare changed, so the name of the
'    'sample is defined based again on the name of the folder
'
'    Dim Question As Integer
'
'    Question = MsgBox("Would you like to rename this sample based on the name of the folder where it is?", vbYesNo)
'
'    on error resume Next
'    If Question = 6 Then
'        SampleName = Dir(FolderPath, vbDirectory)
'            Box1_Start.TextBox2.Value = SampleName 'Inserts a name for the sample based on the name of the folder where it is stored
'
'    End If
    
'End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim Response As Integer
    
    If CloseMode = vbFormControlMenu Then
                
        Response = MsgBox("Do you really want to end the program execution?", vbYesNo)
            If Response = vbNo Then
                Cancel = True
            ElseIf Response = vbYes Then
                Call UnloadAll
                    End
            End If
            
    End If
    
End Sub
