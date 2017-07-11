Attribute VB_Name = "Calculations"
Option Explicit

Sub ConvertAbsolute()
    
    Dim AppUpdt As Boolean
    
    AppUpdt = Application.ScreenUpdating
        
        Application.ScreenUpdating = False
    
        Call ConvertUncertantiesTo("Absolute", BlkCalc_Sh)
        Call ConvertUncertantiesTo("Absolute", SlpStdBlkCorr_Sh)
        Call ConvertUncertantiesTo("Absolute", SlpStdCorr_Sh)
        
        Call FormatBlkCalc
        Call FormatSlpStdBlkCorr
        Call FormatSlpStdCorr
        
    Application.ScreenUpdating = AppUpdt

End Sub

Sub ConvertPercentage()
    
    Application.ScreenUpdating = False
        Call ConvertUncertantiesTo("Percentage", BlkCalc_Sh)
        Call ConvertUncertantiesTo("Percentage", SlpStdBlkCorr_Sh)
        Call ConvertUncertantiesTo("Percentage", SlpStdCorr_Sh)
        
        Call FormatBlkCalc(False)
        Call FormatSlpStdBlkCorr(False)
        Call FormatSlpStdCorr(False)

    Application.ScreenUpdating = True

End Sub
Sub ConvertUncertantiesTo(UncertantiesType As String, Sh As Worksheet)
    'This procedure will take all uncertanties in BlkCalc, SlpStdBlkCorr and SlpStdCorr, and
    'convert to relative (%) or absolute.
    
    Dim RangeUnion As Range
'    Dim RangeUnionHeaders As Range
    Dim SearchStr As Variant
    Dim Message As Integer
    Dim a As Range
    Dim UpdtScreen As Boolean
    
    If mwbk Is Nothing Then
        Call PublicVariables
    End If
    
    'Columns of uncertainty in BlkCalc sheet
    
    Select Case Sh.Name
    
        Case BlkCalc_Sh.Name
            
            With BlkCalc_Sh
                Set RangeUnion = Application.Union( _
                .Range(BlkColumn21Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn21Std & BlkCalc_HeaderLine + 1).End(xlDown)), _
                .Range(BlkColumn41Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn41Std & BlkCalc_HeaderLine + 1).End(xlDown)), _
                .Range(BlkColumn61Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn61Std & BlkCalc_HeaderLine + 1).End(xlDown)), _
                .Range(BlkColumn71Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn71Std & BlkCalc_HeaderLine + 1).End(xlDown)), _
                .Range(BlkColumn81Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn81Std & BlkCalc_HeaderLine + 1).End(xlDown)), _
                .Range(BlkColumn321Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn321Std & BlkCalc_HeaderLine + 1).End(xlDown)), _
                .Range(BlkColumn381Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn381Std & BlkCalc_HeaderLine + 1).End(xlDown)), _
                .Range(BlkColumn4Comm1Std & BlkCalc_HeaderLine + 1, .Range(BlkColumn4Comm1Std & BlkCalc_HeaderLine + 1).End(xlDown)))
                
'                Set RangeUnionHeaders = Application.Union( _
'                .Range(BlkColumn21Std & BlkCalc_HeaderLine), _
'                .Range(BlkColumn41Std & BlkCalc_HeaderLine), _
'                .Range(BlkColumn61Std & BlkCalc_HeaderLine), _
'                .Range(BlkColumn71Std & BlkCalc_HeaderLine), _
'                .Range(BlkColumn81Std & BlkCalc_HeaderLine), _
'                .Range(BlkColumn321Std & BlkCalc_HeaderLine), _
'                .Range(BlkColumn381Std & BlkCalc_HeaderLine), _
'                .Range(BlkColumn4Comm1Std & BlkCalc_HeaderLine))
'
                SearchStr = InStr(.Range(BlkColumn21Std & BlkCalc_HeaderLine), "%")
                
            End With
            
        Case SlpStdBlkCorr_Sh.Name
    
            With SlpStdBlkCorr_Sh
                Set RangeUnion = Application.Union( _
                .Range(Column681Std & HeaderRow + 1, .Range(Column681Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column761Std & HeaderRow + 1, .Range(Column761Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column751Std & HeaderRow + 1, .Range(Column751Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column21Std & HeaderRow + 1, .Range(Column21Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column41Std & HeaderRow + 1, .Range(Column41Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column61Std & HeaderRow + 1, .Range(Column61Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column71Std & HeaderRow + 1, .Range(Column71Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column81Std & HeaderRow + 1, .Range(Column81Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column321Std & HeaderRow + 1, .Range(Column321Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column381Std & HeaderRow + 1, .Range(Column381Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column641Std & HeaderRow + 1, .Range(Column641Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column741Std & HeaderRow + 1, .Range(Column741Std & HeaderRow + 1).End(xlDown)), _
                .Range(Column281Std & HeaderRow + 1, .Range(Column281Std & HeaderRow + 1).End(xlDown)))
                
'                Set RangeUnionHeaders = Application.Union( _
'                .Range(Column681Std & HeaderRow), _
'                .Range(Column761Std & HeaderRow), _
'                .Range(Column751Std & HeaderRow), _
'                .Range(Column21Std & HeaderRow), _
'                .Range(Column41Std & HeaderRow), _
'                .Range(Column61Std & HeaderRow), _
'                .Range(Column71Std & HeaderRow), _
'                .Range(Column81Std & HeaderRow), _
'                .Range(Column321Std & HeaderRow), _
'                .Range(Column381Std & HeaderRow), _
'                .Range(Column641Std & HeaderRow), _
'                .Range(Column741Std & HeaderRow), _
'                .Range(Column281Std & HeaderRow))
'
                SearchStr = InStr(.Range(Column681Std & HeaderRow), "%")
                
            End With
            
        Case SlpStdCorr_Sh.Name
            With SlpStdCorr_Sh
                Set RangeUnion = Application.Union( _
                .Range(StdCorr_Column681Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column681Std & StdCorr_HeaderRow + 1).End(xlDown)), _
                .Range(StdCorr_Column761Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column761Std & StdCorr_HeaderRow + 1).End(xlDown)), _
                .Range(StdCorr_Column751Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column751Std & StdCorr_HeaderRow + 1).End(xlDown)), _
                .Range(StdCorr_Column21Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column21Std & StdCorr_HeaderRow + 1).End(xlDown)), _
                .Range(StdCorr_Column41Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column41Std & StdCorr_HeaderRow + 1).End(xlDown)), _
                .Range(StdCorr_Column641Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column641Std & StdCorr_HeaderRow + 1).End(xlDown)), _
                .Range(StdCorr_Column741Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column741Std & StdCorr_HeaderRow + 1).End(xlDown)), _
                .Range(StdCorr_Column281Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column281Std & StdCorr_HeaderRow + 1).End(xlDown)))
                
'                Set RangeUnionHeaders = Application.Union( _
'                .Range(StdCorr_Column681Std & StdCorr_HeaderRow), _
'                .Range(StdCorr_Column761Std & StdCorr_HeaderRow), _
'                .Range(StdCorr_Column751Std & StdCorr_HeaderRow), _
'                .Range(StdCorr_Column21Std & StdCorr_HeaderRow), _
'                .Range(StdCorr_Column41Std & StdCorr_HeaderRow), _
'                .Range(StdCorr_Column641Std & StdCorr_HeaderRow), _
'                .Range(StdCorr_Column741Std & StdCorr_HeaderRow), _
'                .Range(StdCorr_Column281Std & StdCorr_HeaderRow))
'
                SearchStr = InStr(.Range(StdCorr_Column681Std & HeaderRow), "%")

            End With

    
    End Select
    
    On Error GoTo ErrHandler
    
    UpdtScreen = Application.ScreenUpdating
    
    Select Case UncertantiesType
            
        Case "Percentage" 'Convert to relative (percentage)
            If SearchStr <> 0 Then
                
                Application.ScreenUpdating = True: Application.GoTo Sh.Range("A1"): Application.ScreenUpdating = UpdtScreen
                
                    SearchStr = MsgBox("Data errors are absolute? Please, take a look at the table behind " & _
                    "this message, otherwise you might have to reduce all your data again.", vbYesNo, "Relative or absolute errors")
                        
                        If SearchStr = 6 Then
                            
                            If MsgBox("Are you sure?", vbYesNo) = vbYes Then
                            
                                For Each a In RangeUnion
                                    If WorksheetFunction.IsNumber(a) Then
                                        a = 100 * (a / a.Offset(, -1))
                                    End If
                                Next
                                            
                            End If
'                            For Each a In RangeUnionHeaders
'                                a = Replace(a.Value, "(abs)", "(%)")
'                            Next

                        End If
                        
            Else
            
                For Each a In RangeUnion
                    If WorksheetFunction.IsNumber(a) Then
                        a = 100 * (a / a.Offset(, -1))
                    End If
                Next
    
'                For Each a In RangeUnionHeaders
'                    a = Replace(a.Value, "(abs)", "(%)")
'                Next

            End If
        
        Case "Absolute"
            If SearchStr = 0 Then
                
                Application.ScreenUpdating = True: Application.GoTo Sh.Range("A1"): Application.ScreenUpdating = UpdtScreen
                
                    SearchStr = MsgBox("Data errors are in percentage? Please, take a look at the table behind " & _
                    "this message, otherwise you might have to reduce all your data again.", vbYesNo, "Relative or absolute errors")

                        If SearchStr = 6 Then

                            If MsgBox("Are you sure?", vbYesNo) = vbYes Then
                                
                                For Each a In RangeUnion
                                    If WorksheetFunction.IsNumber(a) Then
                                        a = a * (a.Offset(, -1) / 100)
                                    End If
                                Next
                                    
                        End If
'                            For Each a In RangeUnionHeaders
'                                a = Replace(a.Value, "(%)", "(abs)")
'                            Next

                        End If
                        
            Else
            
                For Each a In RangeUnion
                    If WorksheetFunction.IsNumber(a) = True Then
                        a = a * (a.Offset(, -1) / 100)
                    End If
                Next
    
'                For Each a In RangeUnionHeaders
'                    a = Replace(a.Value, "(%)", "(abs)")
'                Next

            End If
                        
    End Select
    
    Exit Sub
    
ErrHandler:
    MsgBox ("Uncertainties were not converted to " & UncertantiesType & "in " & Sh.Name & ".")
        End
    
End Sub

Sub CalcAllSlpStd_BlkCorr()
    'This program loads the list of external standards and samples + internal standards, and calls CalcSlpBlkCorr
    ' or CalcExtStdBlkCorr to calculate ratios and isotopes signal intensities. Different programs to reduce
    'external standards and samples + internal standards were necessary because only one blank is considered
    'for external standard (that analysed immediately before the external standard).
    
    'UNCERTANTIES ARE AUTOMATICALLY CONVERTED TO ABSOLUTE BY ConvertUncertantiesTo subprocedure at the end
    'of CalcAllSlpStdBlkCorr
    
    Dim a As Variant
    Dim C As Integer
    Dim d As Range
    Dim E As Variant
    Dim f As Variant
    Dim G As Range
    Dim H As Double
    Dim i As Double
    Dim J As Double
    Dim K As Integer 'Offset between columns in raw data file (238 to 232, 202 to 204, etc)
    
    Dim RangeUnionHeaders As Range

'    Dim Blk1 As Integer 'Row of the blank 1 in BlkCalc
'    Dim Blk2 As Integer 'Row of the blank 2 in BlkCalc
'    Dim bSlp As Workbook 'Samples and INTERNAL standards workbooks opened
'    Dim bStd1 As Workbook 'Standards workbooks opened (Std before sample)
'    Dim bStd2 As Workbook 'Standards workbooks opened (Std after sample)
    Dim SlpStd() As Integer 'Array with samples and INTERNAL standards (treated as samples) IDs only
    Dim ExtStd() As Integer 'Array with external standards IDs only
    Dim BlkAverage(1 To 18) As Double
    Dim AnalysesListNumber As Integer 'Variable that stores the identification of the sample in AnalysesList
    Dim SlpStd1Std2(1 To 3) As Workbook 'Array with the opened worksheets of sample (or internal standard) and external standards (before and after)
    
    If SlpStdBlkCorr_Sh Is Nothing Then
        Call PublicVariables
    End If

    If IsArrayEmpty(BlkFound) = True Then
        Call IdentifyFileType
    End If
    
    If IsArrayEmpty(PathsNamesIDsTimesCycles) = True Then
        Call SetPathsNamesIDsTimesCycles
    End If

    SlpStdBlkCorr_Sh.Cells.Clear
    
    Call LoadSamListMap
    Call LoadStdListMap
         
    ReDim Preserve SlpStd(1 To UBound(SlpFound) + UBound(IntStdFound) + 2) As Integer
    
    C = 2

    For Each a In SlpFound 'Samples IDs are copied to a different array (Blanks) which accepts only numbers (IDs)
        SlpStd(C - 1) = SamList_Sh.Range(a).Offset(, 1)
        C = C + 1
    Next
    
    For Each a In IntStdFound 'Internal standards IDs are copied to a different array (SlpStd) which accepts only numbers (IDs)
        SlpStd(C - 1) = SamList_Sh.Range(a).Offset(, 1)
        C = C + 1
    Next
    
    If Detector206_UPb = "Faraday Cup" Then
            H = mVtoCPS_UPb
        ElseIf Detector206_UPb = "MIC" Then
            H = 1
        Else
            MsgBox "Please, indicate if 206Pb was analyzed using Faraday cup or Ion counter."
                Application.GoTo StartANDOptions_Sh.Range("A1")
                    End
    End If
    
    'Now, each raw sample and standard (internal) data file will be opened and processed
    
    C = HeaderRow
    
    SlpStdBlkCorr_Sh.Cells.Clear
    
    For Each a In SlpStd
        C = C + 1
        
        'Below, a is the sample ID, c is a counter for the lines where data will be pasted
        'and h is the VtoCPS constant, if Detector206_UPb = "Faraday Cup".
        Call CalcSlp_BlkCorr(a, H, True, C)
                                    
    Next
    
    ''''''''''''''''''''''''''''''''''''
    'The same procedures for samples and internal standards will be now repeated to external standards
    
    ReDim ExtStd(1 To UBound(StdFound) + 1) As Integer
    
    C = 2
    
    For Each a In StdFound 'External standards IDs are copied to a different array (SlpStd) which accepts only numbers (IDs)
        ExtStd(C - 1) = SamList_Sh.Range(a).Offset(, 1)
        C = C + 1
    Next
    
    'Now, each external standard data file will be opened and processed
    
    C = SlpStdBlkCorr_Sh.Range("A" & HeaderRow + 1).End(xlDown).Row
    
    For Each a In ExtStd
        
        C = C + 1
        
        'Below, a is the sample ID, c is a counter for the line where data will be pasted
        'and h is the VtoCPS constant, if Detector206_UPb = "Faraday Cup".
        Call CalcExtStd_BlkCorr(a, H, True, C)
                                    
    Next
    
    Call ExternalReproSamples
    
Exit Sub
    
ErrHandler:
    MsgBox "Plese, check the blank raw data. Only numbers in isotopes signal range are accepted."
        ''Application.DisplayAlerts = true
            'Application.ScreenUpdating = False
                End
    
End Sub


Sub CalcSlp_BlkCorr(ByVal a As Integer, ByVal H As Double, Optional ByVal CallCommonCalc = True, _
Optional ByVal C As Integer, Optional ByVal CloseAnalysis = True)

    'About the procedure arguments, a is the sample ID, C is a counter for the lines where data will be pasted in
    'and H is the VtoCPS constant, if Detector206_UPb = "Faraday Cup", and CallCommonCalc (true or false) forces
    'the procedure to call CommonCalc or not.
    
    'This procedure opens all samples and internal standards data files, removes blank from them and
    'then calculates all the ratios (76, 68, 28, 64, 74) and average (204Pb) necessary to the sample.
    'These informations are pasted in SlpStdBlkCorr. A different program is used to do the same with
    'External standards because their blanks are only the blank analysed closest to them and not an
    'average, like happens with samples and internal standards.
    
    'Uncertanties are absolute

    Dim f As Variant
    Dim G As Range
    Dim O As Range
    Dim i As Double, i2 As Double, i3 As Double, i4 As Double, i5 As Double
    Dim J As Double
    Dim K As Integer 'Offset between columns in raw data file (238 to 232, 202 to 204, etc)
    Dim factor As Double

    Dim E As Integer
    Dim d As Range
    Dim Blk1 As Integer 'Row number in BlkCalc_Sh where of the first sample blank
    Dim Blk2 As Integer 'Row number in BlkCalc_Sh where of the second sample blank
    Dim AnalysesListNumber As Integer
    
    Dim SearchStr As Integer 'Variable used to search for (%) in headers
    Dim RangeUnionHeaders As Range 'Range with headers
    Dim A_Header As Range 'Variable used to loop through headers
    Dim NumCycles As Integer
    
'    'The code below clears the entire SlpStdBlkCorr sheet and changes the uncertanties headers units to "(abs)".
'        SearchStr = InStr(SlpStdBlkCorr_Sh.Range(StdCorr_Column681Std & HeaderRow), "(%)")
    
    If IsArrayEmpty(PathsNamesIDsTimesCycles) Then
        Call SetPathsNamesIDsTimesCycles
    End If
    
'    If SearchStr <> 0 Then
'        With SlpStdBlkCorr_Sh
'
'            Set RangeUnionHeaders = Application.Union( _
'            .Range(Column681Std & HeaderRow), _
'            .Range(Column761Std & HeaderRow), _
'            .Range(Column751Std & HeaderRow), _
'            .Range(Column21Std & HeaderRow), _
'            .Range(Column41Std & HeaderRow), _
'            .Range(Column641Std & HeaderRow), _
'            .Range(Column741Std & HeaderRow), _
'            .Range(Column281Std & HeaderRow))
'
'            For Each A_Header In RangeUnionHeaders
'                A_Header.Value = Replace(A_Header.Value, " (%)", " (abs)")
'            Next
'        End With
'    End If
    
    'The lines inside the if structure must only be executed only if the
    'CallCommonCalc is called.
    If CallCommonCalc = True Then
        With SlpStdBlkCorr_Sh
            .Range(ColumnID & C) = PathsNamesIDsTimesCycles(ID, a) 'Copies the ID of the analysis to SlpStdBlkCorr_Sh
            .Range(ColumnSlpName & C) = PathsNamesIDsTimesCycles(FileName, a) 'Copies the file name of the analysis to SlpStdBlkCorr_Sh
        End With
    End If
    
            On Error Resume Next
                Set WBSlp = Workbooks.Open(PathsNamesIDsTimesCycles(RawDataFilesPaths, a)) 'ActiveWorkbook
                    If Err.Number <> 0 Then
                        MsgBox MissingFile1 & PathsNamesIDsTimesCycles(RawDataFilesPaths, a) & MissingFile2
                            Call UpdateFilesAddresses
                                Call UnloadAll
                                    End
                    End If
            On Error GoTo 0
            
            Call ClearCycles(WBSlp, PathsNamesIDsTimesCycles(Cycles, a)) 'Any cycles that should be discarded from the sample will be now
            
            For E = 1 To UBound(AnalysesList) 'For each structure used to find the sample or internal standard inside AnalysesList
                If a = AnalysesList(E).Sample Then 'There is a problem here because the blank for standard can be changed, it's
                                                   'not necessarly the same as the sample
                    AnalysesListNumber = E 'Using this variable I am able to retrieve from AnalysesList all the IDs that I must know
                    E = UBound(AnalysesList) 'A beautiful solution to end the if structure
                End If
            Next
           
            With BlkCalc_Sh
                
                If .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown) = "" Then
                    MsgBox ("You need at least one blank to reduce your data.")
                        Application.GoTo BlkCalc_Sh.Range("A1")
                            End
                End If
                
                For Each d In .Range(BlkColumnID & BlkCalc_HeaderLine + 1, .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown))
                    
                    If AnalysesList(AnalysesListNumber).Blk1 = d Then
                        Blk1 = d.Row
                    End If

                    If AnalysesList(AnalysesListNumber).Blk2 = d Then
                        Blk2 = d.Row
                    End If

                Next
                    
            End With
            
            With WBSlp.Worksheets(1)
                
                NumCycles = update_numcycles(WBSlp)
                
                Call CyclesTime(.Range(RawCyclesTimeRange_function(NumCycles)), NumCycles)
                
                'Subtracting 202 blank average from sample (and internal standard) signal
                If Isotope202Analyzed_UPb = True Then
                    If BlanksRecordedSamples_UPb = False Then
                        i = WorksheetFunction.Average(BlkCalc_Sh.Range(BlkColumn2 & Blk1), BlkCalc_Sh.Range(BlkColumn2 & Blk2)) 'UPDATE
                    Else
                        i = BlkCalc_Sh.Range(BlkColumn2 & Blk1)
                    End If
                        
                    If i < 0 Then i = 0
                    
                    If Detector202_UPb = "Faraday Cup" Then
                        factor = mVtoCPS_UPb
                    ElseIf Detector202_UPb = "MIC" Then
                        factor = 1
                    End If
                    
                    For Each G In .Range(Raw202Range(NumCycles))
                        If Not G = "" Then
                            
                            G = G * factor - i
                                                    
                                If G <= 0 Then
                                    G = 1
                                End If
                        End If
                    Next
                End If
                
                'Subtracting 202 blank average divided by 4.35 from sample 204 (and internal standard) signal
                If Isotope204Analyzed_UPb = True Then
                
                    If Detector204_UPb = "Faraday Cup" Then
                            factor = mVtoCPS_UPb
                        ElseIf Detector204_UPb = "MIC" Then
                            factor = 1
                    End If
                
                    If BlanksRecordedSamples_UPb = False Then
                        i = WorksheetFunction.Average(BlkCalc_Sh.Range(BlkColumn2 & Blk1), BlkCalc_Sh.Range(BlkColumn2 & Blk2))
                            i2 = WorksheetFunction.Average(BlkCalc_Sh.Range(BlkColumn4Comm & Blk1), BlkCalc_Sh.Range(BlkColumn4Comm & Blk2))
                    Else
                        i = BlkCalc_Sh.Range(BlkColumn2 & Blk1)
                            i2 = BlkCalc_Sh.Range(BlkColumn4Comm & Blk1)
                    End If
                    
                    If Isotope202Analyzed_UPb = False Then 'I am not sure if this is necessary, because if the cell is empty, then its value is 0.
                        i = 0
                    End If
                    
                    If i < 0 Then i = 0
                        If i2 < 0 Then i2 = 0
    
                    For Each G In .Range(Raw204Range(NumCycles))
                            
                        If BlkSlp202 = True And Isotope202Analyzed_UPb = True Then  'Means that 202 from both sample and blank will be considered below
                            Extra202 = .Cells(G.Row, .Range(RawHg202Range).Column) 'The rawHg202RAnge here does not need to be updated because only its column it's necessary
                        ElseIf BlkSlp202 = False Or Isotope202Analyzed_UPb = False Then
                            Extra202 = 0
                        End If
    
                        If Not G = "" Then
                            G = G * factor - ((i + Extra202) / RatioMercury_UPb) - i2
                                If G <= 0 Then
                                    G = 1
                                End If
                        End If
                    Next
                    
                End If
                                                    
                'Subtracting 206 blank average from sample (and internal standard) signal and multiplying signal by mvtoCPS
                'constant, if 206 was analysed using MIC
'                If Detector206_UPb = "MIC" Then
                
                If BlanksRecordedSamples_UPb = False Then
                    i = WorksheetFunction.Average(BlkCalc_Sh.Range(BlkColumn6 & Blk1), BlkCalc_Sh.Range(BlkColumn6 & Blk2))
                Else
                    i = BlkCalc_Sh.Range(BlkColumn6 & Blk1)
                End If
                
                If i < 0 Then i = 0
                
                If Detector206_UPb = "Faraday Cup" Then
                    factor = mVtoCPS_UPb
                ElseIf Detector206_UPb = "MIC" Then
                    factor = 1
                End If
                
                For Each G In .Range(Raw206Range(NumCycles))
                    If Not G = "" Then
                        G = G * factor - i
                            If G <= 0 Then
                                G = 1
                            End If
                    End If
                Next
                                    
                'Subtracting 207 blank average from sample (and internal standard) signal
                If BlanksRecordedSamples_UPb = False Then
                    i = WorksheetFunction.Average(BlkCalc_Sh.Range(BlkColumn7 & Blk1), BlkCalc_Sh.Range(BlkColumn7 & Blk2))
                Else
                    i = BlkCalc_Sh.Range(BlkColumn7 & Blk1)
                End If
                
                    If i < 0 Then i = 0
                    
                If Detector207_UPb = "Faraday Cup" Then
                        factor = mVtoCPS_UPb
                ElseIf Detector207_UPb = "MIC" Then
                        factor = 1
                End If
                    
                For Each G In .Range(Raw207Range(NumCycles))
                    If Not G = "" Then
                        G = G * factor - i
                            If G <= 0 Then
                                G = 1
                            End If
                    End If
                Next

                'Subtracting 208 blank average from sample (and internal standard) signal
                If Isotope208Analyzed_UPb = True Then
                    
                    If BlanksRecordedSamples_UPb = False Then
                        i = WorksheetFunction.Average(BlkCalc_Sh.Range(BlkColumn8 & Blk1), BlkCalc_Sh.Range(BlkColumn8 & Blk2))
                    Else
                        i = BlkCalc_Sh.Range(BlkColumn8 & Blk1)
                    End If
                    
                        If i < 0 Then i = 0
                        
                    If Detector208_UPb = "Faraday Cup" Then
                            factor = mVtoCPS_UPb
                    ElseIf Detector208_UPb = "MIC" Then
                            factor = 1
                    End If
                    
                    For Each G In .Range(Raw208Range(NumCycles))
                        If Not G = "" Then
                            G = G * factor - i
                                If G <= 0 Then
                                    G = 0
                                End If
                        End If
                    Next
                End If
                
                'Multiplying 232 from sample (and internal standard) signal by MvtoCPS constant
                If Isotope232Analyzed_UPb = True Then
                
                    If Detector232_UPb = "Faraday Cup" Then
                            factor = mVtoCPS_UPb
                    ElseIf Detector232_UPb = "MIC" Then
                            factor = 1
                    End If
                    
                    For Each G In .Range(Raw232Range(NumCycles))
                        If Not G = "" Then
                            G = G * factor
                                If G <= 0 Then
                                    G = 1
                                End If
                        End If
                    Next
                End If
                
                'Multiplying 238 from sample (and internal standard) signal by MvtoCPS constant
                If Detector238_UPb = "Faraday Cup" Then
                        factor = mVtoCPS_UPb
                ElseIf Detector238_UPb = "MIC" Then
                        factor = 1
                End If
                    
                For Each G In .Range(Raw238Range(NumCycles))
                    If Not G = "" Then
                        G = G * factor
                            If G <= 0 Then
                                G = 1
                            End If
                    End If
                Next
                
                If CallCommonCalc = True Then
                    Call CommonCalcSlpExtStd_BlkCorr(WBSlp.Worksheets(1), C, SpotRaster_UPb.Value, Blk1, Blk2)
                End If
                                                                
            End With
            
    If CloseAnalysis = True Then
        WBSlp.Close savechanges:=False
    End If
    
End Sub

Sub CalcExtStd_BlkCorr(ByVal a As Integer, ByVal H As Double, Optional ByVal CallCommonCalc = True, _
Optional ByVal C As Integer, Optional ByVal CloseAnalysis = True)

    'About the procedure arguments, a is the sample ID, c is a counter for the lines where data will be pasted
    'and h is the VtoCPS constant, if Detector206_UPb = "Faraday Cup".
    
    'This procedure opens all samples and internal standards data files, removes blank from them and
    'then calculates all the ratios (76, 68, 28, 64, 74) and average (204Pb) necessary to the sample.
    'These informations are pasted in SlpStdBlkCorr. A different program is used to do the same with
    'External standards because their blanks are only the blank analysed closes to the them and not an
    'average, like happens with samples and internal standards.

    Dim f As Variant
    Dim G As Range
    Dim O As Range
    Dim i As Double
    Dim i2 As Double
    Dim J As Double
    Dim K As Integer 'Offset between columns in raw data file (238 to 232, 202 to 204, etc)

    Dim E As Integer
    Dim d As Range
    Dim Blk1 As Integer 'Row number in BlkCalc_Sh where of the first sample blank
    Dim Blk2 As Integer 'Row number in BlkCalc_Sh where of the second sample blank
    Dim AnalysesListNumber As Integer
    
    Dim SearchStr As Integer 'Variable used to search for (%) in headers
    Dim RangeUnionHeaders As Range 'Range with headers
    Dim A_Header As Range 'Variable used to loop through headers
    Dim NumCycles As Integer
    Dim factor As Long
    
    'The code below clears the entire SlpStdBlkCorr sheet and changes the uncertanties headers units to "(abs)".
        SearchStr = InStr(SlpStdBlkCorr_Sh.Range(StdCorr_Column681Std & HeaderRow), "(%)")
        
    If IsArrayEmpty(PathsNamesIDsTimesCycles) Then
        Call SetPathsNamesIDsTimesCycles
    End If
    
    If SearchStr <> 0 Then
        With SlpStdBlkCorr_Sh
             
            Set RangeUnionHeaders = Application.Union( _
            .Range(Column681Std & HeaderRow), _
            .Range(Column761Std & HeaderRow), _
            .Range(Column751Std & HeaderRow), _
            .Range(Column21Std & HeaderRow), _
            .Range(Column41Std & HeaderRow), _
            .Range(Column641Std & HeaderRow), _
            .Range(Column741Std & HeaderRow), _
            .Range(Column281Std & HeaderRow))
    
            For Each A_Header In RangeUnionHeaders
                A_Header.Value = Replace(A_Header.Value, " (%)", " (abs)")
            Next
        End With
    End If
          
    If CallCommonCalc = True Then
        With SlpStdBlkCorr_Sh
            .Range(ColumnID & C) = PathsNamesIDsTimesCycles(ID, a)
            .Range(ColumnSlpName & C) = PathsNamesIDsTimesCycles(FileName, a)
        End With
    End If
          
            On Error Resume Next
                Set WBSlp = Workbooks.Open(PathsNamesIDsTimesCycles(RawDataFilesPaths, a)) 'ActiveWorkbook
                    If Err.Number <> 0 Then
                        MsgBox MissingFile1 & PathsNamesIDsTimesCycles(RawDataFilesPaths, a) & MissingFile2
                            Call UpdateFilesAddresses
                                Call UnloadAll
                                    End
                    End If
            On Error GoTo 0

            Call ClearCycles(WBSlp, PathsNamesIDsTimesCycles(Cycles, a)) 'Any cycles that should be discarded from the sample will be now
            
            For E = 1 To UBound(AnalysesList_std) 'For each structure used to find the external standard inside AnalysesList_std
                If a = AnalysesList_std(E).Std Then 'There is a problem here because the blank for standard can be changed, it's
                                                   'not necessarly the same as the sample
                    AnalysesListNumber = E 'Using this variable I am able to retrieve from AnalysesList all the IDs that I must know
                    E = UBound(AnalysesList_std) 'A beautiful solution to end the if structure
                End If
            Next
                            
            With BlkCalc_Sh
                
                If .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown) = "" Then
                    MsgBox ("You need at least two blanks to reduce your data.")
                        Application.GoTo BlkCalc_Sh.Range(BlkColumnID & BlkCalc_HeaderLine)
                            End
                End If
                
                For Each d In .Range(BlkColumnID & BlkCalc_HeaderLine + 1, .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown))
                    
                    If AnalysesList_std(AnalysesListNumber).Blk1 = d Then
                        Blk1 = d.Row
                    End If

                Next
            
            End With
                'Debug.Assert a <> 2
                With WBSlp.Worksheets(1)
                                   
                    NumCycles = update_numcycles(WBSlp)
                    
                    Call CyclesTime(.Range(RawCyclesTimeRange_function(NumCycles)), NumCycles)
                    
                    'Subtracting 202 blank average from external standard signal
                    If Isotope202Analyzed_UPb = True Then
                        i = BlkCalc_Sh.Range(BlkColumn2 & Blk1)
                            If i < 0 Then i = 0
                            
                    If Detector202_UPb = "Faraday Cup" Then
                            factor = mVtoCPS_UPb
                    ElseIf Detector202_UPb = "MIC" Then
                            factor = 1
                    End If
                    
                        For Each G In .Range(Raw202Range(NumCycles))
                            If Not G = "" Then
                                G = G * factor - i
                                    If G <= 0 Then
                                        G = 1
                                    End If
                            End If
                        Next
                    End If
                    
                    '204Pb
                    'Subtracting 202 blank average divided by 4.35 from sample (and internal standard) 204 signal
                    'k = -(.Range(RawPb204Range).Column - .Range(RawHg202Range).Column / RatioMercury_UPb)
                    If Isotope204Analyzed_UPb = True Then
                    
                        If Detector204_UPb = "Faraday Cup" Then
                                factor = mVtoCPS_UPb
                            ElseIf Detector204_UPb = "MIC" Then
                                factor = 1
                        End If

                        i = BlkCalc_Sh.Range(BlkColumn2 & Blk1)
                            If i < 0 Then i = 0
                        i2 = BlkCalc_Sh.Range(BlkColumn4Comm & Blk1)
                            If i2 < 0 Then i2 = 0
                            
                        If Isotope202Analyzed_UPb = False Then 'I am not sure if this is necessary, because if the cell is empty, then its value is 0.
                            i = 0
                        End If
                            
                        For Each G In .Range(Raw204Range(NumCycles))
                        
                            If BlkSlp202 = True Then 'Means that 202 from both sample and blank will be considered below
                                Extra202 = .Cells(G.Row, .Range(RawHg202Range).Column)
                            ElseIf BlkSlp202 = False Then
                                Extra202 = 0
                            End If
    
                            If Not G = "" Then
                                G = G * factor - ((i + Extra202) / RatioMercury_UPb) - i2 '- ((BlkCalc_Sh.Range(BlkColumn4 & Blk1) - i / RatioMercury_UPb))
                                    If G <= 0 Then
                                        G = 1
                                    End If
                            End If
                        Next
                    End If
                                        
                    'Subtracting 206 blank average from sample (and internal standard) signal and multiplying signal by mvtoCPS
                    'constant, if 206 was analysed using MIC
'                    If Detector206_UPb = "MIC" Then
                        i = BlkCalc_Sh.Range(BlkColumn6 & Blk1)
                            If i < 0 Then i = 0
'
'                        Else: i = 0
'
'                    End If

                    If Detector206_UPb = "Faraday Cup" Then
                        factor = mVtoCPS_UPb
                    ElseIf Detector206_UPb = "MIC" Then
                        factor = 1
                    End If
                    
                    For Each G In .Range(Raw206Range(NumCycles))
                        If Not G = "" Then
                            G = G * factor - i
                                If G <= 0 Then
                                    G = 1
                                End If
                        End If
                    Next
                                                            
                    'Subtracting 207 blank average from sample (and internal standard) signal
                        i = BlkCalc_Sh.Range(BlkColumn7 & Blk1)
                            If i < 0 Then i = 0
                            
                    If Detector207_UPb = "Faraday Cup" Then
                        factor = mVtoCPS_UPb
                    ElseIf Detector207_UPb = "MIC" Then
                        factor = 1
                    End If

                    For Each G In .Range(Raw207Range(NumCycles))
                        If Not G = "" Then
                            G = G * factor - i
                                If G <= 0 Then
                                    G = 1
                                End If
                        End If
                    Next

                    'Subtracting 208 blank average from sample (and internal standard) signal
                    If Isotope208Analyzed_UPb = True Then
                            i = BlkCalc_Sh.Range(BlkColumn8 & Blk1)
                                If i < 0 Then i = 0
                                
                        If Detector208_UPb = "Faraday Cup" Then
                            factor = mVtoCPS_UPb
                        ElseIf Detector208_UPb = "MIC" Then
                            factor = 1
                        End If
    
                        For Each G In .Range(Raw208Range(NumCycles))
                            If Not G = "" Then
                                G = G * factor - i
                                    If G <= 0 Then
                                        G = 0
                                    End If
                            End If
                        Next
                    End If
                    
                    'Multiplying 232 from sample (and internal standard) signal by MvtoCPS constant
                    If Isotope232Analyzed_UPb = True Then
                    
                        If Detector232_UPb = "Faraday Cup" Then
                            factor = mVtoCPS_UPb
                        ElseIf Detector232_UPb = "MIC" Then
                            factor = 1
                        End If
                        
                        For Each G In .Range(Raw232Range(NumCycles))
                            If Not G = "" Then
                                G = G * factor
                                    If G <= 0 Then
                                        G = 1
                                    End If
                            End If
                        Next
                    End If
                    
                    'Multiplying 238 from sample (and internal standard) signal by MvtoCPS constant
                    For Each G In .Range(Raw238Range(NumCycles))
                    
                        If Detector238_UPb = "Faraday Cup" Then
                            factor = mVtoCPS_UPb
                        ElseIf Detector238_UPb = "MIC" Then
                            factor = 1
                        End If
                        
                        If Not G = "" Then
                            G = G * factor
                                If G <= 0 Then
                                    G = 1
                                End If
                        End If
                    Next

                    If CallCommonCalc = True Then
                        Call CommonCalcSlpExtStd_BlkCorr(WBSlp.Worksheets(1), C, SpotRaster_UPb.Value, Blk1)
                    End If
                    
                End With
            
    If CloseAnalysis = True Then
        WBSlp.Close savechanges:=False
    End If
        
End Sub


Sub CommonCalcSlpExtStd_BlkCorr(Sh As Worksheet, ByVal C As Integer, AcquisitionSpotRaster As String, ByVal Blk1 As Integer, Optional ByVal Blk2 As Integer)

    'Sh is the worksheet with raw data and c is a counter for the lines where data will be pasted.

    'Based on signals less blank now the necessary ratios (68, 76, 64, 74, 28) or simple averages (4) will
    'be calculated. The only difference betweeen samples (and internal standard) and external standard is
    'about blank reduction. Besides that, all ratios and erros are calculated on the same way.
    
    'Error propagation here is calculated based on Horstwood (2008), Short Course Series, equation 4.
    'However, only to 75 and 28 ratios is necessary to propagate some error. All other ratios are only
    'simple standard deviations, considering that blank error must be only propagated in SlpStdCorr
    
    'ALL UNCERTANTIES ARE CALCULATED AS PERCENTAGE
    
    Dim a As Double
    Dim b As Double
    Dim OriginalY_ValuesRange As Range 'The range before calling the NonEmptyCellsRange procedure.
    Dim OriginalX_ValuesRange As Range 'The range before calling the NonEmptyCellsRange procedure.
    Dim Y_ValuesRange As Range
    Dim X_ValuesRange As Range
    Dim Y2_ValuesRange As Range
    Dim ClearRange As Range
    Dim temp As Long
    Dim NumCycles As Integer
    
    NumCycles = update_numcycles(Sh.Parent)

     RawPb206Range_updated = Raw206Range(NumCycles)
     
     If Isotope208Analyzed_UPb = True Then
        RawPb208Range_updated = Raw208Range(NumCycles)
     End If
     
     If Isotope232Analyzed_UPb = True Then
        RawTh232Range_updated = Raw232Range(NumCycles)
     End If
     
     RawU238Range_updated = Raw238Range(NumCycles)
     
     If Isotope202Analyzed_UPb = True Then
         RawHg202Range_updated = Raw202Range(NumCycles)
     End If
     
     If Isotope204Analyzed_UPb = True Then
        RawPb204Range_updated = Raw204Range(NumCycles)
     End If
     
     RawPb207Range_updated = Raw207Range(NumCycles)
     RawCyclesTimeRange_updated = RawCyclesTimeRange_function(NumCycles)
    
    With Sh
                
        'In the beginning, Y_ValuesRange and OriginalY_ValuesRange are set to the same range but this changes during code execution
        Set Y_ValuesRange = .Range(.Range(CalculationFirstCell), .Range(CalculationColumn & NumCycles))
        Set OriginalY_ValuesRange = .Range(.Range(CalculationFirstCell), .Range(CalculationColumn & NumCycles))
        
        Set X_ValuesRange = .Range(RawCyclesTimeRange_updated)
        Set OriginalX_ValuesRange = .Range(RawCyclesTimeRange_updated)
        
        Set ClearRange = .Range(OriginalY_ValuesRange, OriginalY_ValuesRange.Offset(, 2))
        
        'Ratio 68
        
        ClearRange.ClearContents
        
        Call MatchValidRangeItems(.Range(RawPb206Range_updated), .Range(RawU238Range_updated), OriginalX_ValuesRange, Sh, .Range(CalculationFirstCell))
        Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
        Set Y2_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 1), OriginalY_ValuesRange.Offset(, 1).Item(1), Sh, True)
        Set X_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 2), OriginalY_ValuesRange.Offset(, 2).Item(1), Sh, True)
        
            Y2_ValuesRange.Copy
                Y_ValuesRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide
                                
                Select Case AcquisitionSpotRaster
                    Case "Spot"
                        'Intercept of 68 trend
                        SlpStdBlkCorr_Sh.Range(Column68 & C) = WorksheetFunction.Intercept(Y_ValuesRange, X_ValuesRange)
                            '68 intercept error multiplied by student´s t factor for 68% confidence
                            SlpStdBlkCorr_Sh.Range(Column681Std & C) = LineFitInterceptError(Y_ValuesRange, X_ValuesRange) * _
                                WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 2)
                        
                        'R
                        SlpStdBlkCorr_Sh.Range(Column68R & C) = WorksheetFunction.Pearson(Y_ValuesRange, X_ValuesRange)
                        'R2
                        SlpStdBlkCorr_Sh.Range(Column68R2 & C) = WorksheetFunction.Power(SlpStdBlkCorr_Sh.Range(Column68R & C), 2)
                            
                    Case "Raster"
                        '68 average
                        SlpStdBlkCorr_Sh.Range(Column68 & C) = WorksheetFunction.Average(Y_ValuesRange)
                        
                            '68 average error propagation multiplied by student´s t factor for 68% confidence
                            SlpStdBlkCorr_Sh.Range(Column681Std & C) = (WorksheetFunction.StDev_S(Y_ValuesRange) / Sqr(WorksheetFunction.count(Y_ValuesRange)) * _
                                WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1))
                                                                        
                        'R
                        SlpStdBlkCorr_Sh.Range(Column68R & C) = WorksheetFunction.Pearson(Y_ValuesRange, X_ValuesRange)
                        'R2
                        SlpStdBlkCorr_Sh.Range(Column68R2 & C) = WorksheetFunction.Power(SlpStdBlkCorr_Sh.Range(Column68R & C), 2)
                 End Select
                
                'SlpStdBlkCorr_Sh.Range(Column281Std & c).Offset(, 1) = WorksheetFunction.Slope(Y_ValuesRange, X_ValuesRange)
        
        'Ratio 76
        ClearRange.ClearContents
        
        Call MatchValidRangeItems(.Range(RawPb207Range_updated), .Range(RawPb206Range_updated), OriginalX_ValuesRange, Sh, .Range(CalculationFirstCell))
        Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
        Set Y2_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 1), OriginalY_ValuesRange.Offset(, 1).Item(1), Sh, True)
        Set X_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 2), OriginalY_ValuesRange.Offset(, 2).Item(1), Sh, True)
        
            Y2_ValuesRange.Copy
                Y_ValuesRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide
                    
                    'SlpStdBlkCorr_Sh.Range(Column76 & c) = WorksheetFunction.Average(Y_ValuesRange)                         'Average of 76 ratio
                    SlpStdBlkCorr_Sh.Range(Column76 & C) = WorksheetFunction.Intercept(Y_ValuesRange, X_ValuesRange)
                                                                  
            '76 average error
            
            'SlpStdBlkCorr_Sh.Range(Column761Std & c) = WorksheetFunction.StDev_S(Y_ValuesRange)
            SlpStdBlkCorr_Sh.Range(Column761Std & C) = LineFitInterceptError(Y_ValuesRange, X_ValuesRange) * WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 2)
        
        'Ratio 75
        SlpStdBlkCorr_Sh.Range(Column75 & C) = _
        SlpStdBlkCorr_Sh.Range(Column68 & C) * SlpStdBlkCorr_Sh.Range(Column76 & C) * RatioUranium_UPb
        
'        Debug.Assert SlpStdBlkCorr_Sh.Range(Column75 & c) <= 30
        
            '75 error
            SlpStdBlkCorr_Sh.Range(Column751Std & C) = _
                SlpStdBlkCorr_Sh.Range(Column75 & C) * _
                Sqr((SlpStdBlkCorr_Sh.Range(Column681Std & C) / SlpStdBlkCorr_Sh.Range(Column68 & C)) ^ 2 + _
                (SlpStdBlkCorr_Sh.Range(Column761Std & C) / SlpStdBlkCorr_Sh.Range(Column76 & C)) ^ 2)
            
        'Rho
            SlpStdBlkCorr_Sh.Range(Column7568Rho & C) = _
            (SlpStdBlkCorr_Sh.Range(Column681Std & C) / SlpStdBlkCorr_Sh.Range(Column68 & C) / _
            (SlpStdBlkCorr_Sh.Range(Column751Std & C) / SlpStdBlkCorr_Sh.Range(Column75 & C)))

        '202 signal intensity
        If Isotope202Analyzed_UPb = True Then
            ClearRange.ClearContents
            
                .Range(RawHg202Range_updated).Copy Destination:=OriginalY_ValuesRange.Item(1)
                    Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
    
    '            '202 average error propagation
    
                        On Error Resume Next
                        
                        SlpStdBlkCorr_Sh.Range(Column2 & C) = WorksheetFunction.Average(Y_ValuesRange) '202 average
                        
                        If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column2 & C) = "n.a."
                        End If
                
                '202 average error propagation
                        
                        Err.Clear
                        
                        SlpStdBlkCorr_Sh.Range(Column21Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                                Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                        
                        If SlpStdBlkCorr_Sh.Range(Column2 & C) = "n.a." Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column21Std & C) = "n.a."
                        End If
                        
                        On Error GoTo 0
        End If
            
           
        '204Pb signal intensity
        If Isotope204Analyzed_UPb = True Then
            ClearRange.ClearContents
            
                .Range(RawPb204Range_updated).Copy Destination:=OriginalY_ValuesRange.Item(1)
                    Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
    
    '            '204 average error propagation
                
                        On Error Resume Next
                        
                        SlpStdBlkCorr_Sh.Range(Column4 & C) = WorksheetFunction.Average(Y_ValuesRange) '204 average
                        
                        If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column4 & C) = "n.a."
                        End If
                
                '204 average error propagation
                        
                        Err.Clear
                        
                        SlpStdBlkCorr_Sh.Range(Column41Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                                Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                        
                        If SlpStdBlkCorr_Sh.Range(Column4 & C) = "n.a." Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column41Std & C) = "n.a."
                        End If
                        
                        Err.Clear
                        
                        On Error GoTo 0
        End If
            
            
        '206Pb signal intensity
                            
        ClearRange.ClearContents
        
            .Range(RawPb206Range_updated).Copy Destination:=OriginalY_ValuesRange.Item(1)
                Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1).Item(1), Sh, True)

'            '206 average error propagation
'            SlpStdBlkCorr_Sh.Range(Column61Std & c) = WorksheetFunction.StDev_S(Y_ValuesRange)
            
                    On Error Resume Next
                    
                    SlpStdBlkCorr_Sh.Range(Column6 & C) = WorksheetFunction.Average(Y_ValuesRange) '206 average
                    
                    If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                        SlpStdBlkCorr_Sh.Range(Column6 & C) = "n.a."
                    End If
            
            '206 average error propagation
                    
                    Err.Clear
                    
                    SlpStdBlkCorr_Sh.Range(Column61Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                        WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                            Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                    
                    If SlpStdBlkCorr_Sh.Range(Column6 & C) = "n.a." Or Err.Number <> 0 Then
                        SlpStdBlkCorr_Sh.Range(Column61Std & C) = "n.a."
                    End If
                    
                    On Error GoTo 0
            
        '207Pb signal intensity
                            
        'SlpStdBlkCorr_Sh.Range(Column7 & c) = WorksheetFunction.Average(.Range(RawPb207Range_updated))
        ClearRange.ClearContents
'        OriginalY_ValuesRange.Clear: OriginalY_ValuesRange.Offset(, 1).Clear 'Cleaning columns used to calculate
        
            .Range(RawPb207Range_updated).Copy Destination:=OriginalY_ValuesRange.Item(1)
                Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
'                    SlpStdBlkCorr_Sh.Range(Column7 & c) = WorksheetFunction.Average(Y_ValuesRange) '207 average
'
'
'            '207 average error propagation
'            SlpStdBlkCorr_Sh.Range(Column71Std & c) = WorksheetFunction.StDev_S(Y_ValuesRange)
            
                    On Error Resume Next
                    
                    SlpStdBlkCorr_Sh.Range(Column7 & C) = WorksheetFunction.Average(Y_ValuesRange) '207 average
                    
                    If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                        SlpStdBlkCorr_Sh.Range(Column7 & C) = "n.a."
                    End If
            
            '207 average error propagation
            
                    Err.Clear
                    
                    SlpStdBlkCorr_Sh.Range(Column71Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                        WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                            Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                    
                    If SlpStdBlkCorr_Sh.Range(Column7 & C) = "n.a." Or Err.Number <> 0 Then
                        SlpStdBlkCorr_Sh.Range(Column71Std & C) = "n.a."
                    End If
                    
                    On Error GoTo 0

        '208Pb signal intensity
        If Isotope208Analyzed_UPb = True Then
            'SlpStdBlkCorr_Sh.Range(Column8 & c) = WorksheetFunction.Average(.Range(RawPb208Range_updated))
            OriginalY_ValuesRange.Clear: OriginalY_ValuesRange.Offset(, 1).Clear 'Cleaning columns used to calculate
            
                .Range(RawPb208Range_updated).Copy Destination:=OriginalY_ValuesRange.Item(1)
                    Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
                        
                        On Error Resume Next
                        
                        SlpStdBlkCorr_Sh.Range(Column8 & C) = WorksheetFunction.Average(Y_ValuesRange)
                        
                        If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column8 & C) = "n.a."
                        End If
                
                '208 average error propagation
                        
                        Err.Clear
                        
                        SlpStdBlkCorr_Sh.Range(Column81Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                                Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                            
                        
                        If SlpStdBlkCorr_Sh.Range(Column8 & C) = "n.a." Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column81Std & C) = "n.a."
                        End If
                        
                        On Error GoTo 0
        End If
        
        '232Th signal intensity
        If Isotope232Analyzed_UPb = True Then
                            
            'SlpStdBlkCorr_Sh.Range(Column32 & c) = WorksheetFunction.Average(.Range(RawTh232Range_updated))
            OriginalY_ValuesRange.Clear: OriginalY_ValuesRange.Offset(, 1).Clear 'Cleaning columns used to calculate
            
                .Range(RawTh232Range_updated).Copy Destination:=OriginalY_ValuesRange.Item(1)
                    Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
    '                    SlpStdBlkCorr_Sh.Range(Column32 & c) = WorksheetFunction.Average(Y_ValuesRange) '232 average
    
    '            '232 average error propagation
    '            SlpStdBlkCorr_Sh.Range(Column321Std & c) = WorksheetFunction.StDev_S(Y_ValuesRange)
                
                        On Error Resume Next
                        
                        SlpStdBlkCorr_Sh.Range(Column32 & C) = WorksheetFunction.Average(Y_ValuesRange) '232 average
                        
                        If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column32 & C) = "n.a."
                        End If
                
                '232 average error propagation
                        
                        Err.Clear
                        
                        SlpStdBlkCorr_Sh.Range(Column321Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                                Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                        
                        If SlpStdBlkCorr_Sh.Range(Column32 & C) = "n.a." Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column321Std & C) = "n.a."
                        End If
                        
                        On Error GoTo 0
        End If
            
        '238U signal intensity
                            
        'SlpStdBlkCorr_Sh.Range(Column38 & c) = WorksheetFunction.Average(.Range(RawU238Range_updated))
        OriginalY_ValuesRange.Clear: OriginalY_ValuesRange.Offset(, 1).Clear 'Cleaning columns used to calculate
        
            .Range(RawU238Range_updated).Copy Destination:=OriginalY_ValuesRange.Item(1)
                Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
                    SlpStdBlkCorr_Sh.Range(Column38 & C) = WorksheetFunction.Average(Y_ValuesRange) '238 average

            
            '238 average error propagation
            SlpStdBlkCorr_Sh.Range(Column381Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                    Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
            
                    On Error Resume Next
                    
                    SlpStdBlkCorr_Sh.Range(Column38 & C) = WorksheetFunction.Average(Y_ValuesRange) '238 average
                    
                    If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                        SlpStdBlkCorr_Sh.Range(Column38 & C) = "n.a."
                    End If
            
            '238 average error propagation
                    
                    Err.Clear
                    
                    SlpStdBlkCorr_Sh.Range(Column381Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                        WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                            Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                    
                    If SlpStdBlkCorr_Sh.Range(Column38 & C) = "n.a." Or Err.Number <> 0 Then
                        SlpStdBlkCorr_Sh.Range(Column381Std & C) = "n.a."
                    End If
                    
                    On Error GoTo 0

        If Isotope204Analyzed_UPb = True Then
            'Ratio 64
    
            OriginalY_ValuesRange.Clear: OriginalY_ValuesRange.Offset(, 1).Clear 'Cleaning columns used to calculate
            
            Call MatchValidRangeItems(.Range(RawPb206Range_updated), .Range(RawPb204Range_updated), OriginalX_ValuesRange, Sh, .Range(CalculationFirstCell))
            Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
            Set Y2_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 1), OriginalY_ValuesRange.Offset(, 1).Item(1), Sh, True)
            Set X_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 2), OriginalY_ValuesRange.Offset(, 2).Item(1), Sh, True)
            
                Y2_ValuesRange.Copy
                    Y_ValuesRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide
                    
    '                    SlpStdBlkCorr_Sh.Range(Column64 & c) = WorksheetFunction.Average(Y_ValuesRange) 'Average of 64 ratio
    '
    '            '64 average error propagation
    '            SlpStdBlkCorr_Sh.Range(Column641Std & c) = WorksheetFunction.StDev_S(Y_ValuesRange)
                
                        On Error Resume Next
                        
                        SlpStdBlkCorr_Sh.Range(Column64 & C) = WorksheetFunction.Average(Y_ValuesRange) '64 average
                        
                        If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column64 & C) = "n.a."
                        End If
                
                '64 average error propagation
                        
                        Err.Clear
                        
                        SlpStdBlkCorr_Sh.Range(Column641Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                                Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                        
                        If SlpStdBlkCorr_Sh.Range(Column64 & C) = "n.a." Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column641Std & C) = "n.a."
                        End If
                        
                        
                        On Error GoTo 0
            
            'Ratio 74
            
    '        OriginalY_ValuesRange.Clear 'Cleaning column used to calculate
    '
    '        .Range(RawPb207Range_updated).Copy Destination:=.Range(CalculationFirstCell) 'Copying 207 signal to CalculationColumn
    '            .Range(RawPb204Range_updated).Copy
    '                .Range(CalculationFirstCell).PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide 'Pasting 204 signal divinding 207
                    
            OriginalY_ValuesRange.Clear: OriginalY_ValuesRange.Offset(, 1).Clear 'Cleaning columns used to calculate
            
            Call MatchValidRangeItems(.Range(RawPb207Range_updated), .Range(RawPb204Range_updated), OriginalX_ValuesRange, Sh, .Range(CalculationFirstCell))
            Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
            Set Y2_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 1), OriginalY_ValuesRange.Offset(, 1).Item(1), Sh, True)
            Set X_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 2), OriginalY_ValuesRange.Offset(, 2).Item(1), Sh, True)
            
                Y2_ValuesRange.Copy
                    Y_ValuesRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide
                    
    '                    SlpStdBlkCorr_Sh.Range(Column74 & c) = WorksheetFunction.Average(Y_ValuesRange) 'Average of 74 ratio
    '
    '            '74 average error propagation
    '            SlpStdBlkCorr_Sh.Range(Column741Std & c) = WorksheetFunction.StDev_S(Y_ValuesRange)
                
                        On Error Resume Next
                        
                        SlpStdBlkCorr_Sh.Range(Column74 & C) = WorksheetFunction.Average(Y_ValuesRange) '74 average
                        
                        If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column74 & C) = "n.a."
                        End If
                
                '74 average error propagation
                        
                        Err.Clear
                        
                        SlpStdBlkCorr_Sh.Range(Column741Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                                Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                        
                        If SlpStdBlkCorr_Sh.Range(Column74 & C) = "n.a." Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column741Std & C) = "n.a."
                        End If
                        
                        On Error GoTo 0
        End If

        'Ratio 28
        If Isotope232Analyzed_UPb = True Then
            OriginalY_ValuesRange.Clear: OriginalY_ValuesRange.Offset(, 1).Clear 'Cleaning columns used to calculate
            
            Call MatchValidRangeItems(.Range(RawTh232Range_updated), .Range(RawU238Range_updated), OriginalX_ValuesRange, Sh, .Range(CalculationFirstCell))
            Set Y_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange, OriginalY_ValuesRange.Item(1), Sh, True)
            Set Y2_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 1), OriginalY_ValuesRange.Offset(, 1).Item(1), Sh, True)
            Set X_ValuesRange = NonEmptyCellsRange(OriginalY_ValuesRange.Offset(, 2), OriginalY_ValuesRange.Offset(, 2).Item(1), Sh, True)
            
                Y2_ValuesRange.Copy
                    Y_ValuesRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlDivide
                    
    '                    SlpStdBlkCorr_Sh.Range(Column28 & c) = WorksheetFunction.Average(Y_ValuesRange)
    '
    '            '28 error propagation - No blank correction is necessary for these two isotopes
    '            SlpStdBlkCorr_Sh.Range(Column281Std & c) = WorksheetFunction.StDev_S(Y_ValuesRange)
                
                        On Error Resume Next
                        
                        SlpStdBlkCorr_Sh.Range(Column28 & C) = WorksheetFunction.Average(Y_ValuesRange) '28 average
                        
                        If WorksheetFunction.Average(Y_ValuesRange) = 0 Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column28 & C) = "n.a."
                        End If
                
                '28 average error propagation
                        
                        Err.Clear
                        
                        SlpStdBlkCorr_Sh.Range(Column281Std & C) = WorksheetFunction.StDev_S(Y_ValuesRange) * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(Y_ValuesRange) - 1) / _
                                Sqr(WorksheetFunction.count(Y_ValuesRange)) 'Standard error
                        
                        If SlpStdBlkCorr_Sh.Range(Column28 & C) = "n.a." Or Err.Number <> 0 Then
                            SlpStdBlkCorr_Sh.Range(Column281Std & C) = "n.a."
                        End If
                        
                        On Error GoTo 0
        End If
    End With
    
End Sub
Sub CalcBlank()

    'This program only deal with blank processing, opening every blank file and
    'calculating averages and standard errors, and copying theses calculations
    'to BlkCalc sheet
    
    Dim a As Variant
    Dim C As Integer
    Dim analysis_workbook As Workbook 'The workbook opened
    Dim d As Single
    Dim E As Double
    Dim SearchStr As Integer
    Dim A_Header As Range
    Dim RangeUnionHeaders As Range
    Dim SelectedRange As Range
    Dim CountCells As Integer
    Dim NumCycles As Integer
    
    Dim Blanks() As Integer 'Array with blanks IDs
    Dim VarCovarArray As Variant
    
    ''Application.DisplayAlerts = False
    ''Application.ScreenUpdating = False
    
    If FolderPath_UPb Is Nothing Then
        Call PublicVariables
    End If

    If IsArrayEmpty(BlkFound) = True Then
        Call IdentifyFileType
    End If
    
    If IsArrayEmpty(PathsNamesIDsTimesCycles) = True Then
        Call SetPathsNamesIDsTimesCycles
    End If
                      
    With BlkCalc_Sh: .Cells.Clear: End With
                          
    ReDim Preserve Blanks(1 To UBound(BlkFound) + 1) As Integer
    
    C = 2

    For Each a In BlkFound 'Blanks IDs are copied to a different array (Blanks) which accepts only numbers (IDs)
        Blanks(C - 1) = SamList_Sh.Range(a).Offset(, 1)
        C = C + 1
    Next
            
    'Now, each raw blank data file will be opened and processed
    C = 1
        
    For Each a In Blanks
        C = C + 1
            
            On Error Resume Next
                Set analysis_workbook = Workbooks.Open(PathsNamesIDsTimesCycles(RawDataFilesPaths, a)) 'ActiveWorkbook
                    If Err.Number <> 0 Then
                        MsgBox MissingFile1 & PathsNamesIDsTimesCycles(RawDataFilesPaths, a) & MissingFile2
                            Call UpdateFilesAddresses
                                Call UnloadAll
                                    End
                    End If
            On Error GoTo 0

            Call ClearCycles(analysis_workbook, PathsNamesIDsTimesCycles(Cycles, a)) 'Any cycles that should be discarded from the blanks will be now
            
            ''Application.DisplayAlerts = False
            
                
                'Below, averages of the isotopes cps will be evaluated from the raw data files and copied to UPb reduction workbook
                With BlkCalc_Sh
                
                'On Error GoTo ErrHandler
                
                    NumCycles = update_numcycles(analysis_workbook)
                    
                    .Range(BlkSlpName & C) = PathsNamesIDsTimesCycles(FileName, a)
                    
                    .Range(BlkColumnID & C) = a
                    
                    '202-----------------------------------------------------------------------------------
                    If Isotope202Analyzed_UPb = True Then
                        
                        If Detector202_UPb = "Faraday Cup" Then
                                d = mVtoCPS_UPb
                            ElseIf Detector202_UPb = "MIC" Then
                                d = 1
                        End If
                        
                        Set SelectedRange = analysis_workbook.Sheets(1).Range(Raw202Range(NumCycles))
                        CountCells = WorksheetFunction.count(SelectedRange)
                        
                        E = WorksheetFunction.Average(SelectedRange) * d
                        
                            If Not E < 0 Then .Range(BlkColumn2 & C) = E Else: .Range(BlkColumn2 & C) = 0
                                
                            'Below it is the standard error (standard deviation divided by square root of  number os analyses)
                            'multiplied by student's t factor of the assigned confidence
                            E = WorksheetFunction.StDev_S(SelectedRange) * d * _
                                WorksheetFunction.T_Inv_2T(ConfLevel, CountCells - 1) / _
                                    Sqr(CountCells)
                                    
                                If Not E < 0 Then .Range(BlkColumn21Std & C) = E Else: .Range(BlkColumn21Std & C) = 0
                                
                    End If
                    
                    '204-----------------------------------------------------------------------------------
                    If Isotope204Analyzed_UPb = True Then
                    
                        If Detector204_UPb = "Faraday Cup" Then
                                d = mVtoCPS_UPb
                            ElseIf Detector204_UPb = "MIC" Then
                                d = 1
                        End If
                        
                        Set SelectedRange = analysis_workbook.Sheets(1).Range(Raw204Range(NumCycles))
                        CountCells = WorksheetFunction.count(SelectedRange)
                        
                        E = WorksheetFunction.Average(SelectedRange) * d
                            If Not E < 0 Then .Range(BlkColumn4 & C) = E Else: .Range(BlkColumn4 & C) = 0
                        
                            'Below it is the standard error (standard deviation divided by square root of  number os analyses)
                            'multiplied by student's t factor of the assigned confidence
                            E = WorksheetFunction.StDev_S(SelectedRange) * d * _
                                WorksheetFunction.T_Inv_2T(ConfLevel, CountCells - 1) / _
                                    Sqr(CountCells)
                                    
                                If Not E < 0 Then .Range(BlkColumn41Std & C) = E Else: .Range(BlkColumn41Std & C) = 0
                    End If
                    
                    '206-----------------------------------------------------------------------------------
                    If Detector206_UPb = "Faraday Cup" Then
                            d = mVtoCPS_UPb
                        ElseIf Detector206_UPb = "MIC" Then
                            d = 1
                    End If
                                        
                    Set SelectedRange = analysis_workbook.Sheets(1).Range(Raw206Range(NumCycles))
                    CountCells = WorksheetFunction.count(SelectedRange)

                        .Range(BlkColumn6 & C) = WorksheetFunction.Average(SelectedRange) * d
                        
                            'Below it is the standard error (standard deviation divided by square root of  number os analyses)
                            'multiplied by student's t factor of the assigned confidence
                            .Range(BlkColumn61Std & C) = WorksheetFunction.StDev_S(SelectedRange) * d * _
                                WorksheetFunction.T_Inv_2T(ConfLevel, CountCells - 1) / _
                                    Sqr(CountCells)
                                        
                    '207-----------------------------------------------------------------------------------
                    If Detector207_UPb = "Faraday Cup" Then
                            d = mVtoCPS_UPb
                        ElseIf Detector207_UPb = "MIC" Then
                            d = 1
                    End If
                    
                    Set SelectedRange = analysis_workbook.Sheets(1).Range(Raw207Range(NumCycles))
                    CountCells = WorksheetFunction.count(SelectedRange)

                        E = WorksheetFunction.Average(SelectedRange) * d
                        
                            If Not E < 0 Then .Range(BlkColumn7 & C) = E Else: .Range(BlkColumn7 & C) = 0
                
                            'Below it is the standard error (standard deviation divided by square root of  number os analyses)
                            'multiplied by student's t factor of the assigned confidence
                            E = WorksheetFunction.StDev_S(SelectedRange) * d * _
                            WorksheetFunction.T_Inv_2T(ConfLevel, CountCells - 1) / _
                                Sqr(CountCells)
                                
                            If Not E < 0 Then .Range(BlkColumn71Std & C) = E Else: .Range(BlkColumn71Std & C) = 0
                    
                    '208-----------------------------------------------------------------------------------
                    If Isotope208Analyzed_UPb = True Then
                        
                        If Detector208_UPb = "Faraday Cup" Then
                                d = mVtoCPS_UPb
                            ElseIf Detector208_UPb = "MIC" Then
                                d = 1
                        End If
                    
                        Set SelectedRange = analysis_workbook.Sheets(1).Range(Raw208Range(NumCycles))
                        CountCells = WorksheetFunction.count(SelectedRange)
                        
                            .Range(BlkColumn8 & C) = WorksheetFunction.Average(SelectedRange) * d
                            
                                'Below it is the standard error (standard deviation divided by square root of  number os analyses)
                                'multiplied by student's t factor of the assigned confidence
                                .Range(BlkColumn81Std & C) = WorksheetFunction.StDev_S(SelectedRange) * d * _
                                    WorksheetFunction.T_Inv_2T(ConfLevel, CountCells - 1) / _
                                        Sqr(CountCells)
                    End If
                    '232-----------------------------------------------------------------------------------
                    If Isotope232Analyzed_UPb = True Then
                        
                        If Detector232_UPb = "Faraday Cup" Then
                                d = mVtoCPS_UPb
                            ElseIf Detector232_UPb = "MIC" Then
                                d = 1
                        End If
                    
                        Set SelectedRange = analysis_workbook.Sheets(1).Range(Raw232Range(NumCycles))
                        CountCells = WorksheetFunction.count(SelectedRange)
                                    
                            'Below it is the standard error (standard deviation divided by square root of  number os analyses)
                            'multiplied by student's t factor of the assigned confidence
                            .Range(BlkColumn32 & C) = WorksheetFunction.Average(SelectedRange) * d
        
                                .Range(BlkColumn321Std & C) = WorksheetFunction.StDev_S(SelectedRange) * d * _
                                    WorksheetFunction.T_Inv_2T(ConfLevel, CountCells - 1) / _
                                        Sqr(CountCells) 'Standard error
                    End If
                    '238-----------------------------------------------------------------------------------
                    If Detector238_UPb = "Faraday Cup" Then
                            d = mVtoCPS_UPb
                        ElseIf Detector238_UPb = "MIC" Then
                            d = 1
                    End If
                    
                    Set SelectedRange = analysis_workbook.Sheets(1).Range(Raw238Range(NumCycles))
                    CountCells = WorksheetFunction.count(SelectedRange)
                    
                        .Range(BlkColumn38 & C) = WorksheetFunction.Average(SelectedRange) * d
 
                            'Below it is the standard error (standard deviation divided by square root of  number os analyses)
                            'multiplied by student's t factor of the assigned confidence
                            .Range(BlkColumn381Std & C) = WorksheetFunction.StDev_S(SelectedRange) * d * _
                                WorksheetFunction.T_Inv_2T(ConfLevel, CountCells - 1) / _
                                    Sqr(CountCells) 'Standard error
                    
                    .Range(BlkColumn4Comm & C) = .Range(BlkColumn4 & C) - .Range(BlkColumn2 & C) / RatioMercury_UPb
                    
                        .Range(BlkColumn4Comm1Std & C) = Sqr(.Range(BlkColumn21Std & C) ^ 2 + .Range(BlkColumn41Std & C) ^ 2)
                    
                End With
                                            
        analysis_workbook.Close savechanges:=False 'Close raw blank data file without saving any modifications
        
    Next
        
    With BlkCalc_Sh
        .Range("A" & BlkCalc_HeaderLine + 1, .Range("A" & BlkCalc_HeaderLine + 1).End(xlDown)).NumberFormat = "0"
        .Range("B" & BlkCalc_HeaderLine + 1, .Range("R" & BlkCalc_HeaderLine + 1).End(xlDown)).NumberFormat = "0.00"
        .Range("A" & BlkCalc_HeaderLine, .Range("R" & BlkCalc_HeaderLine).End(xlDown)).HorizontalAlignment = xlCenter
    End With
    
    
Exit Sub

ErrHandler:
    MsgBox "Plese, check the blank raw data. Only numbers in isotopes signal range are accepted."
        End
End Sub

Sub CalcAllSlp_StdCorr()
    
    Dim a As Integer
    Dim counter As Integer
    Dim SamplesID As Range
    Dim Sample As Range
    
    If SlpStdCorr_Sh Is Nothing Then 'We need some public variables, so we must be sure that they were set
        Call PublicVariables
    End If

    With SlpStdCorr_Sh 'The lines below are going to write headers and format cells.
        
        .Cells.Clear
                
        With .Range(StdCorr_Column68R2 & HeaderRow).Characters(Start:=2, Length:=1).Font
            .Superscript = True
            .Bold = True
        End With
            
    End With

    On Error Resume Next
        If AnalysesList(0).Sample = "" Then
            Call LoadSamListMap
        End If
    On Error GoTo 0
    
    counter = StdCorr_HeaderRow + 1
        
    'ID of sample or internal standard will be pasted to SlpStdCorr
    For a = 1 To UBound(AnalysesList) - 1
        
        SlpStdCorr_Sh.Range(StdCorr_ColumnID & counter) = AnalysesList(a).Sample 'ID of sample or internal standard will be pasted to SlpStdCorr
                        
            SlpStdCorr_Sh.Range(StdCorr_TetaFactor & counter) = TetaFactor(a)
            
            counter = counter + 1
    Next
    
    If SlpStdCorr_Sh.Range(StdCorr_ColumnID & StdCorr_HeaderRow + 1).End(xlDown) = "" Then
        MsgBox "There are no analyses in SlpStdBlkCorr sheet. Please, check it."
            Application.GoTo SlpStdBlkCorr_Sh.Range("A1")
                End
        Else
            Set SamplesID = SlpStdCorr_Sh.Range(StdCorr_ColumnID & StdCorr_HeaderRow + 1, SlpStdCorr_Sh.Range(StdCorr_ColumnID & StdCorr_HeaderRow + 1).End(xlDown))
    End If
    
    For Each Sample In SamplesID
        
        Call CalcSlp_StdCorr(Sample, Sample.Row, Sample.Offset(, 2))
    
    Next

End Sub


Sub CalcSlp_StdCorr(ByVal a As Integer, ByVal C As Integer, ByVal Teta As Double)
    
    'This program calculates the sample ratios corrected by the external standard, as well the errors.
    'A is the sample ID, C is the row where data should be pasted in SlpStdCorr_Sh and Teta is the
    'constant calculated considering the analysis time of sample and external standard.
    
    'Error propagation here is calculated based on Horstwood (2008), Short Course Series, equation 4.
    'ALL UNCERTANTIES ARE CALCULATED AS PERCENTAGE.
    
    'Updated 11/09/2015 - 68 75 concordance added
    
    Dim E As Integer
    Dim counter As Integer
    Dim AnalysesListNumber As Integer
    
    Dim P As Range 'Column68 or Column76 or Column2 or Column4 or Column64 or Column74 or Column28
    Dim P2 As Range 'Column68 or Column76 or Column2 or Column4 or Column64 or Column74 or Column28
    Dim PP As Range 'Column68 or Column76 or Column2 or Column4 or Column64 or Column74 or Column28
    Dim Q As Range 'Standard deviation of the measured quantity to which P is equal
    Dim QQ As Range 'Standard deviation of the measured quantity to which P is equal
        
    Dim Blk21Std_Blk1 As Double '202 1 std for Blk1
    Dim Blk2_Blk1 As Double '202 for Blk1
    Dim Blk41Std_Blk1 As Double '204 1 std for Blk1
    Dim Blk4_Blk1 As Double '204 for Blk1
    Dim Blk61Std_Blk1 As Double '206 1 std for Blk1
    Dim Blk6_Blk1 As Double '206 for Blk1
    Dim Blk71Std_Blk1 As Double '207 1 std for Blk1
    Dim Blk7_Blk1 As Double '207 for Blk1
    Dim Blk21Std_Blk2 As Double '202 1 std for Blk2
    Dim Blk2_Blk2 As Double '202 for Blk2
    Dim Blk41Std_Blk2 As Double '204 1 std for Blk2
    Dim Blk4_Blk2 As Double '204 for Blk2
    Dim Blk61Std_Blk2 As Double '206 1 std for Blk2
    Dim Blk6_Blk2 As Double '206 for Blk2
    Dim Blk71Std_Blk2 As Double '207 1 std for Blk2
    Dim Blk7_Blk2 As Double '207 for Blk2
    
    Dim Std681Std_Std1 As Double '68 1 std for Std1
    Dim Std68_Std1 As Double '68 for Std1
    Dim Std761Std_Std1 As Double '76 1 std for Std1
    Dim Std76_Std1 As Double '76 for Std1
    Dim Std751Std_Std1 As Double '75 1 std for Std1
    Dim Std75_Std1 As Double '75 for Std1
    Dim Std681Std_Std2 As Double '68 1 std for Std2
    Dim Std68_Std2 As Double '68 for Std2
    Dim Std761Std_Std2 As Double '76 1 std for Std2
    Dim Std76_Std2 As Double '76 for Std2
    Dim Std751Std_Std2 As Double '75 1 std for Std2
    Dim Std75_Std2 As Double '75 for Std2
    
    Dim ExtStd68 As Double 'ExtStd68 ratio
    Dim ExtStd681Std As Double 'ExtStd681Std
    Dim ExtStd75 As Double 'ExtStd75 ratio
    Dim ExtStd751Std As Double 'ExtStd751Std
    Dim ExtStd76 As Double 'ExtStd76 ratio
    Dim ExtStd761Std As Double 'ExtStd761Std
    
    Dim ExtStd68Reproducibility As Double
    Dim ExtStd75Reproducibility As Double
    Dim ExtStd76Reproducibility As Double
        
    Dim Names As Variant
    Dim StdName As String
    Dim Slp As Integer 'Sample row in SlpStdBlkCorr_Sh
    Dim Blk1 As Integer 'Blk1 row in BlkCorr_Sh
    Dim Blk2 As Integer 'Blk2 row in BlkCorr_Sh
    Dim Std1 As Integer 'Std1 row in SlpStdBlkCorr_Sh
    Dim Std2 As Integer 'Std2 row in SlpStdBlkCorr_Sh
    Dim SampleID As Range 'Samples ID in the column ID of SlpStdCorr
    Dim BlankID As Range 'Blanks IDs in the column ID of BlkCalc_Sh
    Dim StandardID As Range 'Blanks IDs in the column ID of BlkCalc_Sh
    
    Dim d As Integer
    
    Dim SearchStr As Integer 'Variable used to search for (%) in headers
    Dim RangeUnionHeaders As Range 'Range with headers
    Dim A_Header As Range 'Variable used to loop through headers
'Debug.Assert a <> 184
    If SlpStdCorr_Sh Is Nothing Then
        Call PublicVariables
    End If

    If IsArrayEmpty(StdFound) = True Then
        Call IdentifyFileType
    End If

    If IsArrayEmpty(PathsNamesIDsTimesCycles) Then
        Call SetPathsNamesIDsTimesCycles
    End If

    'on error resume Next
        If Not WorksheetFunction.IsNumber(AnalysesList(1).Sample) = True Then
            Call LoadSamListMap
        End If
    'On Error GoTo 0
        
    'The code below clears the entire SlpStdBlkCorr sheet and changes the uncertanties headers units to "(abs)".
        SearchStr = InStr(SlpStdCorr_Sh.Range(StdCorr_Column681Std & StdCorr_HeaderRow), "(%)")
    
    If SearchStr <> 0 Then
        With SlpStdCorr_Sh
             
            Set RangeUnionHeaders = Application.Union( _
            .Range(StdCorr_Column681Std & StdCorr_HeaderRow), _
            .Range(StdCorr_Column761Std & StdCorr_HeaderRow), _
            .Range(StdCorr_Column751Std & StdCorr_HeaderRow), _
            .Range(StdCorr_Column21Std & StdCorr_HeaderRow), _
            .Range(StdCorr_Column41Std & StdCorr_HeaderRow), _
            .Range(StdCorr_Column641Std & StdCorr_HeaderRow), _
            .Range(StdCorr_Column741Std & StdCorr_HeaderRow), _
            .Range(StdCorr_Column281Std & StdCorr_HeaderRow))
    
            For Each A_Header In RangeUnionHeaders
                A_Header.Value = Replace(A_Header.Value, " (%)", " (abs)")
            Next
        End With
    End If
    
    With SlpStdCorr_Sh
        .Range(StdCorr_ColumnID & C) = PathsNamesIDsTimesCycles(ID, a)
        .Range(StdCorr_SlpName & C) = PathsNamesIDsTimesCycles(FileName, a)
    End With
    
    counter = 1
    
    'The following 6 lines are used just to check if UPbStd was initialized
    On Error Resume Next
        counter = LBound(UPbStd)
            If Err.Number <> 0 Then
                Call Load_UPbStandardsTypeList
            End If
    On Error GoTo 0
    
    'The 6 lines below are necessary to adentify the number of the external standard in UpbStd
    For counter = LBound(UPbStd) To UBound(UPbStd)
        If UPbStd(counter).StandardName = ExternalStandard_UPb Then
            StdName = counter
                counter = UBound(UPbStd)
        End If
    Next

    counter = 2
            
    For E = 1 To UBound(AnalysesList) 'For each structure used to find the sample inside AnalysesList
        If a = AnalysesList(E).Sample Then
            AnalysesListNumber = E 'Using this variable I am able to retrieve from AnalysesList all the IDs that I must know
            E = UBound(AnalysesList) 'A beautiful solution to end the if structure
        End If
    Next

    'Lines of code necessary to find Slp row in SlpStdBlkCorr_Sh
    With SlpStdBlkCorr_Sh

        If .Range(ColumnID & HeaderRow + 1).End(xlDown) = "" Then
            MsgBox ("There is no data in SlpStdBlkCorr sheet. Please, check it.")
                Application.GoTo SlpStdBlkCorr_Sh.Range("A1")
                    End
        End If

        For Each SampleID In .Range(ColumnID & HeaderRow + 1, .Range(ColumnID & HeaderRow + 1).End(xlDown))

            If AnalysesList(AnalysesListNumber).Sample = SampleID Then
                Slp = SampleID.Row
            End If

        Next
    
        If Slp = 0 Then
            MsgBox "Sample with ID=" & SampleID & " was not found in SlpStdBlkCorr sheet."
                End
        End If
        
    End With

    'Lines of code necessary to find Blk1 and Blk2 rows in BlkCalc_Sh
    With BlkCalc_Sh

        If .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown) = "" Then
            MsgBox ("You need at least two blanks to reduce your data.")
                Application.GoTo BlkCalc_Sh.Range("A1")
                    End
        End If

        For Each BlankID In .Range(BlkColumnID & BlkCalc_HeaderLine + 1, .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown))

            If AnalysesList(AnalysesListNumber).Blk1 = BlankID Then
                Blk1 = BlankID.Row
            End If

            If BlanksRecordedSamples_UPb = False Then
                If AnalysesList(AnalysesListNumber).Blk2 = BlankID Then
                    Blk2 = BlankID.Row
                End If
            End If

        Next
        
        If BlanksRecordedSamples_UPb = False And Blk2 = 0 Then
            MsgBox "Blank with ID=" & BlankID & " was not found in SlpStdBlkCorr sheet."
                End
        End If
        
        If Blk1 = 0 Then
            MsgBox "Blank with ID=" & BlankID & " was not found in SlpStdBlkCorr sheet."
                End
        End If
    
    End With

    'Lines of code necessary to find Std1 and Std2 rows in SlpStdBlkCorr_Sh
    With SlpStdBlkCorr_Sh

        If .Range(ColumnID & HeaderRow + 1) = "" Then
            MsgBox ("There is no data in SlpStdBlkCorr sheet. Please, check it.")
                Application.GoTo SlpStdBlkCorr_Sh.Range("A1")
                    End
        End If

        For Each StandardID In .Range(ColumnID & HeaderRow + 1, .Range(ColumnID & HeaderRow + 1).End(xlDown))

            If AnalysesList(AnalysesListNumber).Std1 = StandardID Then
                Std1 = StandardID.Row
            End If

            If AnalysesList(AnalysesListNumber).Std2 = StandardID Then
                Std2 = StandardID.Row
            End If

        Next
        
        If Std1 = 0 Or Std2 = 0 Then
            MsgBox "External standard with ID=" & StandardID & " was not found in SlpStdBlkCorr sheet."
                End
        End If
    
    End With
        
    ExtStd68 = UPbStd(StdName).Ratio68
    ExtStd75 = UPbStd(StdName).Ratio75
    ExtStd76 = UPbStd(StdName).Ratio76

    On Error Resume Next
    
        'Std1----------------------------------------------------------------
        
            Std76_Std1 = SlpStdBlkCorr_Sh.Range(Column76 & Std1)
            
                If Err.Number <> 0 Then
                    Std76_Std1 = 0
                End If
            
            Err.Clear
                    
            Std68_Std1 = SlpStdBlkCorr_Sh.Range(Column68 & Std1)
            
                If Err.Number <> 0 Then
                    Std68_Std1 = 0
                End If
            
            Err.Clear
                    
            Std75_Std1 = Std68_Std1 * Std76_Std1 * RatioUranium_UPb
            
                If Err.Number <> 0 Then
                    Std75_Std1 = 0
                End If
            
        'Std2---------------------------------------------------------------

            Err.Clear

            Std76_Std2 = SlpStdBlkCorr_Sh.Range(Column76 & Std2)
            
                If Err.Number <> 0 Then
                    Std76_Std2 = 0
                End If
            
            Err.Clear
                                
            Std68_Std2 = SlpStdBlkCorr_Sh.Range(Column68 & Std2)
            
                If Err.Number <> 0 Then
                    Std68_Std2 = 0
                End If
            
            Err.Clear
            
            Std75_Std2 = Std68_Std2 * Std76_Std2 * RatioUranium_UPb
            
                If Err.Number <> 0 Then
                    Std75_Std2 = 0
                End If
        
    On Error GoTo 0
    
    Select Case ErrBlank_UPb
    
        Case True 'User wants that blank error be propagated into sample uncertanties
                        
            'For Blk1
            Blk21Std_Blk1 = BlkCalc_Sh.Range(BlkColumn21Std & Blk1)
                Blk2_Blk1 = BlkCalc_Sh.Range(BlkColumn2 & Blk1)
                    
                    If Blk2_Blk1 <= 0 Then
                        Blk21Std_Blk1 = 0
                        Blk2_Blk1 = 1
                    End If
            
            Blk41Std_Blk1 = BlkCalc_Sh.Range(BlkColumn41Std & Blk1)
                Blk4_Blk1 = BlkCalc_Sh.Range(BlkColumn4 & Blk1)
                
                    If Blk4_Blk1 <= 0 Then
                        Blk41Std_Blk1 = 0
                        Blk4_Blk1 = 1
                    End If
            
            If Detector206_UPb = "MIC" Then
                Blk61Std_Blk1 = BlkCalc_Sh.Range(BlkColumn61Std & Blk1)
                    Blk6_Blk1 = BlkCalc_Sh.Range(BlkColumn6 & Blk1)
                
                    If Blk6_Blk1 <= 0 Then
                        Blk61Std_Blk1 = 0
                        Blk6_Blk1 = 1
                    End If
                Else
                    Blk61Std_Blk1 = 0
                    Blk6_Blk1 = 1
            End If
                
            Blk71Std_Blk1 = BlkCalc_Sh.Range(BlkColumn71Std & Blk1)
                Blk7_Blk1 = BlkCalc_Sh.Range(BlkColumn7 & Blk1)
                
                    If Blk7_Blk1 <= 0 Then
                        Blk71Std_Blk1 = 0
                        Blk7_Blk1 = 1
                    End If
            
            'For Blk2
            Blk21Std_Blk2 = BlkCalc_Sh.Range(BlkColumn21Std & Blk2)
                Blk2_Blk2 = BlkCalc_Sh.Range(BlkColumn2 & Blk2)
            
                    If Blk2_Blk2 <= 0 Then
                        Blk21Std_Blk2 = 0
                        Blk2_Blk2 = 1
                    End If
            
            Blk41Std_Blk2 = BlkCalc_Sh.Range(BlkColumn41Std & Blk2)
                Blk4_Blk2 = BlkCalc_Sh.Range(BlkColumn4 & Blk2)
            
                    If Blk4_Blk2 <= 0 Then
                        Blk41Std_Blk2 = 0
                        Blk4_Blk2 = 1
                    End If
            
            If Detector206_UPb = "MIC" Then
                Blk61Std_Blk2 = BlkCalc_Sh.Range(BlkColumn61Std & Blk2)
                    Blk6_Blk2 = BlkCalc_Sh.Range(BlkColumn6 & Blk2)
                    
                    If Blk6_Blk2 <= 0 Then
                        Blk61Std_Blk2 = 0
                        Blk6_Blk2 = 1
                    End If
                        
                Else
                    Blk61Std_Blk2 = 0
                    Blk6_Blk2 = 1
            End If

            Blk71Std_Blk2 = BlkCalc_Sh.Range(BlkColumn71Std & Blk2)
                Blk7_Blk2 = BlkCalc_Sh.Range(BlkColumn7 & Blk2)
                
                    If Blk7_Blk2 <= 0 Then
                        Blk71Std_Blk2 = 0
                        Blk7_Blk2 = 1
                    End If
                     
        Case False 'User don't want that blank error be propagated into sample uncertanties
        
            Blk21Std_Blk1 = 0
            Blk2_Blk1 = 1
            Blk41Std_Blk1 = 0
            Blk4_Blk1 = 1
            Blk61Std_Blk1 = 0
            Blk6_Blk1 = 1
            Blk71Std_Blk1 = 0
            Blk7_Blk1 = 1
            Blk21Std_Blk2 = 0
            Blk2_Blk2 = 1
            Blk41Std_Blk2 = 0
            Blk4_Blk2 = 1
            Blk61Std_Blk2 = 0
            Blk6_Blk2 = 1
            Blk71Std_Blk2 = 0
            Blk7_Blk2 = 1
            
    End Select
    
    Select Case ExtStdRepro_UPb
        
        Case True
            
            If ExtStd68MSWD > 1 Then
                ExtStd68Reproducibility = Sqr(ExtStd68MSWD)
            Else
                ExtStd68Reproducibility = 1
            End If
            
            If ExtStd75MSWD > 1 Then
                ExtStd75Reproducibility = Sqr(ExtStd75MSWD)
            Else
                ExtStd75Reproducibility = 1
            End If
            
            If ExtStd76MSWD > 1 Then
                ExtStd76Reproducibility = Sqr(ExtStd76MSWD)
            Else
                ExtStd76Reproducibility = 1
            End If
            
        Case False
        
            ExtStd68Reproducibility = 1
            ExtStd75Reproducibility = 1
            ExtStd76Reproducibility = 1
            
    End Select

    Select Case ErrExtStdCert_UPb 'User wants that external standard (certified) error be propagated into sample uncertanties
        Case True
            
            ExtStd681Std = UPbStd(StdName).Ratio68Error
            ExtStd751Std = UPbStd(StdName).Ratio75Error
            ExtStd761Std = UPbStd(StdName).Ratio76Error
        
        Case False
            
            ExtStd681Std = 0
            ExtStd751Std = 0
            ExtStd761Std = 0

    End Select


    Select Case ErrExtStd_UPb
    
        'User wants that external standard (analyzed) error be propagated into sample uncertanties
        Case True
            
            'If error is not 1 StDev, then it is divided by 2
            If UPbStd(StdName).RatioErrors12s = 2 Then
                ExtStd681Std = ExtStd681Std / 2
                ExtStd751Std = ExtStd751Std / 2
                ExtStd761Std = ExtStd761Std / 2
            End If
            
            'If error is not absolute, then it is converted to absolute
            If UPbStd(StdName).RatioErrorsAbs = False Then
                ExtStd681Std = ExtStd68 * ExtStd681Std / 100
                ExtStd751Std = ExtStd75 * ExtStd751Std / 100
                ExtStd761Std = ExtStd76 * ExtStd761Std / 100
            End If
                        
            'Std1---------------------------------------------------------------------------
            On Error Resume Next
                        
                Std761Std_Std1 = _
                    (Std76_Std1 ^ (1 - Teta)) * _
                    (1 - Teta) * _
                    ((SlpStdBlkCorr_Sh.Range(Column761Std & Std1) / (Std76_Std1)))

'                Std761Std_Std1 = SlpStdBlkCorr_Sh.Range(Column761Std & Std1) / (Std76_Std1)

                    
                If Err.Number <> 0 Then
                    Std761Std_Std1 = 0
                End If
            
            Err.Clear
            
                Std681Std_Std1 = _
                    (Std68_Std1 ^ (1 - Teta)) * _
                    (1 - Teta) * _
                    ((SlpStdBlkCorr_Sh.Range(Column681Std & Std1) / (Std68_Std1)))
                    
'                Std681Std_Std1 = SlpStdBlkCorr_Sh.Range(Column681Std & Std1) / (Std68_Std1)

                
                If Err.Number <> 0 Then
                    Std681Std_Std1 = 0
                End If
                            
            Err.Clear
                
                Std751Std_Std1 = Std75_Std1 * Sqr( _
                (Std681Std_Std1 / Std68_Std1) ^ 2 + _
                (Std761Std_Std1 / Std76_Std1) ^ 2)
                
                If Err.Number <> 0 Then
                    Std751Std_Std1 = 0
                End If
            
            Err.Clear
                        
            'Sdt2---------------------------------------------------------------------------
                                        
            Err.Clear
            
                Std761Std_Std2 = _
                    (Std76_Std2 ^ (1 - Teta)) * _
                    (1 - Teta) * _
                    ((SlpStdBlkCorr_Sh.Range(Column761Std & Std2) / (Std76_Std2)))

'                Std761Std_Std2 = SlpStdBlkCorr_Sh.Range(Column761Std & Std2) / (Std76_Std2)
                
                If Err.Number <> 0 Then
                    Std761Std_Std2 = 0
                End If
            
            
            Err.Clear
            
                Std681Std_Std2 = _
                    (Std68_Std2 ^ (1 - Teta)) * _
                    (1 - Teta) * _
                    ((SlpStdBlkCorr_Sh.Range(Column681Std & Std2) / (Std68_Std2)))

'                Std681Std_Std2 = SlpStdBlkCorr_Sh.Range(Column681Std & Std2) / (Std68_Std2)
            
                If Err.Number <> 0 Then
                    Std681Std_Std2 = 0
                End If
                        
            Err.Clear
                
                Std751Std_Std2 = Std75_Std2 * Sqr( _
                (Std681Std_Std2 / Std68_Std2) ^ 2 + _
                (Std761Std_Std2 / Std76_Std2) ^ 2)
                
                If Err.Number <> 0 Then
                    Std751Std_Std2 = 0
                End If
            
            Err.Clear
            
            On Error GoTo 0

        Case False
                        
            Std681Std_Std1 = 0
            'Std68_Std1 = 1
            Std761Std_Std1 = 0
            'Std76_Std1 = 1
            Std751Std_Std1 = 0
            'Std75_Std1 = 1
            
            Std681Std_Std2 = 0
            'Std68_Std2 = 1
            Std761Std_Std2 = 0
            'Std76_Std2 = 1
            Std751Std_Std2 = 0
            'Std75_Std2 = 1

    End Select

    With SlpStdCorr_Sh
    
        'Ratio 68 corrected by standard
        Set P = SlpStdBlkCorr_Sh.Range(Column68 & Slp)
        
            Set P2 = .Range(StdCorr_Column68 & C)
        
        On Error Resume Next
            P2 = P * ExtStd68 / _
                (((Std68_Std1 ^ (1 - Teta)) * _
                    (Std68_Std2 ^ Teta)))
        
            If Err.Number <> 0 Then
                P2 = "n.a."
            End If
        
                Err.Clear
        
        On Error GoTo 0
        
'        SlpStdBlkCorr_Sh.Activate
'        P.Select
'        .Activate
'        P2.Select
        
        '68 error propagation
        Set Q = SlpStdBlkCorr_Sh.Range(Column681Std & Slp)

            On Error Resume Next
            .Range(StdCorr_Column681Std & C) = P2 * _
                ExtStd68Reproducibility * _
                    Sqr( _
                        (Q / P) ^ 2 + _
                        (Blk61Std_Blk1 / Blk6_Blk1) ^ 2 + _
                        (Blk61Std_Blk2 / Blk6_Blk2) ^ 2 + _
                        (Std681Std_Std1 / (Std68_Std1 ^ (1 - Teta))) ^ 2 + _
                        (Std681Std_Std2 / (Std68_Std2 ^ Teta)) ^ 2 + _
                        (ExtStd681Std / ExtStd68) ^ 2)
                         
                If Err.Number <> 0 Then
                    .Range(StdCorr_Column681Std & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
        
            
            'Application.SendKeys "^g ^a {DEL}"

        'R
        .Range(StdCorr_Column68R & C) = SlpStdBlkCorr_Sh.Range(Column68R & Slp)
                
        'R2
        .Range(StdCorr_Column68R2 & C) = SlpStdBlkCorr_Sh.Range(Column68R2 & Slp)
                                
        'Ratio 76 corrected by standard
        Set P = SlpStdBlkCorr_Sh.Range(Column76 & Slp)
            
        Set P2 = .Range(StdCorr_Column76 & C)
            
            On Error Resume Next
            P2 = P * (ExtStd76 / _
                ((Std76_Std1 ^ (1 - Teta)) * _
                    (Std76_Std2 ^ Teta)))
                    
                If Err.Number <> 0 Then
                    P2 = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
                                 
        '76 error propagation
        Set Q = SlpStdBlkCorr_Sh.Range(Column761Std & Slp)
            
            On Error Resume Next
            .Range(StdCorr_Column761Std & C) = P2 * _
                ExtStd76Reproducibility * _
                    Sqr( _
                        (Q / P) ^ 2 + _
                        (Blk61Std_Blk1 / Blk6_Blk1) ^ 2 + _
                        (Blk61Std_Blk2 / Blk6_Blk2) ^ 2 + _
                        (Blk71Std_Blk1 / Blk7_Blk1) ^ 2 + _
                        (Blk71Std_Blk2 / Blk7_Blk2) ^ 2 + _
                        (Std761Std_Std1 / (Std76_Std1 ^ (1 - Teta))) ^ 2 + _
                        (Std761Std_Std2 / (Std76_Std2 ^ Teta)) ^ 2 + _
                        (ExtStd761Std / ExtStd76) ^ 2)
            
            ''''''''''''''''''''''''''''''''''''''''''
            'Debug.Print "76"
            'Debug.Print Q / P
            'Debug.Print Std761Std_Std1 / Std76_Std1
            'Debug.Print Std761Std_Std2 / Std76_Std2
            'Debug.Print ExtStd761Std / ExtStd76

                If Err.Number <> 0 Then
                    .Range(StdCorr_Column761Std & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
                        
        'Ratio 75 corrected by standard
        Set P = SlpStdBlkCorr_Sh.Range(Column75 & Slp)
        
            Set P2 = .Range(StdCorr_Column75 & C)
            
            On Error Resume Next
            P2 = P * ExtStd75 / _
                ((Std75_Std1 ^ (1 - Teta)) * _
                    (Std75_Std2 ^ Teta))
                
                If Err.Number <> 0 Then
                    P2 = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
'        SlpStdBlkCorr_Sh.Activate
'        P.Select
'        .Activate
'        P2.Select
            
        '75 error propagation
        Set Q = SlpStdCorr_Sh.Range(StdCorr_Column681Std & C)
            Set QQ = SlpStdCorr_Sh.Range(StdCorr_Column761Std & C)
        
        Set P = SlpStdCorr_Sh.Range(StdCorr_Column68 & C)
            Set PP = SlpStdCorr_Sh.Range(StdCorr_Column76 & C)
            
            On Error Resume Next
            .Range(StdCorr_Column751Std & C) = P2 * _
                ExtStd75Reproducibility * _
                    Sqr( _
                        (Q / P) ^ 2 + _
                        (QQ / PP) ^ 2 + _
                        (ExtStd751Std / ExtStd75) ^ 2)
            
                If Err.Number <> 0 Then
                    .Range(StdCorr_Column751Std & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
                        
        'Rho
            On Error Resume Next
            SlpStdCorr_Sh.Range(StdCorr_Column7568Rho & C) = _
                (SlpStdCorr_Sh.Range(StdCorr_Column681Std & C) / SlpStdCorr_Sh.Range(StdCorr_Column68 & C)) / _
                (SlpStdCorr_Sh.Range(StdCorr_Column751Std & C) / SlpStdCorr_Sh.Range(StdCorr_Column75 & C))
                
                If Err.Number <> 0 Then
                    SlpStdCorr_Sh.Range(StdCorr_Column7568Rho & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0

        '202 isotope
        Set P = .Range(StdCorr_Column2 & C)
        
            Set P2 = SlpStdBlkCorr_Sh.Range(Column2 & Slp)
        
                P = P2
                        
        '202 isotope error propagation
        Set Q = SlpStdBlkCorr_Sh.Range(Column21Std & Slp)

            On Error Resume Next
            .Range(StdCorr_Column21Std & C) = P * Sqr( _
            (Q / P2) ^ 2 + _
            (Blk21Std_Blk1 / Blk2_Blk1) ^ 2 + _
            (Blk21Std_Blk2 / Blk2_Blk2) ^ 2)
            
                If Err.Number <> 0 Then
                    .Range(StdCorr_Column21Std & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
            
        '204 isotope
        Set P = .Range(StdCorr_Column4 & C)
        
            Set P2 = SlpStdBlkCorr_Sh.Range(Column4 & Slp)
            
                P = P2
        
        '204 isotope error propagation
        Set Q = SlpStdBlkCorr_Sh.Range(Column41Std & Slp)
            
            On Error Resume Next
            .Range(StdCorr_Column41Std & C) = P * Sqr( _
            (Q / P2) ^ 2 + _
            (Blk41Std_Blk1 / Blk4_Blk1) ^ 2 + _
            (Blk41Std_Blk2 / Blk4_Blk2) ^ 2)
            
                If Err.Number <> 0 Then
                    .Range(StdCorr_Column41Std & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
            
        'Ratio 64 corrected by standard
        Set P = .Range(StdCorr_Column64 & C)
        
            Set P2 = SlpStdBlkCorr_Sh.Range(Column64 & Slp)
            
                P = P2
        
        '64 error propagation
        Set Q = SlpStdBlkCorr_Sh.Range(Column641Std & Slp)
            
            On Error Resume Next
            .Range(StdCorr_Column641Std & C) = P * Sqr( _
            (Q / P2) ^ 2 + _
            (Blk61Std_Blk1 / Blk6_Blk1) ^ 2 + _
            (Blk61Std_Blk2 / Blk6_Blk2) ^ 2 + _
            (Blk41Std_Blk1 / Blk4_Blk1) ^ 2 + _
            (Blk41Std_Blk2 / Blk4_Blk2) ^ 2)
            
                If Err.Number <> 0 Then
                    .Range(StdCorr_Column641Std & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
            
        'Ratio 74 corrected by standard
        Set P = .Range(StdCorr_Column74 & C)
            
            Set P2 = SlpStdBlkCorr_Sh.Range(Column74 & Slp)
        
                P = P2
                
        '74 error propagation
        Set Q = SlpStdBlkCorr_Sh.Range(Column741Std & Slp)
            
            On Error Resume Next
            .Range(StdCorr_Column741Std & C) = P * Sqr( _
            (Q / P2) ^ 2 + _
            (Blk71Std_Blk1 / Blk7_Blk1) ^ 2 + _
            (Blk71Std_Blk2 / Blk7_Blk2) ^ 2 + _
            (Blk41Std_Blk1 / Blk4_Blk1) ^ 2 + _
            (Blk41Std_Blk2 / Blk4_Blk2) ^ 2)
            
                If Err.Number <> 0 Then
                    .Range(StdCorr_Column741Std & C) = "n.a."
                End If
                
                    Err.Clear
                    
            On Error GoTo 0
            
        'Fraction of common 206Pb
        
        
        'Ratio 28
        If Isotope232Analyzed_UPb = True Then
            Set P = .Range(StdCorr_Column28 & C)
            
                Set P2 = SlpStdBlkCorr_Sh.Range(Column28 & Slp)
            
                    P = P2
                    
            '28 error propagation
            Set Q = .Range(StdCorr_Column281Std & C)
            
                Q = SlpStdBlkCorr_Sh.Range(Column281Std & Slp)
        End If
            
        'The code below used to depends on Isoplot 4.15 to calculate 6/4 from stacey & Krammers single stage model and
        'the ages in Ma based on 68, 75 and 76 ratios. However, these function were implemented in Chronus
        
        '206* (%) - Common 206Pb based on Stacey & Kramers single stage model implemented by Ludwig
        Set P = .Range(StdCorr_ColumnF206 & C)
        Set Q = .Range(StdCorr_Column64 & C)
        
            If Q = "n.a." Then
                P = "n.a."
            Else
                On Error Resume Next
                    P = 100 * (Chronus_SingleStagePbR(Chronus_AgePb6U8(.Range(StdCorr_Column68 & C)), 1) / Q)
                        If Err.Number <> 0 Then
                            P = "n.a."
                        End If
                On Error GoTo 0
            End If
                
        'Age 68
        Set P = .Range(StdCorr_Column68AgeMa & C)
        Set Q = .Range(StdCorr_Column68 & C)

            If Q = "n.a." Then
                P = "n.a."
            Else
                On Error Resume Next
'                        P = AgePb6U8(Q)
                    P = Chronus_AgePb6U8(Q.Value)
                        If Err.Number <> 0 Then
                            P = "n.a."
                        End If
                On Error GoTo 0
            End If

        'Age 68 1 std
        Set P = .Range(StdCorr_Column68AgeMa1std & C)
        Set Q = .Range(StdCorr_Column681Std & C)
        Set P2 = .Range(StdCorr_Column68 & C)

            If Q = "n.a." Then
                P = "n.a."
            Else
                On Error Resume Next
'                        P = AgePb6U8(P2 + Q) - AgePb6U8(P2)
                    P = Chronus_AgePb6U8(P2.Value + Q.Value) - Chronus_AgePb6U8(P2.Value)
                        If Err.Number <> 0 Then
                            P = "n.a."
                        End If
                On Error GoTo 0
            End If


        'Age 75
        Set P = .Range(StdCorr_Column75AgeMa & C)
        Set Q = .Range(StdCorr_Column75 & C)

            If Q = "n.a." Then
                P = "n.a."
            Else
                On Error Resume Next
'                        P = AgePb7U5(Q)
                    P = Chronus_AgePb7U5(Q.Value)
                        If Err.Number <> 0 Then
                            P = "n.a."
                        End If
                On Error GoTo 0
            End If

        'Age 75 1 std
        Set P = .Range(StdCorr_Column75AgeMa1std & C)
        Set Q = .Range(StdCorr_Column751Std & C)
        Set P2 = .Range(StdCorr_Column75 & C)

            If Q = "n.a." Then
                P = "n.a."
            Else
                On Error Resume Next
'                        P = AgePb7U5(P2 + Q) - AgePb7U5(P2)
                        P = Chronus_AgePb7U5(P2.Value + Q.Value) - Chronus_AgePb7U5(P2.Value)
                        If Err.Number <> 0 Then
                            P = "n.a."
                        End If
                On Error GoTo 0
            End If

        'Age 76
        Set P = .Range(StdCorr_Column76AgeMa & C)
        Set Q = .Range(StdCorr_Column76 & C)

            If Q = "n.a." Then
                P = "n.a."
            Else
                On Error Resume Next
'                        P = agepb76(Q)
                    P = Chronus_AgePb76(Q.Value)
                        If Err.Number <> 0 Then
                            P = "n.a."
                        End If
                On Error GoTo 0
            End If

        'Age 76 1 std
        Set P = .Range(StdCorr_Column76AgeMa1std & C)
        Set Q = .Range(StdCorr_Column761Std & C)
        Set P2 = .Range(StdCorr_Column76 & C)

            If Q = "n.a." Then
                P = "n.a."
            Else
                On Error Resume Next
'                        P = agepb76(P2 + Q) - agepb76(P2)
                    P = Chronus_AgePb76(P2.Value + Q.Value) - Chronus_AgePb76(P2.Value)
                        If Err.Number <> 0 Then
                            P = "n.a."
                        End If
                On Error GoTo 0
            End If
            
        '68 and 76 age concordance
        On Error Resume Next
            
            .Range(StdCorr_Column6876Conc & C) = _
                100 * _
                (1 - (.Range(StdCorr_Column68AgeMa & C) / _
                .Range(StdCorr_Column76AgeMa & C)))
                        
            If Err.Number <> 0 Then
                .Range(StdCorr_Column6876Conc & C) = "n.a."
            End If
            
            .Range(StdCorr_Column6875Conc & C) = _
                100 * _
                (1 - (.Range(StdCorr_Column68AgeMa & C) / _
                .Range(StdCorr_Column75AgeMa & C)))
                        
            If Err.Number <> 0 Then
                .Range(StdCorr_Column6875Conc & C) = "n.a."
            End If
            
        On Error GoTo 0
           
    End With
    
End Sub

Sub StandardDeviationTest(Sh As Worksheet, Test68 As Boolean, LineFit68 As Boolean, Test76 As Boolean, StdDevLimit As Integer, Optional Test28 As Boolean = False, _
Optional Test74 As Boolean = False, Optional Test64 As Boolean = False, Optional RunningAgain As Boolean = False, Optional TestingAll As Boolean = False)
    
    'This procedure takes some ranges in Plot_Sh and calculates the standard deviation of them. Then, this is used
    'to eliminate some rows with data that deviates a lot from the average.
    
    'Arguments
    'Sh is the worksheet with the data
    'Test from 68 to 64 as boolean gives the user the possibility to choose which isotopes should be tested.
    
    'Updated 28092015 - ClearRowArray now is redimensioned everytime an item is added.
    
    Dim StdDev68Ypts As Double 'Standard deviation of Yi points in relation to YiPredicted, based on a linear relationship
    Dim UpperLimit As Double 'YiPredicted + StdDev68Ypts
    Dim LowerLimit As Double 'YiPredicted - StdDev68Ypts
    Dim YiPred 'Yi predicted based on a linear relationship
    Dim x As Range 'Range of independent variable. MUST HAVE ONLY ONE AREA!
    Dim Xcouter As Integer 'Number of rows in range X
    Dim YAverage As Double 'Average of Y points
    Dim YStdDev As Double
    Dim Y68 As Range, Y76 As Range, Y28 As Range, Y74 As Range, Y64 As Range 'Ranges of dependent variables. MUST HAVE ONLY ONE AREA!
    Dim counter As Integer
    Dim CellInRange As Range
    Dim lineSlope As Double
    Dim lineIntercept As Double
    Dim IsThereEmptyElement As Boolean
    Dim StdDevTestMsg As Integer
    Dim FailedCycles As Long 'Number of cycles that failed the standard deviation test.
    Dim ScreenUpdt As Boolean 'Variable used to store application.screenupdating state
    Dim EnaEvent As Boolean 'Variable used to store application.enableevents state
    
    Dim ClearRowArray() As Variant 'Row with items that didn´t pass the standard deviation test and must be eliminated
        ReDim ClearRowArray(1 To 1) As Variant
    Dim ClearRowArray_Unique() As Variant 'The same as ClearRowArray sans duplicate values.
        ReDim ClearRowArray_Unique(1 To 1) As Variant
    
    If Sh Is Nothing Then
        If MsgBox("You must select an worksheet to run the standard deviation test. Would you like to skip the standard deviation " _
            & "test (the program will keep running)?", vbYesNo) = vbYes Then
            Exit Sub
        Else
            End
        End If
    End If
    
    If RawNumberCycles_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    Set x = Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + 1, Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb)
    Set Y68 = Sh.Range(Plot_Column68 & Plot_HeaderRow + 1, Plot_Column68 & Plot_HeaderRow + RawNumberCycles_UPb)
    Set Y76 = Sh.Range(Plot_Column76 & Plot_HeaderRow + 1, Plot_Column76 & Plot_HeaderRow + RawNumberCycles_UPb)
    Set Y28 = Sh.Range(Plot_Column28 & Plot_HeaderRow + 1, Plot_Column28 & Plot_HeaderRow + RawNumberCycles_UPb)
    Set Y74 = Sh.Range(Plot_Column74 & Plot_HeaderRow + 1, Plot_Column74 & Plot_HeaderRow + RawNumberCycles_UPb)
    Set Y64 = Sh.Range(Plot_Column64 & Plot_HeaderRow + 1, Plot_Column64 & Plot_HeaderRow + RawNumberCycles_UPb)
    
    'Standard deviation test for 68 ratios
    If Test68 = True Then
    
        If Not SpotRaster_UPb = "" Then
                    
            If LineFit68 = True And SpotRaster_UPb = "Raster" And RunningAgain = False Then
                
                If MsgBox("You set the laser analysis type as Raster, case when normally there is minimum frationation between " & _
                "238U and 206Pb. So, the program calculates the 6/8 ratio by doing a simple average and a standard deviation. " & _
                "Setting linear fit for this ratio forces the program to calculate the standard deviation test using a linear fit " & _
                "for this data. Do you still want to proceed?", vbYesNo) = vbNo Then
                
                    Box6_StdDevTest.Show
                        Call UnloadAll: End
                End If
                        
            ElseIf LineFit68 = False And SpotRaster_UPb = "Spot" And RunningAgain = False Then
            
                If MsgBox("You set the laser analysis type as Spot, case when normally there is a large frationation between " & _
                "238U and 206Pb. So, the program calculates the 6/8 ratio using a linear fit. Setting linear fit as false" & _
                " for this ratio forces the program to do the Standard Deviation Test based on a simple standard deviation " & _
                "of this data (not based on a linear fit). Do you still want to proceed?", vbYesNo) = vbNo Then
                
                    Box6_StdDevTest.Show
                        Call UnloadAll: Exit Sub
                End If
        
            End If
            
            Select Case LineFit68
                
                Case True
                'In this case, we assume that 68 fractionated during ablation and that a linear regression is a suitable
                'model to his data. So, we calculate the standard deviation of the points in relation to the line fit. In
                'or der to apply the StdDev test to the data, we must compare each data point to its predicted value based
                'on the line fit.
                
                StdDev68Ypts = LineFitStdDev(Y68, x)
                
                    For counter = 1 To Y68.count
                        
                        YiPred = LineFitYiPred(Y68, x, x.Item(counter))
                        UpperLimit = YiPred + StdDev68Ypts * StdDevLimit
                        LowerLimit = YiPred - StdDev68Ypts * StdDevLimit
                                        
                        If Not IsEmpty(Y68.Item(counter)) Then
                        
                            If Y68.Item(counter) > UpperLimit Or Y68.Item(counter) < LowerLimit Then
                            
                                ClearRowArray(UBound(ClearRowArray)) = Y68.Item(counter).Row
                                    ReDim Preserve ClearRowArray(1 To UBound(ClearRowArray) + 1)
                                    
                            End If
                        End If
                    Next
                
                Case False
                'In this case, we assume that the 238U and 206Pb were not fractionated
                    '68 average
                    YAverage = WorksheetFunction.Average(Y68)
                    
                    '68 error
                    YStdDev = WorksheetFunction.StDev_S(Y68)
                    
                    UpperLimit = YAverage + StdDevLimit * YStdDev
                    LowerLimit = YAverage - StdDevLimit * YStdDev
            
                        For counter = 1 To Y68.count
                            
                            If Not IsEmpty(Y68.Item(counter)) Then
                            
                                If Y68.Item(counter) > UpperLimit Or Y68.Item(counter) < LowerLimit Then
                                    ClearRowArray(UBound(ClearRowArray)) = Y68.Item(counter).Row
                                        ReDim Preserve ClearRowArray(1 To UBound(ClearRowArray) + 1)
                                End If
                                
                            End If
                        Next
            End Select
        End If
            
    End If
    
    'Standard deviation test for 76 ratios
    If Test76 = True Then
        
        YAverage = WorksheetFunction.Average(Y76)
        YStdDev = WorksheetFunction.StDev_S(Y76)
        UpperLimit = YAverage + StdDevLimit * YStdDev
        LowerLimit = YAverage - StdDevLimit * YStdDev
        
            For counter = 1 To Y76.count
            
                If Not IsEmpty(Y76.Item(counter)) Then
                                           
                    If Y76.Item(counter) > UpperLimit Or Y76.Item(counter) < LowerLimit Then
                        ClearRowArray(UBound(ClearRowArray)) = Y76.Item(counter).Row
                            ReDim Preserve ClearRowArray(1 To UBound(ClearRowArray) + 1)
                    End If
                            
                End If
                  
            Next
            
    End If
    
    'Standard deviation test for 28 ratios
    If Test28 = True Then
        
        YAverage = WorksheetFunction.Average(Y28)
        YStdDev = WorksheetFunction.StDev_S(Y28)
        UpperLimit = YAverage + StdDevLimit * YStdDev
        LowerLimit = YAverage - StdDevLimit * YStdDev
        
            For counter = 1 To Y28.count
            
                If Not IsEmpty(Y28.Item(counter)) Then
                                                
                    If Y28.Item(counter) > UpperLimit Or Y28.Item(counter) < LowerLimit Then
                        ClearRowArray(UBound(ClearRowArray)) = Y28.Item(counter).Row
                            ReDim Preserve ClearRowArray(1 To UBound(ClearRowArray) + 1)
                    End If
                            
                End If
                
            Next
            
    End If

    'Standard deviation test for 74 ratios
    If Test74 = True Then
        
        YAverage = WorksheetFunction.Average(Y74)
        YStdDev = WorksheetFunction.StDev_S(Y74)
        UpperLimit = YAverage + StdDevLimit * YStdDev
        LowerLimit = YAverage - StdDevLimit * YStdDev
        
            For counter = 1 To Y74.count
                
                If Not IsEmpty(Y74.Item(counter)) Then
                                                
                    If Y74.Item(counter) > UpperLimit Or Y74.Item(counter) < LowerLimit Then
                        ClearRowArray(UBound(ClearRowArray)) = Y74.Item(counter).Row
                            ReDim Preserve ClearRowArray(1 To UBound(ClearRowArray) + 1)
                    End If
                            
                 End If
                            
            Next
            
    End If

    'Standard deviation test for 64 ratios
    If Test64 = True Then
        
        YAverage = WorksheetFunction.Average(Y64)
        YStdDev = WorksheetFunction.StDev_S(Y64)
        UpperLimit = YAverage + StdDevLimit * YStdDev
        LowerLimit = YAverage - StdDevLimit * YStdDev
        
            For counter = 1 To Y64.count
            
                If Not IsEmpty(Y64.Item(counter)) Then
                                                
                    If Y64.Item(counter) > UpperLimit Or Y64.Item(counter) < LowerLimit Then
                        ClearRowArray(UBound(ClearRowArray)) = Y64.Item(counter).Row
                            ReDim Preserve ClearRowArray(1 To UBound(ClearRowArray) + 1)
                    End If
                    
                End If
                            
            Next
            
    End If
    
    'Cleaning rows that didn't pass the test
    
    ClearRowArray_Unique = Array_Unique(ClearRowArray)
    
    'Information to the user
    If IsEmpty(ClearRowArray_Unique(LBound(ClearRowArray_Unique))) = True Then 'All cells passed the test.
            If TestingAll = False Then
                MsgBox "All cycles passed the " & StdDevLimit & " standard deviation test.", , StdDevLimit & " Standard Deviation Test - " & Plot_Sh.Name
            End If
                Exit Sub
    Else
        FailedCycles = UBound(ClearRowArray_Unique) + 1
    End If
    
    If Not IsEmpty(ClearRowArray_Unique(LBound(ClearRowArray_Unique))) = True Then
        If (NumElements(ClearRowArray_Unique, 1) - 1) / (counter - 1) > 0.5 Then 'The "- 1" is related to the problem that there is always one item more in ClearRowArray than necessary
            If MsgBox("More than 50% of rows contain data that didn't pass the " & StdDevLimit & _
                "standard deviation test. Do you still want to clear these rows?", vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    End If
        
    If ClearRowArray_Unique(UBound(ClearRowArray_Unique)) = 0 Then 'Always the last element will be empty because of the redim statement.
        IsThereEmptyElement = DeleteArrayElement(ClearRowArray_Unique, UBound(ClearRowArray_Unique), True)
    End If
    
    If IsArrayAllNumeric(ClearRowArray_Unique, False) = False Then 'This is just a simple check to avoid some kind of problem.
        MsgBox "There is at least one element in ClearRowArray_Unique array that is not a number. StandardDeviationTest program will stop." & _
            " The test can't continue.", , StdDevLimit & " Standard Deviation Test - " & Plot_Sh.Name
                Exit Sub
    End If
    
    'Deleting cells
    counter = 1
        
        'Enable lines below related to screenupdating if you'd like that the user be able to see
        'outliers being deleted
'        ScreenUpdt = Application.ScreenUpdating
'            Application.ScreenUpdating = True
                For counter = LBound(ClearRowArray_Unique) To UBound(ClearRowArray_Unique)

                    EnaEvent = Application.EnableEvents
                        Application.EnableEvents = False
                    
                    Sh.Range(Plot_FirstColumn & ClearRowArray_Unique(counter), Plot_LastColumn & ClearRowArray_Unique(counter)).Clear
                
                    Application.EnableEvents = EnaEvent
                Next
                
                Call ResultsPreviewCalculation
'            Application.ScreenUpdating = ScreenUpdt
            
    If TestingAll = False Then
        StdDevTestMsg = MsgBox(FailedCycles & " cycle(s) failed the " & StdDevLimit & " standard deviation test. " & _
              "Would you like to run the test once again?", vbYesNo, StdDevLimit & " Standard Deviation Test - " & Plot_Sh.Name)
    Else
'        StdDevTestMsg = MsgBox(FailedCycles & " cycle(s) failed the " & StdDevLimit & " standard deviation test. " & _
'              "Would you like to run the test once again?", vbYesNoCancel, StdDevLimit & " Standard Deviation Test - " & Plot_Sh.Name)
    End If
    
        If StdDevTestMsg = vbYes Then
            Call StandardDeviationTest(Plot_Sh, Test68, LineFit68, Test76, StdDevLimit, Test28, Test74, Test64, True)
        ElseIf StdDevTestMsg = vbCancel Then
            Call UnloadAll
                End
        End If
   
    End Sub

Sub ExternalReproSamples()

    'This program will look for external standard ID in SlpStdBlkCorr_Sh, copy all their 68, 76 and 75
    'ratios and calculate a weighted average, uncertainties and MSWD of these ratios using WtdAv Isoplot
    'built-in function. This will be stored above the first row with data from samples and standards
    'in SlpStdBlkCorr_Sh.
    
    Dim ExtStd() As Integer 'Array with external standards IDs only
    Dim a As Variant
    Dim C As Long
    Dim d As Long
    Dim f As Variant
    Dim counter As Long
    Dim IDsRange As Range 'Range with IDs of all samples and standards (internal and external) are in SlpStdBlkCorr_Sh
    Dim LastRow As Integer 'Last row of IDsRange
    Dim E As Long
    Dim StdName As String
    Dim AnalysesListNumber As Long
    Dim Blk1Row As Long
    Dim ExtStd68 As Double 'ExtStd68 ratio
    Dim ExtStd681Std As Double 'ExtStd681Std
    Dim ExtStd75 As Double 'ExtStd75 ratio
    Dim ExtStd751Std As Double 'ExtStd751Std
    Dim ExtStd76 As Double 'ExtStd76 ratio
    Dim ExtStd761Std As Double 'ExtStd761Std
    
    'Ranges of data only from external standard, already copied to the new range
    Dim Ratio68Range As Range
    Dim Ratio75Range As Range
    Dim Ratio76Range As Range
    
    Dim FirstNewRow As Long 'First row where data from standards will be copied to
    
    'Columns where data from standard will be copied to
    Dim Ratio68ColumnNew As String
    Dim Ratio75ColumnNew As String
    Dim Ratio76ColumnNew As String
    Dim Ratio681StdColumnNew As String
    Dim Ratio751StdColumnNew As String
    Dim Ratio761StdColumnNew As String
    Dim Std681Std As Range '68 1 std for Std1
    Dim Std68 As Range '68 for Std1
    Dim Std761Std As Range '76 1 std for Std1
    Dim Std76 As Range '76 for Std1
    Dim Std751Std As Range '75 1 std for Std1
    Dim Std75 As Range '75 for Std1
    Dim Blk61Std As Double '206 1 std for Blk1
    Dim Blk6 As Double '206 for Blk1
    Dim Blk71Std As Double '207 1 std for Blk1
    Dim Blk7 As Double '207 for Blk1
 
    If SlpStdBlkCorr_Sh Is Nothing Then
        Call PublicVariables
    End If

    If IsArrayEmpty(StdFound) = True Then
        Call IdentifyFileType
    End If
    
    On Error Resume Next
        If AnalysesList_std(0).Std = "" Then
            Call LoadStdListMap
        End If
    On Error GoTo 0

    ReDim ExtStd(1 To UBound(StdFound) + 1) As Integer
    
    C = 2

        For Each a In StdFound 'External standards IDs are copied to a different array (SlpStd) which accepts only numbers (IDs)
            ExtStd(C - 1) = SamList_Sh.Range(a).Offset(, 1)
            C = C + 1
        Next
        
    counter = 1
    
    'The following 6 lines are used just to check if UPbStd was initialized
    On Error Resume Next
        counter = LBound(UPbStd)
            If Err.Number <> 0 Then
                Call Load_UPbStandardsTypeList
            End If
    On Error GoTo 0
    
    'The 6 lines below are necessary to adentify the number of the external standard in UpbStd
    For counter = LBound(UPbStd) To UBound(UPbStd)
        If UPbStd(counter).StandardName = ExternalStandard_UPb Then
            StdName = counter
                counter = UBound(UPbStd)
        End If
    Next

    Ratio68ColumnNew = "A"
    Ratio681StdColumnNew = "B"
    Ratio75ColumnNew = "C"
    Ratio751StdColumnNew = "D"
    Ratio76ColumnNew = "E"
    Ratio761StdColumnNew = "F"
    
    'Certified primary standard ratios and uncertainties
    ExtStd68 = UPbStd(StdName).Ratio68
    ExtStd75 = UPbStd(StdName).Ratio75
    ExtStd76 = UPbStd(StdName).Ratio76
        ExtStd681Std = UPbStd(StdName).Ratio68Error
        ExtStd751Std = UPbStd(StdName).Ratio75Error
        ExtStd761Std = UPbStd(StdName).Ratio76Error
    
    FirstNewRow = 2
        
    With SlpStdBlkCorr_Sh
        
        LastRow = .Range(ColumnID & HeaderRow + 1).End(xlDown).Row
        
        Set IDsRange = .Range(.Range(ColumnID & HeaderRow + 1), .Range(ColumnID & HeaderRow + 1).End(xlDown))
        
        C = 1
        d = FirstNewRow
        
            ExtStd68Repro.ClearContents
            ExtStd75Repro.ClearContents
            ExtStd76Repro.ClearContents
                
                ExtStd68Repro1std.ClearContents
                ExtStd75Repro1std.ClearContents
                ExtStd76Repro1std.ClearContents
                
        For Each a In ExtStd

            'For each structure used to find the external standard inside AnalysesList_std
            For E = 1 To UBound(AnalysesList_std)
                If a = AnalysesList_std(E).Std Then 'There is a problem here because the blank for standard can be changed, it's
                                                   'not necessarly the same as the sample
                    AnalysesListNumber = E 'Using this variable I am able to retrieve from AnalysesList all the IDs that I must know
                    E = UBound(AnalysesList_std) 'A beautiful solution to end the if structure
                End If
            Next
            
            'The code inside the with block below is used to find the row of the blank used to correct the primary standard
            'being processed
            With BlkCalc_Sh
                
                If .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown) = "" Then
                    MsgBox ("You need at least two blanks to reduce your data.")
                        Application.GoTo BlkCalc_Sh.Range(BlkColumnID & BlkCalc_HeaderLine)
                            End
                End If
                
                For Each f In .Range(BlkColumnID & BlkCalc_HeaderLine + 1, .Range(BlkColumnID & BlkCalc_HeaderLine + 1).End(xlDown))
                    
                    If AnalysesList_std(AnalysesListNumber).Blk1 = f Then
                        Blk1Row = f.Row
                    End If

                Next
            
            End With

             
             For C = 1 To WorksheetFunction.count(IDsRange)
                
                If a = IDsRange.Item(C) Then

                    Set Std68 = .Range(Ratio68ColumnNew & LastRow + d)
                    Set Std681Std = .Range(Ratio681StdColumnNew & LastRow + d)
                    Set Std75 = .Range(Ratio75ColumnNew & LastRow + d)
                    Set Std751Std = .Range(Ratio751StdColumnNew & LastRow + d)
                    Set Std76 = .Range(Ratio76ColumnNew & LastRow + d)
                    Set Std761Std = .Range(Ratio761StdColumnNew & LastRow + d)
                    
                        Std68 = .Range(Column68 & IDsRange.Item(C).Row)
                        Std75 = .Range(Column75 & IDsRange.Item(C).Row)
                        Std76 = .Range(Column76 & IDsRange.Item(C).Row)

                    'Primary standard 68 uncertainty evaluation
                    With BlkCalc_Sh 'REPETION OF THE LINES BELOW
                         Blk6 = .Range(BlkColumn6 & Blk1Row)
                         Blk61Std = .Range(BlkColumn61Std & Blk1Row)
                         Blk7 = .Range(BlkColumn7 & Blk1Row)
                         Blk71Std = .Range(BlkColumn71Std & Blk1Row)
                    End With
            
                        If .Range(Column6 & IDsRange.Item(C).Row) / Abs(Blk6) > CutOffRatio Or Blk6 < 0 Then
                            Blk61Std = 0
                        End If
                                                    
                            Std681Std = Std68 * Sqr( _
                                (.Range(Column681Std & IDsRange.Item(C).Row) / Std68) ^ 2 + _
                                (Blk61Std / Blk6) ^ 2 + _
                                (ExtStd681Std / ExtStd68) ^ 2)
                                                    
                    'Primary standard 76 uncertainty evaluation
                    With BlkCalc_Sh 'REPETION OF THE LINES BELOW - UPDATE
                         Blk6 = .Range(BlkColumn6 & Blk1Row)
                         Blk61Std = .Range(BlkColumn61Std & Blk1Row)
                         Blk7 = .Range(BlkColumn7 & Blk1Row)
                         Blk71Std = .Range(BlkColumn71Std & Blk1Row)
                    End With
                                                    
                        If .Range(Column6 & IDsRange.Item(C).Row) / Abs(Blk6) > CutOffRatio Or Blk6 < 0 Then
                            Blk61Std = 0
                        End If
                            
                        If .Range(Column7 & IDsRange.Item(C).Row) / Abs(Blk7) > CutOffRatio Or Blk7 < 0 Then
                            Blk71Std = 0
                        End If
                            
                        Std761Std = Std76 * Sqr( _
                            (.Range(Column761Std & IDsRange.Item(C).Row) / Std76) ^ 2 + _
                            (Blk61Std / Blk6) ^ 2 + _
                            (Blk71Std / Blk7) ^ 2 + _
                            (ExtStd761Std / ExtStd76) ^ 2)
                        
                    'Primary standard 75 uncertainty evaluation
                        Std751Std = Std75 * Sqr( _
                            (.Range(Column751Std & IDsRange.Item(C).Row) / Std75) ^ 2 + _
                            (Std681Std / Std68) ^ 2 + _
                            (Std761Std / Std76) ^ 2 + _
                            (ExtStd751Std / ExtStd75) ^ 2)
                        
                    
                    C = WorksheetFunction.count(IDsRange)
                
                End If
                                            
            Next
            
            d = d + 1
        
        Next
        
        Set Ratio68Range = .Range(.Range(Ratio68ColumnNew & LastRow + FirstNewRow), .Range(Ratio681StdColumnNew & LastRow + d - 1))
        Set Ratio75Range = .Range(.Range(Ratio75ColumnNew & LastRow + FirstNewRow), .Range(Ratio751StdColumnNew & LastRow + d - 1))
        Set Ratio76Range = .Range(.Range(Ratio76ColumnNew & LastRow + FirstNewRow), .Range(Ratio761StdColumnNew & LastRow + d - 1))
                    
'            ExtStd68Repro = WorksheetFunction.Average(Ratio68Range)
'            ExtStd75Repro = WorksheetFunction.Average(Ratio75Range)
'            ExtStd76Repro = WorksheetFunction.Average(Ratio76Range)
            
            ExtStd68Repro.FormulaArray = "=wtdav(" & Ratio68Range.Address(False, False) & ",FALSE, FALSE,1,TRUE,FALSE,1)"
                ExtStd68Repro.Copy: ExtStd68Repro.PasteSpecial xlPasteValues
            ExtStd75Repro.FormulaArray = "=wtdav(" & Ratio75Range.Address(False, False) & ",FALSE, FALSE,1,TRUE,FALSE,1)"
                ExtStd75Repro.Copy: ExtStd75Repro.PasteSpecial xlPasteValues
            ExtStd76Repro.FormulaArray = "=wtdav(" & Ratio76Range.Address(False, False) & ",FALSE, FALSE,1,TRUE,FALSE,1)"
                ExtStd76Repro.Copy: ExtStd76Repro.PasteSpecial xlPasteValues
                
            Application.CutCopyMode = False

            'Below, standard deviation of sample will be multiplied by student's t factor
            'depending on the number of sample analuisis (degrees of freedom)
            
'            ExtStd68Repro1std = WorksheetFunction.StDev_S(Ratio68Range) * _
'                                WorksheetFunction.T_Inv_2T(ConfLevel, NumElements(StdFound))
'            ExtStd75Repro1std = WorksheetFunction.StDev_S(Ratio75Range) * _
'                                WorksheetFunction.T_Inv_2T(ConfLevel, NumElements(StdFound))
'            ExtStd76Repro1std = WorksheetFunction.StDev_S(Ratio76Range) * _
'                                WorksheetFunction.T_Inv_2T(ConfLevel, NumElements(StdFound))
'
'            ExtStd68Repro1std.Offset(1) = 100 * ExtStd68Repro1std / ExtStd68Repro
'            ExtStd75Repro1std.Offset(1) = 100 * ExtStd75Repro1std / ExtStd75Repro
'            ExtStd76Repro1std.Offset(1) = 100 * ExtStd76Repro1std / ExtStd76Repro
     
                Ratio68Range.ClearContents
                Ratio75Range.ClearContents
                Ratio76Range.ClearContents
    
    End With
    
End Sub

