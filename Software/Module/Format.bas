Attribute VB_Name = "Format"
Option Explicit
Sub FormatMainSh()
    
    'Updated 06112015
    
    Dim SearchStr As Long

    If SlpStdCorr_Sh Is Nothing Then
        Call PublicVariables
    End If
    
    SearchStr = InStr(SlpStdCorr_Sh.Range(StdCorr_Column681Std & HeaderRow), "%")
        
    Call FormatSamList
    Call FormatStartANDOptions
    Call FormatFinalReport
    
    If SearchStr = 0 Then
        Call FormatBlkCalc(True)
        Call FormatSlpStdBlkCorr(True)
        Call FormatSlpStdCorr(True)
    Else
        Call FormatBlkCalc(False)
        Call FormatSlpStdBlkCorr(False)
        Call FormatSlpStdCorr(False)
    End If
    
End Sub

Sub FormatSamList()
    
    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If

    With SamList_Sh 'The lines below are going to format cells.
    
        .Range(SamList_FilePath & SamList_HeadersLine2) = "Addresses of samples raw data "
        .Range(SamList_FileName & SamList_HeadersLine2) = "File Name"
        .Range(SamList_ID & SamList_HeadersLine2) = "ID"
        .Range(SamList_FirstCycleTime & SamList_HeadersLine2) = "First Cycle Time"
        .Range(SamList_Cycles & SamList_HeadersLine2) = "Cycles"
        .Range(SamList_StdID & SamList_HeadersLine1, SamList_BlkID & SamList_HeadersLine1).Merge
            .Range(SamList_StdID & SamList_HeadersLine1, SamList_BlkID & SamList_HeadersLine1) = "Std map"
                .Range(SamList_StdID & SamList_HeadersLine2) = "Standard"
                .Range(SamList_BlkID & SamList_HeadersLine2) = "Blank"
        .Range(SamList_SlpID & SamList_HeadersLine1, SamList_Blk2ID & SamList_HeadersLine1).Merge
        .Range(SamList_SlpID & SamList_HeadersLine1, SamList_Blk2ID & SamList_HeadersLine1).HorizontalAlignment = xlCenter
            .Range(SamList_SlpID & SamList_HeadersLine1, SamList_Blk2ID & SamList_HeadersLine1) = "Slp map"
                .Range(SamList_SlpID & SamList_HeadersLine2) = "Sample"
                .Range(SamList_Std1ID & SamList_HeadersLine2) = "Standard 1"
                .Range(SamList_Std2ID & SamList_HeadersLine2) = "Standard 2"
                .Range(SamList_Blk1ID & SamList_HeadersLine2) = "Blank 1"
                .Range(SamList_Blk2ID & SamList_HeadersLine2) = "Blank 2"
                                
        .Range(SamList_Std2ID & SamList_HeadersLine1, .Range(SamList_FilePath & SamList_HeadersLine2).End(xlDown)).Columns.AutoFit
        .Range(SamList_Cycles & SamList_HeadersLine2).Columns.ColumnWidth = 16
        .Range(SamList_FilePath & SamList_HeadersLine2).ColumnWidth = 60.86
        .Range(SamList_FilePath & ":" & SamList_Std2ID).Font.Strikethrough = False
        .Range(SamList_FirstCycleTime & SamList_HeadersLine2, .Range(SamList_FirstCycleTime & SamList_HeadersLine2).End(xlDown)).NumberFormat = "dd/mm/yyyy hh:mm:ss.000"
        .Columns(5).ColumnWidth = 14
        
        With .Range(SamList_FilePath & SamList_FirstLine, .Range(SamList_Blk2ID & SamList_FirstLine).End(xlDown))
            .Font.Italic = False
            .Font.Bold = False
            With .Font
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
        End With
        
        .Range(SamList_ID & SamList_HeadersLine1, .Range(SamList_FilePath & SamList_FirstLine).End(xlDown)).Interior.ThemeColor = xlThemeColorAccent5  'Blue color for column A to C in SamList_Sh
        .Range(SamList_Cycles & SamList_HeadersLine1, .Range(SamList_FirstCycleTime & SamList_FirstLine).End(xlDown)).Interior.ColorIndex = 6  'Yellow color for column D to E in SamList_Sh
        .Range(SamList_StdID & SamList_HeadersLine1, .Range(SamList_BlkID & SamList_FirstLine).End(xlDown)).Interior.ColorIndex = 3  'Red color for columns F and G in SamList_Sh
        .Range(SamList_SlpID & SamList_HeadersLine1, .Range(SamList_Blk2ID & SamList_FirstLine).End(xlDown)).Interior.ColorIndex = 4  'Red color for columns H to K in SamList_Sh
    
        .Range(SamList_Cycles & SamList_FirstLine, .Range(SamList_Cycles & SamList_FirstLine).End(xlDown)).HorizontalAlignment = xlRight
        .Range(SamList_FilePath & SamList_FirstLine, .Range(SamList_FilePath & SamList_FirstLine).End(xlDown)).HorizontalAlignment = xlLeft
        .Range(SamList_FileName & ":" & SamList_FirstCycleTime).HorizontalAlignment = xlCenter
        .Range(SamList_StdID & ":" & SamList_Blk2ID).HorizontalAlignment = xlCenter
        .Range(SamList_FilePath & SamList_HeadersLine1, SamList_Blk2ID & SamList_HeadersLine2).Font.Bold = True
        .Range(SamList_FilePath & SamList_HeadersLine1, SamList_Blk2ID & SamList_HeadersLine2).HorizontalAlignment = xlCenter

        Application.GoTo .Range("A" & SamList_FirstLine)
            
            With ActiveWindow
                .SplitColumn = 0
                .SplitRow = SamList_HeadersLine2
                .FreezePanes = True
            End With

    End With

End Sub

Sub FormatPlot(TargetSh As Worksheet)

    Dim RangeUnion As Range

    If mwbk Is Nothing Then
        Call PublicVariables
    End If

    Application.GoTo TargetSh.Range("A1")
    
    'Code to set the ranges for the isotopes signal in the sheet where they will be plotted
    With TargetSh
        .Range(Plot_IDCell).Offset(, -1) = "ID"
        .Range(Plot_ColumnCyclesTime & Plot_HeaderRow) = "Cycles time"
        .Range(Plot_Column75 & Plot_HeaderRow) = "207/235"
        .Range(Plot_Column68 & Plot_HeaderRow) = "206/238"
        .Range(Plot_Column76 & Plot_HeaderRow) = "207/206"
        .Range(Plot_Column2 & Plot_HeaderRow) = "Hg202 cps"
        .Range(Plot_Column4 & Plot_HeaderRow) = "204 cps"
        .Range(Plot_Column6 & Plot_HeaderRow) = "Pb206 cps"
        .Range(Plot_Column7 & Plot_HeaderRow) = "Pb207 cps"
        .Range(Plot_Column8 & Plot_HeaderRow) = "Pb208 cps"
        .Range(Plot_Column32 & Plot_HeaderRow) = "Th232 cps"
        .Range(Plot_Column38 & Plot_HeaderRow) = "U238 cps"
        .Range(Plot_Column64 & Plot_HeaderRow) = "206/204"
        .Range(Plot_Column74 & Plot_HeaderRow) = "207/204"
        .Range(Plot_Column28 & Plot_HeaderRow) = "232/238"
        
        'The following lines must be changed if the Plot_ResultsPreview constant be changed
        .Range("S4").FormulaR1C1 = "68"
        .Range("S5").FormulaR1C1 = "76"
        .Range("T3").FormulaR1C1 = "Ratios"
        .Range("U3").FormulaR1C1 = "1s (abs)"
        .Range("V3").FormulaR1C1 = "1s (%)"
        .Range("W3").FormulaR1C1 = "R"
        .Range("X3").FormulaR1C1 = "R2"
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                
        .Range(Plot_ColumnCyclesTime & Plot_HeaderRow, .Range(Plot_ColumnCyclesTime & Plot_HeaderRow).End(xlToRight)).Font.Bold = True
        .Range(Plot_IDCell).Offset(, -1).Font.Bold = True
        .Range(Plot_AnalysisName).Font.Bold = True
        
        'Number format in Plot_Sh
        .Range(.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + 1), _
            .Range(Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb)).NumberFormat = "0.000"
        
        .Range(.Range(Plot_Column2 & Plot_HeaderRow + 1), _
            .Range(Plot_Column38 & Plot_HeaderRow + RawNumberCycles_UPb)).NumberFormat = "0.0"
        
        .Range(.Range(Plot_Column64 & Plot_HeaderRow + 1), _
            .Range(Plot_Column76 & Plot_HeaderRow + RawNumberCycles_UPb)).NumberFormat = "0.000"
        
        .Cells.Columns.AutoFit
        .Cells.HorizontalAlignment = xlCenter
        
        With .Range(Plot_ResultsPreview)
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        
    
            With .Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
        End With
        
        'The lines below will format the results preview box. They must be changed if the Plot_ResultsPreview constant be changed!
        .Range("S4").Font.Bold = True
        .Range("S5").Font.Bold = True
        .Range("T3").Font.Bold = True
        .Range("U3").Font.Bold = True
        .Range("V3").Font.Bold = True
        .Range("W3").Font.Bold = True
        .Range("X3").Font.Bold = True
        
        Set RangeUnion = Application.Union(.Range("T4"), .Range("T5"), .Range("W4"), .Range("W5"), .Range("X4"), .Range("X5"))
            RangeUnion.NumberFormat = "0.000"
        Set RangeUnion = Application.Union(.Range("U4"), .Range("U5"))
            RangeUnion.NumberFormat = "0.00000"
        Set RangeUnion = Application.Union(.Range("V4"), .Range("V5"))
            RangeUnion.NumberFormat = "0.00"
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        .Range(Plot_ResultsPreview).Columns.AutoFit
        
    End With
            
            With ActiveWindow
                .SplitColumn = 1
                .SplitRow = Plot_HeaderRow
                .FreezePanes = True
            End With

End Sub

Sub FormatBlkCalc(Optional AbsoluteUncertainty As Boolean = True)

    'Procedure that formats the BlkCalc sheet.
    
    'pdated 26092015 - Autofilter arrows are always shown

    Dim RangeUnion As Range
    Dim UncertantyText As String
    Dim SearchStr As Variant
    Dim ScreenUpdt As Boolean

    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If
    
    If AbsoluteUncertainty = True Then
        UncertantyText = "abs"
    ElseIf AbsoluteUncertainty = False Then
        UncertantyText = "%"
    End If

    With BlkCalc_Sh 'The lines below are going to write headers and format cells.
    
'    'Checking if uncertainties in SlpStdBlkCorr sheet agree with AbsoluteUncertainty option
'    SearchStr = InStr(.Range(BlkColumn21Std & HeaderRow), "%")
'        If SearchStr = 0 And UncertantyText = "abs" Then
'
'            Application.ScreenUpdating = True: .Activate: Application.ScreenUpdating = ScreenUpdt
'
'                If MsgBox("Are uncertainties in this sheet percentage?", vbYesNo) = vbNo Then
'                    UncertantyText = "abs"
'                End If
'        End If

        .Range(BlkColumnID & BlkCalc_HeaderLine) = "ID"
        .Range(BlkSlpName & BlkCalc_HeaderLine) = "Sample Name"
        .Range(BlkColumn2 & BlkCalc_HeaderLine) = "202 cps"
        .Range(BlkColumn21Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        .Range(BlkColumn4 & BlkCalc_HeaderLine) = "204 cps"
        .Range(BlkColumn41Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        .Range(BlkColumn6 & BlkCalc_HeaderLine) = "206 cps"
        .Range(BlkColumn61Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        .Range(BlkColumn7 & BlkCalc_HeaderLine) = "207 cps"
        .Range(BlkColumn71Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        .Range(BlkColumn8 & BlkCalc_HeaderLine) = "208 cps"
        .Range(BlkColumn81Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        .Range(BlkColumn32 & BlkCalc_HeaderLine) = "232 cps"
        .Range(BlkColumn321Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        .Range(BlkColumn38 & BlkCalc_HeaderLine) = "238 cps"
        .Range(BlkColumn381Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        .Range(BlkColumn4Comm & BlkCalc_HeaderLine) = "204* cps"
        .Range(BlkColumn4Comm1Std & BlkCalc_HeaderLine) = "1 std (" & UncertantyText & ")"
        
        .Range(BlkColumnID & BlkCalc_HeaderLine, BlkColumn4Comm1Std & BlkCalc_HeaderLine).Font.Bold = True
        
        .Range(BlkColumnID & BlkCalc_HeaderLine + 1, .Range(BlkColumn4Comm1Std & BlkCalc_HeaderLine).End(xlDown)).NumberFormat = "0"
        .Range(BlkColumn2 & BlkCalc_HeaderLine + 1, .Range(BlkColumn4Comm1Std & BlkCalc_HeaderLine).End(xlDown)).NumberFormat = "0.00"
        
        'Lines to display the autofilter arrows
        With .Range(BlkColumnID & BlkCalc_HeaderLine, BlkColumn4Comm1Std & BlkCalc_HeaderLine)
            
            If BlkCalc_Sh.AutoFilterMode = False Then
                .AutoFilter
            End If
            
        End With

        .Cells.Columns.AutoFit
        .Cells.HorizontalAlignment = xlCenter

        Application.GoTo .Range("A" & BlkCalc_HeaderLine + 1)
            
            With ActiveWindow
                .SplitColumn = 1
                .SplitRow = BlkCalc_HeaderLine
                .FreezePanes = True
            End With
        
    End With

End Sub

Sub FormatFinalReport()

    'Procedure that formats the finalreport sheet.
    
    'Updated 26092015 - Chronus info formatting updated
    
    Dim RangeUnion As Range
    Dim SearchStr As Variant
    Dim UncertantyText As String
    Dim AutoFitRange As Range
    
    If mwbk Is Nothing Then
        Call PublicVariables
    End If

    If FinalReport_Sh Is Nothing Then
        On Error Resume Next
            Set FinalReport_Sh = mwbk.Sheets(FinalReport_Sh_Name)
            
                If Err.Number <> 0 Then
                    Exit Sub
                End If
        
        On Error GoTo 0
    End If
    
    With FinalReport_Sh
        
        .Range(FR_ChronusVersion) = .Range(FR_ChronusVersion).Value & " " & ChronusVersion
        
        'Ages in Ma
        Set RangeUnion = Application.Union( _
            .Range(FR_ColumnAge76 & ":" & FR_ColumnAge76), _
            .Range(FR_ColumnAge76 & ":" & FR_ColumnAge76), _
            .Range(FR_ColumnAge762StdAbs & ":" & FR_ColumnAge762StdAbs), _
            .Range(FR_ColumnAge68 & ":" & FR_ColumnAge68), _
            .Range(FR_ColumnAge682StdAbs & ":" & FR_ColumnAge682StdAbs), _
            .Range(FR_ColumnAge75 & ":" & FR_ColumnAge75), _
            .Range(FR_ColumnAge752StdAbs & ":" & FR_ColumnAge752StdAbs), _
            .Range(FR_ColumnAge208232 & ":" & FR_ColumnAge208232), _
            .Range(FR_ColumnAge2082322StdAbs & ":" & FR_ColumnAge2082322StdAbs))
        
            RangeUnion.NumberFormat = "0"
        
        'Uncertainties in percentage
        Set RangeUnion = Application.Union( _
            .Range(FR_Column641Std & ":" & FR_Column641Std), _
            .Range(FR_ColumnTera2382061Std & ":" & FR_ColumnTera2382061Std), _
            .Range(FR_ColumnTera761Std & ":" & FR_ColumnTera761Std), _
            .Range(FR_Column2082061Std & ":" & FR_Column2082061Std), _
            .Range(FR_ColumnWeth751Std & ":" & FR_ColumnWeth751Std), _
            .Range(FR_ColumnWeth681Std & ":" & FR_ColumnWeth681Std), _
            .Range(FR_Column2082321Std & ":" & FR_Column2082321Std), _
            .Range(FR_Column6876DiscPercent & ":" & FR_Column6876DiscPercent))
            
            RangeUnion.NumberFormat = "0.00"
        '204Pb cps
        .Range(FR_Column204PbCps & ":" & FR_Column204PbCps).NumberFormat = "0"
        
        '206 CPS
        .Range(FR_Column206PbmV & ":" & FR_Column206PbmV).NumberFormat = "0.0000"
        
        'ThU
        .Range(FR_ColumnThU & ":" & FR_ColumnThU).NumberFormat = "0.000"
        
        '238206 ratio
        .Range(FR_ColumnTera238206 & ":" & FR_ColumnTera238206).NumberFormat = "0.0"
        
        '208206 ratio
        .Range(FR_ColumnTera208206 & ":" & FR_ColumnTera208206).NumberFormat = "0.0"

        '75 wetherill ratio
        .Range(FR_ColumnWeth75 & ":" & FR_ColumnWeth75).NumberFormat = "0.000"

        '68 wetherill ratio
        .Range(FR_ColumnWeth68 & ":" & FR_ColumnWeth68).NumberFormat = "0.0000"

        'Rho
        .Range(FR_ColumnWethRho & ":" & FR_ColumnWethRho).NumberFormat = "0.00"

'        With .Cells.Interior
'            .Pattern = xlNone
'            .TintAndShade = 0
'            .PatternTintAndShade = 0
'        End With
'
'        With .Cells.Font
'            .ColorIndex = xlAutomatic
'            .TintAndShade = 0
'        End With
        
        .Cells.FormatConditions.Delete

        Set AutoFitRange = .Range(.Range(FR_ColumnSlpName & FR_HeaderRow + 1), _
        .Range(FR_Column6876DiscPercent & FR_HeaderRow + 1).End(xlDown))
        
        AutoFitRange.Columns.AutoFit
        
        .Range(FR_ChronusVersion).Columns.AutoFit
        
        'The following lines will delete the columns that are still not being filled
        .Columns("E").EntireColumn.Delete
        .Columns("H:M").EntireColumn.Delete
        .Columns("M:N").EntireColumn.Delete
        .Columns("S:T").EntireColumn.Delete
        
        'Formatting the Chronus info
        
        With .Range("A1").Font
            .Color = -16711681
            .TintAndShade = 0
        End With
        
        With .Range("B1").Font
            .Color = -11489280
            .TintAndShade = 0
        End With
        
        With .Range("C1").Font
            .Color = -1003520
            .TintAndShade = 0
        End With
        
        With .Range("A1:D1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With

    End With

End Sub

Sub FormatSlpStdBlkCorr(Optional AbsoluteUncertainty As Boolean = True)

    'Procedure that formats the SlpStdBlkCorr
    
    'Updated 26092015 - Autofilter arrows always displayed
    
    Dim RangeUnion As Range
    Dim SearchStr As Variant
    Dim UncertantyText As String
    Dim ScreenUpdt As Boolean

    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If

    If AbsoluteUncertainty = True Then
        UncertantyText = "abs"
    ElseIf AbsoluteUncertainty = False Then
        UncertantyText = "%"
    End If

    ScreenUpdt = Application.ScreenUpdating

    With SlpStdBlkCorr_Sh
    
'    'Checking if uncertainties in SlpStdBlkCorr sheet agree with AbsoluteUncertainty option
'    SearchStr = InStr(.Range(Column681Std & HeaderRow), "%")
'        If SearchStr = 0 And UncertantyText = "abs" Then
'
'            Application.ScreenUpdating = True: .Activate: Application.ScreenUpdating = ScreenUpdt
'
'                If MsgBox("Are uncertainties in this sheet percentage?", vbYesNo) = vbNo Then
'                    UncertantyText = "abs"
'                End If
'        End If

        .Range(ColumnExtStdRepro & ExtStdReproRow) = "Reproducibility of primary standard - " & StandardName_UPb
        .Range(ColumnExtStd68 & ExtStdReproRow + 1) = "ratio 6/8 1 sigma"
        .Range(ColumnExtStd75 & ExtStdReproRow + 1) = "ratio 7/5 1 sigma"
        .Range(ColumnExtStd76 & ExtStdReproRow + 1) = "ratio 7/6 1 sigma"
        .Range(ColumnWtdAvLabels & ExtStdReproRow + 2) = "Wtd Mean (from internal errs)"
        .Range(ColumnWtdAvLabels & ExtStdReproRow + 3) = "68%-conf. err. of mean"
        .Range(ColumnWtdAvLabels & ExtStdReproRow + 4) = "MSWD"
        .Range(ColumnWtdAvLabels & ExtStdReproRow + 5) = "Samples rejected"
        .Range(ColumnWtdAvLabels & ExtStdReproRow + 6) = "Probability of fit"
        
        .Range(ColumnID & HeaderRow) = "ID"
        .Range(ColumnSlpName & HeaderRow) = "Sample Name"
        .Range(Column68 & HeaderRow) = "ratio 6/8"
        .Range(Column681Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column68R & HeaderRow) = "R"
        .Range(Column68R2 & HeaderRow) = "R2"
        .Range(Column76 & HeaderRow) = "ratio 7/6"
        .Range(Column761Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column75 & HeaderRow) = "ratio 7/5"
        .Range(Column751Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column2 & HeaderRow) = "202 cps"
        .Range(Column21Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column4 & HeaderRow) = "204 cps"
        .Range(Column41Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column6 & HeaderRow) = "206 cps"
        .Range(Column61Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column7 & HeaderRow) = "207 cps"
        .Range(Column71Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column8 & HeaderRow) = "208 cps"
        .Range(Column81Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column32 & HeaderRow) = "232 cps"
        .Range(Column321Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column38 & HeaderRow) = "238 cps"
        .Range(Column381Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column64 & HeaderRow) = "ratio 6/4"
        .Range(Column641Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column74 & HeaderRow) = "ratio 7/4"
        .Range(Column741Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column28 & HeaderRow) = "ratio 2/8"
        .Range(Column281Std & HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(Column7568Rho & HeaderRow) = "Rho"
        
        .Range("A" & HeaderRow, .Range("A" & HeaderRow).End(xlToRight)).Font.Bold = True
        .Range("A" & HeaderRow, .Range("A" & HeaderRow).End(xlToRight)).HorizontalAlignment = xlCenter
        
        Set RangeUnion = Application.Union( _
            .Range(ColumnExtStd68 & ExtStdReproRow + 1), _
            .Range(ColumnExtStd75 & ExtStdReproRow + 1), _
            .Range(ColumnExtStd76 & ExtStdReproRow + 1))

        With RangeUnion
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .WrapText = True
        End With
        
        .Range(ColumnExtStdRepro & ExtStdReproRow).Font.Bold = True
        .Range(ColumnExtStdRepro & ExtStdReproRow).HorizontalAlignment = xlLeft
        
        Set RangeUnion = Application.Union( _
            .Range(ColumnWtdAvLabels & ExtStdReproRow + 2), _
            .Range(ColumnWtdAvLabels & ExtStdReproRow + 3), _
            .Range(ColumnWtdAvLabels & ExtStdReproRow + 4), _
            .Range(ColumnWtdAvLabels & ExtStdReproRow + 5), _
            .Range(ColumnWtdAvLabels & ExtStdReproRow + 6))

            With RangeUnion
                .Font.Bold = True
                .HorizontalAlignment = xlLeft
            End With
        
        'Isotopic ratios format
        Set RangeUnion = Application.Union( _
            ExtStd68Repro.Item(1), _
            ExtStd75Repro.Item(1), _
            ExtStd76Repro.Item(1))
            
            With RangeUnion
                .NumberFormat = "0.0000"
                .HorizontalAlignment = xlCenter
            End With
            
        'Isotopic ratios uncertainties format
        Set RangeUnion = Application.Union( _
            ExtStd68Repro.Item(2), _
            ExtStd75Repro.Item(2), _
            ExtStd76Repro.Item(2))
            
            With RangeUnion
                RangeUnion.NumberFormat = "0.0000"
                .HorizontalAlignment = xlCenter
            End With

        'MSWD
        Set RangeUnion = Application.Union( _
            ExtStd68Repro.Item(3), _
            ExtStd75Repro.Item(3), _
            ExtStd76Repro.Item(3))
            
            With RangeUnion
                RangeUnion.NumberFormat = "0"
                .HorizontalAlignment = xlCenter
            End With

        'Samples rejected format
        Set RangeUnion = Application.Union( _
            ExtStd68Repro.Item(4), _
            ExtStd75Repro.Item(4), _
            ExtStd76Repro.Item(4))
            
            With RangeUnion
                RangeUnion.NumberFormat = "0"
                .HorizontalAlignment = xlCenter
            End With
        
        'Probability of fit format
        Set RangeUnion = Application.Union( _
            ExtStd68Repro.Item(5), _
            ExtStd75Repro.Item(5), _
            ExtStd76Repro.Item(5))
            
            With RangeUnion
                RangeUnion.NumberFormat = "0.00E+00"
                .HorizontalAlignment = xlCenter
            End With
    
        With .Range(Column68R2 & HeaderRow).Characters(Start:=2, Length:=1).Font
            .Superscript = True
            .Bold = True
        End With

        .Range(ColumnID & ":" & ColumnID).NumberFormat = "0"
        
        Set RangeUnion = Application.Union(.Range(Column68 & ":" & Column68), .Range(Column76 & ":" & Column76), _
        .Range(Column28 & ":" & Column28), .Range(Column75 & ":" & Column75))
            RangeUnion.NumberFormat = "0.000"
            
        Set RangeUnion = Application.Union( _
            .Range(Column2 & ":" & Column2), _
            .Range(Column4 & ":" & Column4), _
            .Range(Column6 & ":" & Column6), _
            .Range(Column7 & ":" & Column7), _
            .Range(Column8 & ":" & Column8), _
            .Range(Column32 & ":" & Column32), _
            .Range(Column38 & ":" & Column38), _
            .Range(Column64 & ":" & Column64), _
            .Range(Column74 & ":" & Column74))
            
            RangeUnion.NumberFormat = "0.0"

        Set RangeUnion = Application.Union( _
            .Range(Column21Std & ":" & Column21Std), _
            .Range(Column41Std & ":" & Column41Std), _
            .Range(Column61Std & ":" & Column61Std), _
            .Range(Column71Std & ":" & Column71Std), _
            .Range(Column81Std & ":" & Column81Std), _
            .Range(Column321Std & ":" & Column321Std), _
            .Range(Column381Std & ":" & Column381Std), _
            .Range(Column641Std & ":" & Column641Std), _
            .Range(Column741Std & ":" & Column741Std))

            RangeUnion.NumberFormat = "0.0"
        
        SearchStr = InStr(.Range(Column681Std & HeaderRow), "(%)")
        
        Set RangeUnion = Application.Union(.Range(Column681Std & ":" & Column681Std), _
         .Range(Column761Std & ":" & Column761Std), .Range(Column751Std & ":" & Column751Std), _
         .Range(Column281Std & ":" & Column281Std), .Range(Column751Std & ":" & Column751Std), _
         ExtStd68Repro1std, ExtStd75Repro1std, ExtStd76Repro1std)
            
            If SearchStr = 0 Then
                RangeUnion.NumberFormat = "0.0000"
            Else
                RangeUnion.NumberFormat = "0.0"
            End If
            
        Set RangeUnion = Application.Union( _
        .Range(Column68R & ":" & Column68R), _
        .Range(Column68R2 & ":" & Column68R2), _
        .Range(Column7568Rho & ":" & Column7568Rho))
            
            RangeUnion.NumberFormat = "0.00"
                    
        With .Range("A" & HeaderRow, .Range("A" & HeaderRow).End(xlToRight))
            .Font.Bold = True
            
            If SlpStdBlkCorr_Sh.AutoFilterMode = False Then
                .AutoFilter
            End If
            
        End With
        
        .Range("A" & HeaderRow, .Range("A" & HeaderRow).End(xlToRight).End(xlDown)).Sort _
            key1:=.Range(ColumnID & HeaderRow), _
            order1:=xlAscending, _
            Header:=xlYes
            
        With .Range(Column68R2 & HeaderRow).Characters(Start:=2, Length:=1).Font
            .Superscript = True
            .Bold = True
        End With
        
        With .Range(.Range(ColumnID & HeaderRow), .Range(ColumnID & HeaderRow).End(xlToRight).End(xlDown))
            .Columns.AutoFit 'Autofit all columns
            .HorizontalAlignment = xlCenter 'All cells horizontal alignment is set to xlcenter
        End With
        
        'Below the program selects a cell without using the select method. This is very important because
        'it doesn't matter ifthe user is doing anything else in the computer, this program is able
        'to select or activate what I want.
        Application.GoTo .Range("A" & HeaderRow + 1)
            
            With ActiveWindow
                .SplitColumn = 2
                .SplitRow = HeaderRow
                .FreezePanes = True
            End With

    End With

    Call HighlightIntStd(SlpStdBlkCorr_Sh)
    Call HighlightExtStd(SlpStdBlkCorr_Sh)
    Call HighlightNAs(SlpStdBlkCorr_Sh)
    
End Sub

Sub FormatSlpStdCorr(Optional AbsoluteUncertainty As Boolean = True, Optional HighlightAnalysis As Boolean = True)
    
    'Procedure that formats the SlpStdCorr sheet
    
    'Updated 26092015 - Autofilter arrows always displayed
    
    Dim SearchStr As Variant
    Dim RangeUnion As Range
    Dim UncertantyText As String
    Dim ScreenUpdt As Boolean

    If SlpStdCorr_Sh Is Nothing Then
        Call PublicVariables
    End If
    
    If AbsoluteUncertainty = True Then
        UncertantyText = "abs"
    ElseIf AbsoluteUncertainty = False Then
        UncertantyText = "%"
    End If

    With SlpStdCorr_Sh
    
'    'Checking if uncertainties in SlpStdBlkCorr sheet agree with AbsoluteUncertainty option
'    SearchStr = InStr(.Range(StdCorr_Column681Std & HeaderRow), "%")
'        If SearchStr = 0 And UncertantyText = "abs" Then
'
'            Application.ScreenUpdating = True: .Activate: Application.ScreenUpdating = ScreenUpdt
'
'                If MsgBox("Are uncertainties in this sheet percentage?", vbYesNo) = vbNo Then
'                    UncertantyText = "abs"
'                End If
'        End If
    
        .Range(StdCorr_ColumnID & StdCorr_HeaderRow) = "ID"
        .Range(StdCorr_SlpName & StdCorr_HeaderRow) = "Sample Name"
        .Range(StdCorr_TetaFactor & StdCorr_HeaderRow) = "Teta"
        .Range(StdCorr_Column68 & StdCorr_HeaderRow) = "ratio 6/8"
        .Range(StdCorr_Column681Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_Column68R & StdCorr_HeaderRow) = "R"
        .Range(StdCorr_Column68R2 & StdCorr_HeaderRow) = "R2"
        .Range(StdCorr_Column76 & StdCorr_HeaderRow) = "ratio 7/6"
        .Range(StdCorr_Column761Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_Column75 & StdCorr_HeaderRow) = "ratio 7/5"
        .Range(StdCorr_Column751Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_Column2 & StdCorr_HeaderRow) = "202 cps"
        .Range(StdCorr_Column21Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_Column4 & StdCorr_HeaderRow) = "204 cps"
        .Range(StdCorr_Column41Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_Column64 & StdCorr_HeaderRow) = "ratio 6/4"
        .Range(StdCorr_Column641Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_Column74 & StdCorr_HeaderRow) = "ratio 7/4"
        .Range(StdCorr_Column741Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_ColumnF206 & StdCorr_HeaderRow) = "f(206)%"
        .Range(StdCorr_Column28 & StdCorr_HeaderRow) = "ratio 2/8"
        .Range(StdCorr_Column281Std & StdCorr_HeaderRow) = "1 std (" & UncertantyText & ")"
        .Range(StdCorr_Column7568Rho & StdCorr_HeaderRow) = "Rho"
        
        .Range(StdCorr_ColumnF206 & StdCorr_HeaderRow) = "206* (%)"
        .Range(StdCorr_Column68AgeMa & StdCorr_HeaderRow) = "6/8 (Ma)"
        .Range(StdCorr_Column68AgeMa1std & StdCorr_HeaderRow) = "6/8 1 std abs"
        .Range(StdCorr_Column75AgeMa & StdCorr_HeaderRow) = "7/5 (Ma)"
        .Range(StdCorr_Column75AgeMa1std & StdCorr_HeaderRow) = "7/5 1 std abs"
        .Range(StdCorr_Column76AgeMa & StdCorr_HeaderRow) = "7/6 (Ma)"
        .Range(StdCorr_Column76AgeMa1std & StdCorr_HeaderRow) = "7/6 1 std abs"
        .Range(StdCorr_Column6876Conc & StdCorr_HeaderRow) = "6/8 7/6 Conc (%)"
        .Range(StdCorr_Column6875Conc & StdCorr_HeaderRow) = "6/8 7/5 Conc (%)"
        
        .Range("A" & StdCorr_HeaderRow, .Range("A" & StdCorr_HeaderRow).End(xlToRight)).Font.Bold = True
        .Range("A" & StdCorr_HeaderRow, .Range("A" & StdCorr_HeaderRow).End(xlToRight)).HorizontalAlignment = xlCenter
        
        With .Range(StdCorr_Column68R2 & HeaderRow).Characters(Start:=2, Length:=1).Font
            .Superscript = True
            .Bold = True
        End With

        .Range(StdCorr_ColumnID & ":" & StdCorr_ColumnID).NumberFormat = "0"
        
        Set RangeUnion = Application.Union(.Range(StdCorr_Column68 & ":" & StdCorr_Column68), _
        .Range(StdCorr_Column28 & ":" & StdCorr_Column28), .Range(StdCorr_Column75 & ":" & StdCorr_Column75), _
        .Range(StdCorr_ColumnF206 & ":" & StdCorr_ColumnF206))
            RangeUnion.NumberFormat = "0.0000"
            
        .Range(StdCorr_Column76 & ":" & StdCorr_Column76).NumberFormat = "0.00000"
        
        Set RangeUnion = Application.Union(.Range(StdCorr_Column2 & ":" & StdCorr_Column2), .Range(StdCorr_Column4 & ":" & StdCorr_Column4) _
        , .Range(StdCorr_Column64 & ":" & StdCorr_Column64), .Range(StdCorr_Column74 & ":" & StdCorr_Column74), .Range(StdCorr_Column6876Conc & ":" & StdCorr_Column6875Conc))
            RangeUnion.NumberFormat = "0.0"

        Set RangeUnion = Application.Union(.Range(StdCorr_Column21Std & ":" & StdCorr_Column21Std), .Range(StdCorr_Column41Std & ":" & StdCorr_Column41Std) _
        , .Range(StdCorr_Column641Std & ":" & StdCorr_Column641Std), .Range(StdCorr_Column741Std & ":" & StdCorr_Column741Std))
            RangeUnion.NumberFormat = "0.0"
        
        Set RangeUnion = Application.Union(.Range(StdCorr_Column68AgeMa1std & ":" & StdCorr_Column68AgeMa1std), .Range(StdCorr_Column75AgeMa1std & ":" & StdCorr_Column75AgeMa1std) _
        , .Range(StdCorr_Column76AgeMa1std & ":" & StdCorr_Column76AgeMa1std))
            RangeUnion.NumberFormat = "0.0"
        
        Set RangeUnion = Application.Union(.Range(StdCorr_Column68AgeMa & ":" & StdCorr_Column68AgeMa), .Range(StdCorr_Column75AgeMa & ":" & StdCorr_Column75AgeMa) _
        , .Range(StdCorr_Column76AgeMa & ":" & StdCorr_Column76AgeMa))
            RangeUnion.NumberFormat = "0.0"
        
        SearchStr = InStr(.Range(StdCorr_Column681Std & HeaderRow), "(%)")
        
        Set RangeUnion = Application.Union(.Range(StdCorr_Column681Std & ":" & StdCorr_Column681Std), _
         .Range(StdCorr_Column761Std & ":" & StdCorr_Column761Std), .Range(StdCorr_Column751Std & ":" & StdCorr_Column751Std), _
         .Range(StdCorr_Column281Std & ":" & StdCorr_Column281Std), .Range(StdCorr_Column751Std & ":" & StdCorr_Column751Std))
            
            If SearchStr = 0 Then
                RangeUnion.NumberFormat = "0.0000"
            Else
                RangeUnion.NumberFormat = "0.0"
            End If
            
        Set RangeUnion = Application.Union( _
        .Range(StdCorr_Column68R & ":" & StdCorr_Column68R), _
        .Range(StdCorr_Column68R2 & ":" & StdCorr_Column68R2), _
        .Range(StdCorr_TetaFactor & ":" & StdCorr_TetaFactor), _
        .Range(StdCorr_Column7568Rho & ":" & StdCorr_Column7568Rho))
        
            RangeUnion.NumberFormat = "0.00"
                    
        With .Range("A" & StdCorr_HeaderRow, .Range("A" & StdCorr_HeaderRow).End(xlToRight))
            Application.GoTo .Range("A" & 1)
            .Font.Bold = True

            If SlpStdCorr_Sh.AutoFilterMode = False Then
                .AutoFilter
            End If

        End With
        
        With .Range(StdCorr_Column68R2 & StdCorr_HeaderRow).Characters(Start:=2, Length:=1).Font
            .Superscript = True
            .Bold = True
        End With

        Application.GoTo .Range("A" & StdCorr_HeaderRow)
            
            With ActiveWindow
                .SplitColumn = 2
                .SplitRow = StdCorr_HeaderRow
                .FreezePanes = True
            End With

        Cells.HorizontalAlignment = xlCenter 'All cells horizontal alignment is set to xlcenter
        Cells.Columns.AutoFit 'Autofit all StdCorr_columns
        
        'Conditional Formatting for 68 ratios uncertainties
        With .Range(StdCorr_Column681Std & ":" & StdCorr_Column681Std)
        
            .FormatConditions.AddColorScale ColorScaleType:=3
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            
            With .FormatConditions(1)
            
                With .ColorScaleCriteria(1).FormatColor
                    .Color = 8109667
                    .TintAndShade = 0
                End With
                                
                With .ColorScaleCriteria(2).FormatColor
                    .Color = 8711167
                    .TintAndShade = 0
                End With
                                    
                With .ColorScaleCriteria(3).FormatColor
                    .Color = 7039480
                    .TintAndShade = 0
                End With
                
                .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(2).Type = xlConditionValuePercentile
                .ColorScaleCriteria(2).Value = 50

            End With
            
        End With
            
        'Conditional Formatting for 76 ratios uncertainties
        With .Range(StdCorr_Column761Std & ":" & StdCorr_Column761Std)
        
            .FormatConditions.AddColorScale ColorScaleType:=3
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            
            With .FormatConditions(1)
            
                With .ColorScaleCriteria(1).FormatColor
                    .Color = 8109667
                    .TintAndShade = 0
                End With
                                
                With .ColorScaleCriteria(2).FormatColor
                    .Color = 8711167
                    .TintAndShade = 0
                End With
                                    
                With .ColorScaleCriteria(3).FormatColor
                    .Color = 7039480
                    .TintAndShade = 0
                End With
                
                .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(2).Type = xlConditionValuePercentile
                .ColorScaleCriteria(2).Value = 50

            End With
        
        End With
        
        'Conditional Formatting for 75 ratios uncertainties
        With .Range(StdCorr_Column751Std & ":" & StdCorr_Column751Std)
        
            .FormatConditions.AddColorScale ColorScaleType:=3
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            
            With .FormatConditions(1)
            
                With .ColorScaleCriteria(1).FormatColor
                    .Color = 8109667
                    .TintAndShade = 0
                End With
                                
                With .ColorScaleCriteria(2).FormatColor
                    .Color = 8711167
                    .TintAndShade = 0
                End With
                                    
                With .ColorScaleCriteria(3).FormatColor
                    .Color = 7039480
                    .TintAndShade = 0
                End With
                
                .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(2).Type = xlConditionValuePercentile
                .ColorScaleCriteria(2).Value = 50

            End With
            
        End With
            
        'Conditional Formatting for common lead
        With .Range(StdCorr_ColumnF206 & ":" & StdCorr_ColumnF206)
        
            .FormatConditions.AddColorScale ColorScaleType:=3
            .FormatConditions(.FormatConditions.count).SetFirstPriority
            .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
            
            With .FormatConditions(1)
            
                With .ColorScaleCriteria(1).FormatColor
                    .Color = 8109667
                    .TintAndShade = 0
                End With
                                
                With .ColorScaleCriteria(2).FormatColor
                    .Color = 8711167
                    .TintAndShade = 0
                End With
                                    
                With .ColorScaleCriteria(3).FormatColor
                    .Color = 7039480
                    .TintAndShade = 0
                End With
                
                .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
                .ColorScaleCriteria(2).Type = xlConditionValuePercentile
                .ColorScaleCriteria(2).Value = 50

            End With

        End With
        
    End With
    
    If HighlightAnalysis = True Then
        Call HighlightIntStd(SlpStdCorr_Sh)
        Call HighlightExtStd(SlpStdCorr_Sh)
        Call HighlightNAs(SlpStdCorr_Sh)
    End If
    
End Sub

Sub FormatStartANDOptions()

    If InternalStandardCheck_UPb Is Nothing Then
        Call PublicVariables
    End If

    With StartANDOptions_Sh
    
            'Naming cells
            .Range("A1") = "Start"
            .Range("A2") = "Basic Information"
            .Range("A3") = "Sample Name"
            .Range("A4") = "Data reduced on"
            .Range("A5") = "Data reduced by"
            .Range("A6") = "Folder path"
                                            
            .Range("A8") = "Standards"
            .Range("A9") = "External standard"
            .Range("A10") = "Internal standard"
            .Range("A13") = "Data acquired using"
            .Range("A16") = "Isotope detector"
            
            .Range("A16") = "Isotope detector"
            .Range("A17") = "206 Isotope"
            .Range("A20") = "Options"
            
            .Range("A21") = "Constants"
            .Range("A22") = "235U/238U"
            .Range("A23") = "202Hg/204Hg"
            .Range("A24") = "mV to CPS factor"
            
            .Range("A26") = "External Standard"
            .Range("A27") = "Standard"
            .Range("A28") = "Name"
            .Range("A29") = "Mineral"
            .Range("A30") = "Description"
            
            .Range("A31") = "Ratios"
            
            .Range("A33") = "206Pb/238U"
            .Range("A34") = "207Pb/235U"
            .Range("A35") = "207Pb/206Pb"
            
            .Range("A37") = "Concentrations"
            
            .Range("A39") = "U (ppm)"
            .Range("A40") = "Th (ppm)"
            
            .Range("A42") = "Address"
            .Range("A43") = "Raw Data File"
            .Range("A45") = "202"
            .Range("A46") = "204"
            .Range("A47") = "206"
            .Range("A48") = "207"
            .Range("A49") = "208"
            .Range("A50") = "232"
            .Range("A51") = "238"
            .Range("A53") = "Cycles Time"
            .Range("A54") = "Analysis Date"
            .Range("A55") = "Number of Cycles"
            .Range("A56") = "Cycles duration (ss.ms)"
                '.Range("A56").AddComment ("Integration Time")
            
            .Range("B32") = "Ratios"
            .Range("B38") = "Concentration"
            .Range("B44") = "Isotopes Signal"
            
            .Range("C2") = "Analyses names"
            .Range("C3") = "Blanks name"
            .Range("C4") = "Samples name"
            .Range("C5") = "Standards name"
            .Range("C21") = "1std"
            .Range("C32") = "Error"
            .Range("C38") = "Error"
            .Range("C44") = "Header"
            
            .Range("D21") = "Error Propagation"
            .Range("D22") = "Blank"
            .Range("D23") = "Primary standard analyses"
            .Range("D24") = "Primary standard reproducibility (MSWD)"
                '.Range("D24").HorizontalAlignment = xlRight
            .Range("D25") = "Cert. Primary standard"
            .Range("D44") = "Analyzed"
                
            .Range("D32") = "1s or 2s"
            .Range("D38") = "1s or 2s"
            
            .Range("E38") = "abs or %"
            .Range("E32") = "abs or %"
    
            'Cells formatting
            With .Range("A1:G1").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
                          
            With .Range("A2:G18").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorAccent6
                .TintAndShade = 0.599993896298105
                .PatternTintAndShade = 0
            End With
            
            .Range("A1,A2,C2,A8,A13,A16,A20,A21,C21,A26:A27,A31,A37,A42:A43,A53:A56,D21").Font.Bold = True
            .Range("A3:A6,C3:C5,A9:A10,A17:A18,A27:A29,A33:A35,A39:A40,A45:A51,B32:E32,B38:E38,B44:D44").Font.Italic = True
            
            With .Range("A20:G20").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .ThemeColor = xlThemeColorLight2
                .TintAndShade = 0.399975585192419
                .PatternTintAndShade = 0
                
            End With
            
            With .Range("A21:G56").Interior
                .PatternColorIndex = xlAutomatic
                .Color = 15773696
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            
            Columns("A:A").EntireColumn.AutoFit
            .Range("D21", "D24").Columns.AutoFit
            .Range("B4").Columns.AutoFit
            .Range("E38,E32").Columns.AutoFit

            
            .Range("A53:A56").HorizontalAlignment = xlLeft
            .Range("B6").HorizontalAlignment = xlLeft
            
            With .Range("B3:B6,B9:B18,D3:D5").Font
                .Color = -16776961
                .TintAndShade = 0
            End With
            
            With .Range("B22:B24,C22:C24,B28:B30,B33:B35,C33:C35,B39:B40,C39:C40,B45:C51,B53:B56").Font
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
            End With
            
            .Range("B22:B24,C22:C24,B28:B29,B33:B35,C33:C35,B39:B40,C39:C40,B45:C51,B53:B56,E38,E32").HorizontalAlignment = xlCenter
                .Range("B30").HorizontalAlignment = xlLeft
                
            With InternalStandardCheck_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="TRUE,FALSE"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "Please, select true or false from the list."
                .ErrorMessage = "Only true or false."
                .ShowInput = True
                .ShowError = True
            End With
            
            With CheckData_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="TRUE,FALSE"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "Please, select true or false from the list."
                .ErrorMessage = "Only true or false."
                .ShowInput = True
                .ShowError = True
            End With
            
            With ErrBlank_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="TRUE,FALSE"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "Please, select true or false from the list."
                .ErrorMessage = "Only true or false."
                .ShowInput = True
                .ShowError = True
            End With
            
            With ErrExtStd_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="TRUE,FALSE"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "Please, select true or false from the list."
                .ErrorMessage = "Only true or false."
                .ShowInput = True
                .ShowError = True
            End With

            With SpotRaster_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Spot; Raster"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "Please, select spot or raster from the list. " & _
                    "This changes the way ratio 206Pb/238U is calculated (by intercept method or simple average, respectively)."
                .ErrorMessage = "Only Spot or Raster."
                .ShowInput = True
                .ShowError = True
            End With

            With Detector206_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="Faraday Cup; MIC"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "Please, select Faraday Cup or MIC from the list. " & _
                "If Faraday Cup is selected, it's necessary to multiply 206 isotope signal by V to CPS factor."
                .ErrorMessage = "Only Faraday Cup or MIC."
                .ShowInput = True
                .ShowError = True
            End With
            
            With ErrBlank_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="TRUE,FALSE"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "True means blank errors will be propagated."
                .ErrorMessage = "Only true or false."
                .ShowInput = True
                .ShowError = True
            End With
            
            With ErrExtStd_UPb.Validation
                .Delete
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                xlBetween, Formula1:="TRUE,FALSE"
                .IgnoreBlank = True
                .InCellDropdown = True
                .InputTitle = ""
                .ErrorTitle = "Value is not acceptable!"
                .InputMessage = "True means external standard errors will be propagated."
                .ErrorMessage = "Only true or false."
                .ShowInput = True
                .ShowError = True
            End With
            
    End With
    
    Application.GoTo StartANDOptions_Sh.Range("A1")
        
        With ActiveWindow
            .FreezePanes = True
            .SplitColumn = 0
            .SplitRow = 0
        End With

End Sub

Sub HighlightNAs(Sh As Worksheet)

    Dim FindCells As Object
    Dim FirstAddress As String
    Dim NAs As String
    
    NAs = "n.a."
    
    With Sh.Cells
        Set FindCells = .Find(NAs, LookIn:=xlValues)
        If Not FindCells Is Nothing Then
            FirstAddress = FindCells.Address
            Do
                With FindCells
                
                    With .Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = -0.149998474074526
                        .PatternTintAndShade = 0
                    End With
                    
                    With .Font
                        .Color = -16776961
                        .TintAndShade = 0
                        .Bold = True
                    End With
                    
                End With
                
                Set FindCells = .FindNext(FindCells)
            Loop While Not FindCells Is Nothing And FindCells.Address <> FirstAddress
        End If
    End With

End Sub

Sub HighlightIntStd(Sh As Worksheet)
    
    'This procedure will highlight the secondary standards in the sheets indicated
    'following a colorscale with 10 different colors. When more than 10 different
    'standards are analyzed, the color start repeating.

    'Updated 16102015
    
    Dim FindCells As Object
    Dim FirstAddress As String
    Dim IntStdNames() As String
    Dim Counter As Integer
    Dim TCnumber As Long 'ThemeColor number
    
    If InternalStandard_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    IntStdNames = Split(InternalStandard_UPb, ",")
    
    TCnumber = 1
    
    If IsArrayEmpty(IntStdNames) = False Then
        For Counter = LBound(IntStdNames) To UBound(IntStdNames)
            IntStdNames(Counter) = Replace(IntStdNames(Counter), " ", "")
        Next
        
        For Counter = LBound(IntStdNames) To UBound(IntStdNames)
            With Sh.Cells
                Set FindCells = .Find(IntStdNames(Counter), LookIn:=xlValues)
                If Not FindCells Is Nothing Then
                    FirstAddress = FindCells.Address
                    Do
                        With .Range(FindCells.Address, .Range(FindCells.Address).End(xlToRight))
                        
                            With .Interior
                                    .Pattern = xlSolid
                                    .PatternColorIndex = xlAutomatic
                                    .PatternTintAndShade = 0
                                    .ThemeColor = TCnumber
                                    .TintAndShade = -0.25
                            End With
                            
                            With .Font
                                .ThemeColor = xlThemeColorDark1
                                .TintAndShade = 0
                            End With
                            
                        End With

                        Set FindCells = .FindNext(FindCells)
                    Loop While Not FindCells Is Nothing And FindCells.Address <> FirstAddress
                End If
            End With
                
                'These following lines changes the background color for the next standard
                Select Case TCnumber
                    Case Is < 10
                        TCnumber = TCnumber + 1
                    Case Is = 10
                        TCnumber = 1
                End Select

        Next
    End If

End Sub

Sub HighlightExtStd(Sh As Worksheet)

    Dim FindCells As Object
    Dim FirstAddress As String

    If ExternalStandard_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    If Not IsEmpty(ExternalStandardName_UPb) Then
        With Sh.Cells
            Set FindCells = .Find(ExternalStandardName_UPb, LookIn:=xlValues)
            If Not FindCells Is Nothing Then
                FirstAddress = FindCells.Address
                Do
                    With .Range(FindCells.Address, .Range(FindCells.Address).End(xlToRight))
                    
                        With .Interior
                            .Pattern = xlSolid
                            .PatternColorIndex = xlAutomatic
                            .Color = 65535
                            .TintAndShade = 0
                            .PatternTintAndShade = 0
                        End With
                                            
                    End With
                    
                    Set FindCells = .FindNext(FindCells)
                Loop While Not FindCells Is Nothing And FindCells.Address <> FirstAddress
            End If
        End With
    End If

    If Sh.Name = SlpStdBlkCorr_Sh.Name Then
        
        With Sh.Rows("1:1").Interior
            .Pattern = xlNone
        End With
    
    End If

End Sub
