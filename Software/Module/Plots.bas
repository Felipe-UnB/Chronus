Attribute VB_Name = "Plots"
Option Explicit
    
    'Variable available in the Plots module scope
    
    Dim ChartDataX As Range
    
    Dim ChartDataY_RawSignal2 As Range
    Dim ChartDataY_RawSignal4 As Range
    Dim ChartDataY_RawSignal6 As Range
    Dim ChartDataY_RawSignal7 As Range
    Dim ChartDataY_RawSignal8 As Range
    Dim ChartDataY_RawSignal32 As Range
    Dim ChartDataY_RawSignal38 As Range
    Dim ChartDataY_64 As Range
    Dim ChartDataY_74 As Range
    Dim ChartDataY_28 As Range
    Dim ChartDataY_75 As Range
    Dim ChartDataY_68 As Range
    Dim ChartDataY_76 As Range

Sub OpenAnalysisToPlot_ByIDs(ID As Integer, Optional ReopeningInPlot As Boolean = False)

    'This program will open each analysis based on its ID and clear the cycles that shouldn't be used by calling
    ' a different program.
    
    'IMPORTANT OBSERVATIONS
    'For samples, internal and external standards, calcBlank procedure must be already executed succesfully because
    'blank result must be discounted before plotting theses analysis.

    Dim SlpStd() As Integer 'Array with samples and INTERNAL standards (treated as samples) IDs only
    Dim ExtStd() As Integer 'Array with external standards IDs only
    Dim Blk() As Integer 'Array with blanks IDs only
    Dim C As Integer
    Dim a As Variant
    Dim H As Double
    Dim FindIDObj As Object
    Dim B As Range 'used to loop through all cells in Plot_Column75 cells
    Dim d As Double
    Dim SheetsNamesArr() As String
    Dim SheetNameNum As Integer
    Dim TargetSheet As Worksheet
    Dim ShName As String
    Dim LastID As Long
    
    If TypeName(ID) <> "Integer" Then
        MsgBox "The type of the choosen ID (" & ID & ") is " & TypeName(ID) & ", but it must be an integer!"
            End
    End If
    
    Call PublicVariables
        
    If IsArrayEmpty(BlkFound) = True Then
        Call IdentifyFileType
    End If
    
    On Error Resume Next
        If AnalysesList(0).sample = "" Then
            Call LoadSamListMap
            Call LoadStdListMap
        End If
    On Error GoTo 0
    
    If IsArrayEmpty(PathsNamesIDsTimesCycles) = True Then
        Call SetPathsNamesIDsTimesCycles
    End If

    If Detector206_UPb = "Faraday Cup" Then
            H = mVtoCPS_UPb
        ElseIf Detector206_UPb = "MIC" Then
            H = 1
        Else
            MsgBox "Please, indicate if 206Pb was analyzed using Faraday cup or Ion counter."
                Application.GoTo (StartANDOptions_Sh.Range("A1"))
                    End
    End If

    'Adding worksheets to plot data
    Set Plot_Sh = mwbk.Worksheets.Add(, SlpStdBlkCorr_Sh)
        
        Plot_Sh.Select: ActiveWindow.Zoom = 70
        
    Set Plot_ShHidden = mwbk.Worksheets.Add(, Plot_Sh)
        Plot_ShHidden.Visible = xlSheetHidden
        
    'Set FindIDObj = SamList_Sh.Columns(SamList_Sh.Range(SamList_ID & ":" & SamList_ID).Column).Find(ID)
                    
    ReDim SlpStd(1 To UBound(SlpFound) + UBound(IntStdFound) + 2) As Integer
    ReDim ExtStd(1 To UBound(StdFound) + 1) As Integer
    ReDim Blk(1 To UBound(BlkFound) + 1) As Integer
    
    C = 1

        'Samples IDs are copied to a different array (SlpStd) which accepts only numbers (IDs)
        For Each a In SlpFound
            SlpStd(C) = SamList_Sh.Range(a).Offset(, 1)
            C = C + 1
        Next
        
        'Internal standards IDs are copied to a different array (SlpStd) which accepts only numbers (IDs)
        For Each a In IntStdFound
            SlpStd(C) = SamList_Sh.Range(a).Offset(, 1)
            C = C + 1
        Next
            
    C = 1
    
        'External standards IDs are copied to a different array (SlpStd) which accepts only numbers (IDs)
        For Each a In StdFound
            ExtStd(C) = SamList_Sh.Range(a).Offset(, 1)
            C = C + 1
        Next
        
    C = 1
    
        'Blanks IDs are copied to a different array (SlpStd) which accepts only numbers (IDs)
        For Each a In BlkFound
            Blk(C) = SamList_Sh.Range(a).Offset(, 1)
            C = C + 1
        Next
    
    'The name of the worksheet where data will be plotted should be the same as the analysis
    Application.GoTo SamList_Sh.Range("A1") 'This line is just necessary to activate the Samlist_Sh
    With SamList_Sh.Columns(SamList_Sh.Range(SamList_ID & ":" & SamList_ID).Column)
        Set FindIDObj = .Find(ID)
            
            On Error Resume Next
                ShName = SamList_Sh.Range(FindIDObj.Address).Offset(, -1).Text
                    ShName = Left(ShName, 10)
                
                Plot_Sh.Name = ShName & " (" & ID & ")_Plot" 'Name of the Plot_Sh
                    Plot_ShHidden.Name = ShName & " (" & ID & ")_PlotHidden" 'Name of the Plot_ShHidden
                
'                If Err.Number <> 0 And Not FindIDObj Is Nothing Then 'In case renaming worksheet fails but ID is valid
'                    Err.Clear
'                        Plot_Sh.Name = "(" & ID & ")_Plot" 'Name of the Plot_Sh
'                            Plot_ShHidden.Name = "(" & ID & ")_PlotHidden" 'Name of the Plot_ShHidden
'                End If
                
                    If Err.Number = 1004 Then
                        On Error GoTo 0 'This statement, besides restoring normal error handling, it sets err.number=0
                        
                        If ReopeningInPlot = False Then 'This is used basically to check if the analysis is being opened again just because the user wants to all data in the Plot sheet
                        
                            If MsgBox("There is a worksheet with the same name of the analysis you are trying to open, " & _
                            ShName & " (" & ID & ")" & "." & vbNewLine & _
                            "Would you like to reopen this analysis?", vbYesNo) = vbYes Then
                            
                                Application.DisplayAlerts = False
                                    Plot_Sh.Delete: Plot_ShHidden.Delete
                                Application.DisplayAlerts = True
                            
                                    Set Plot_Sh = mwbk.Worksheets(ShName & " (" & ID & ")_Plot")
                                        Set Plot_ShHidden = mwbk.Worksheets(ShName & " (" & ID & ")_PlotHidden")
                                            Plot_Sh.Cells.Clear
                                                Plot_ShHidden.Cells.Clear
                            Else
                                
                                Application.DisplayAlerts = False
                                    Plot_Sh.Delete: Plot_ShHidden.Delete
                                Application.DisplayAlerts = True
                                
                                End
                            
                            End If
                                                    
                        Else
                        
                            Application.DisplayAlerts = False
                                Plot_Sh.Delete: Plot_ShHidden.Delete
                            Application.DisplayAlerts = True
                        
                                    Set Plot_Sh = mwbk.Worksheets(ShName & " (" & ID & ")_Plot")
                                        Set Plot_ShHidden = mwbk.Worksheets(ShName & " (" & ID & ")_PlotHidden")
                                            Plot_Sh.Cells.Clear
                                                Plot_ShHidden.Cells.Clear

                        End If
                        
                    ElseIf Err.Number = 91 Then
                        On Error GoTo 0
                        
                        MsgBox "ID " & ID & " was not found. Please, check it and then retry."
                    
                            Application.DisplayAlerts = False
                                Plot_Sh.Delete: Plot_ShHidden.Delete
                            Application.DisplayAlerts = True
                    
                                Call UnloadAll
                                    End
                        
                    End If
            On Error GoTo 0
    
    End With

    'Lines below will write the name of the analysis and its ID to Plot_sh
    Plot_Sh.Range(Plot_IDCell) = ID 'ID of the analysis
        Plot_Sh.Range(Plot_AnalysisName) = ShName 'Name of the analysis
            Plot_Sh.Range(Plot_AnalysisName).Font.Italic = True
    
    'The lines below will test ID to check the type of analysis (sample, internal standard, external standard or blank).
    'Then, blank signal will be discounted and the analysis will be plotted.
    
    
    'SAMPLES, INTERNAL AND EXTERNAL STANDARDS-------------------------------------------
    
    
    If FindItemInArray(ID, SlpStd) = True Or FindItemInArray(ID, ExtStd) = True Then 'True means that the program is dealing with a sample, internal or external standard analysis
                        
        If FindItemInArray(ID, SlpStd) = True Then
            Call CalcSlp_BlkCorr(ID, H, False, , False)
        
            ElseIf FindItemInArray(ID, ExtStd) = True Then
                Call CalcExtStd_BlkCorr(ID, H, False, , False)
            
        End If
        
        'Data from each analysis discounted the blank is copied to the Plot sheet
        Call Plot_CopyData(WBSlp.Worksheets(1), Plot_Sh)
            Call Plot_CopyData(WBSlp.Worksheets(1), Plot_ShHidden)
        
        WBSlp.Close savechanges:=False
        
        Call Plot_OrdinaryCalculations(Plot_Sh)
            Call Plot_OrdinaryCalculations(Plot_ShHidden)
                                    
                                    
    'BLANK ------------------------------------------------------------
    
    
    ElseIf FindItemInArray(ID, Blk) = True Then 'True means that the program is dealing with a blank analysis
            
            'The lines below open the raw data file
            On Error Resume Next
                Set WBSlp = Workbooks.Open(PathsNamesIDsTimesCycles(RawDataFilesPaths, ID)) 'ActiveWorkbook
                    If Err.Number <> 0 Then
                        MsgBox MissingFile1 & PathsNamesIDsTimesCycles(RawDataFilesPaths, ID) & MissingFile2
                            Call UpdateFilesAddresses
                                Call UnloadAll
                                    End
                    End If
            On Error GoTo 0

        'Blank data is copied to the Plot_sh
            Call ClearCycles(WBSlp, PathsNamesIDsTimesCycles(Cycles, ID))
            
                Call CyclesTime(WBSlp.Worksheets(1).Range(RawCyclesTimeRange))
            
                    Call Plot_CopyData(WBSlp.Worksheets(1), Plot_Sh)
                    
                        Call Plot_CopyData(WBSlp.Worksheets(1), Plot_ShHidden)
            
            WBSlp.Close savechanges:=False
            
            ReDim SheetsNamesArr(1 To 2) As String
                SheetsNamesArr(1) = Plot_Sh.Name
                SheetsNamesArr(2) = Plot_ShHidden.Name
                
            For SheetNameNum = LBound(SheetsNamesArr) To UBound(SheetsNamesArr)
            
                Set TargetSheet = mwbk.Worksheets(SheetsNamesArr(SheetNameNum))
                
                    With TargetSheet
                        
                        'It's necessary to multiply 208Pb, 232Th, 238U e 206Pb (if analysed in faraday cup) by VtoCPS constant.
                        'In opposite to what happens with samples and standards, blank analysis does not have a specific program
                        'to deal with just the calculations below.
                        For Each B In .Range(Plot_Column8 & Plot_HeaderRow + 1, Plot_Column8 & Plot_HeaderRow + RawNumberCycles_UPb)
                            If Not IsEmpty(B) Then
                                B = B * mVtoCPS_UPb
                            End If
                        Next
                        
                        For Each B In .Range(Plot_Column32 & Plot_HeaderRow + 1, Plot_Column32 & Plot_HeaderRow + RawNumberCycles_UPb)
                            If Not IsEmpty(B) Then
                                B = B * mVtoCPS_UPb
                            End If
                        Next
            
                        For Each B In .Range(Plot_Column38 & Plot_HeaderRow + 1, Plot_Column38 & Plot_HeaderRow + RawNumberCycles_UPb)
                            If Not IsEmpty(B) Then
                                B = B * mVtoCPS_UPb
                            End If
                        Next
                        
                        For Each B In .Range(Plot_Column6 & Plot_HeaderRow + 1, Plot_Column6 & Plot_HeaderRow + RawNumberCycles_UPb)
                            If Not IsEmpty(B) Then
                                B = B * H
                            End If
                        Next
                        
                        Call Plot_OrdinaryCalculations(TargetSheet)
            
                    End With
            Next
    
    Else 'Case when analysis associated with the ID being evaluated was not considered sample, neither standard (it was processed).
        
        If ID <= 0 Then
            MsgBox "Invalid ID (" & ID & ")" 'There is not any ID smaller than or equal to 0, only if the user change something
        Else
            MsgBox ShName & " (ID " & ID & ") was not processed because its name is not compatible with the name of samples or standards."
        End If
        
            Application.DisplayAlerts = False
                Plot_Sh.Delete: Plot_ShHidden.Delete
            Application.DisplayAlerts = True

                FailToOpen = True
                    Exit Sub
    
    End If
    
    Call AddCodePlotSh(Plot_Sh)
    
End Sub

Sub testaddcode()

    Call AddCodePlotSh(ActiveSheet)
    
    
End Sub


Sub AddCodePlotSh(Plot_Sh As Worksheet)

    'Created 05042016
    'This procedure will add automatically an event handler for the plot sheets: every time the user delete
    'a cycle, the whole line will be cleared and the RecultsPrevieCalculation will be called


    Dim Code As String
    Dim NextLine As Long
    Dim Plot_Sh_Name As String
    Dim vbcomp As VBComponent
    
    Code = "Public Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
    Code = Code & "'Target is the range that has been changed" & vbCrLf
    Code = Code & "Dim cell As Range" & vbCrLf
    Code = Code & "Dim LastColumn As String" & vbCrLf
    Code = Code & "LastColumn = " & Chr(34) & Plot_LastColumn & Chr(34) & vbCrLf
    Code = Code & "If Target.Column = 1 Then" & vbCrLf
    Code = Code & "Application.EnableEvents = False" & vbCrLf
    Code = Code & "Application.ScreenUpdating = False" & vbCrLf
    Code = Code & "else end" & vbCrLf
    Code = Code & "For Each cell In Target" & vbCrLf
    Code = Code & "Me.Range(cell, Me.Range(LastColumn & cell.Row)).ClearContents" & vbCrLf
    Code = Code & "Next" & vbCrLf
    Code = Code & "Application.EnableEvents = True" & vbCrLf
    Code = Code & "Application.ScreenUpdating = True" & vbCrLf
    Code = Code & "End If" & vbCrLf
    Code = Code & "Application.Run " & Chr(34) & "Chronus_1.2.0.xlam!resultsPreviewCalculation" & Chr(34) & vbCrLf
    Code = Code & "End Sub"
    
    For Each vbcomp In Plot_Sh.Parent.VBProject.VBComponents
        If vbcomp.Name = Plot_Sh.CodeName Then
            Plot_Sh_Name = vbcomp.Name
                Exit For
        End If
    Next
    
    With Plot_Sh.Parent.VBProject.VBComponents(Plot_Sh_Name).CodeModule
        NextLine = .CountOfLines + 1
        .InsertLines NextLine, Code
    End With
End Sub

Sub Plot_PlotAnalysis(Sh As Worksheet, Optional Plot64 As Boolean = True, Optional Plot74 As Boolean = True, _
Optional Plot28 As Boolean = True, Optional Plot75 As Boolean = True, Optional Plot68 As Boolean = True, _
Optional Plot76 As Boolean = True, Optional PlotRawSignal As Boolean = True)
    'This program is based on an example of code from the book VBA programming for dummies.
    'It should plot data present in Plot sheets.
    
    'The six optional variables are used to set if which ratios should be plotted.
    
    'Updated 29092015 - Boolean variables were included so it is possible to choose which ratios should be plotted.
    
    Dim ChartShape_RawSignal As Shape
    Dim ChartShape_64 As Shape
    Dim ChartShape_74 As Shape
    Dim ChartShape_28 As Shape
    Dim ChartShape_75 As Shape
    Dim ChartShape_68 As Shape
    Dim ChartShape_76 As Shape
    
    Dim NewChart_RawSignal As Chart
    Dim NewChart_64 As Chart
    Dim NewChart_74 As Chart
    Dim NewChart_28 As Chart
    Dim NewChart_75 As Chart
    Dim NewChart_68 As Chart
    Dim NewChart_76 As Chart
    
    Dim Cht As ChartObject
    Dim ch As Chart
    Dim a As Series
    Dim ShapesInSheet As Shape
    
    Dim MinimumScaleRange As Range
    Dim MaximumScaleRange As Range
    
    If RawNumberCycles_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    'Advises the user that there charts in the sheet that will be deleted.
    If Sh.Shapes.count <> 0 Then
        If MsgBox("There are charts in " & Sh.Name & ". They will be deleted, would you like to continue?", vbYesNo) = vbNo Then
            End
        Else
            For Each ShapesInSheet In Sh.Shapes
                ShapesInSheet.Delete
            Next
        End If
    End If

    With Sh
        Set ChartDataX = .Range(Plot_ColumnCyclesTime & Plot_HeaderRow + 1, Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb)
        
        'RawSignal
        If PlotRawSignal = True Then
            Set ChartDataY_RawSignal2 = .Range(Plot_Column2 & Plot_HeaderRow + 1, Plot_Column2 & Plot_HeaderRow + RawNumberCycles_UPb)
            Set ChartDataY_RawSignal4 = .Range(Plot_Column4 & Plot_HeaderRow + 1, Plot_Column4 & Plot_HeaderRow + RawNumberCycles_UPb)
            Set ChartDataY_RawSignal6 = .Range(Plot_Column6 & Plot_HeaderRow + 1, Plot_Column6 & Plot_HeaderRow + RawNumberCycles_UPb)
            Set ChartDataY_RawSignal7 = .Range(Plot_Column7 & Plot_HeaderRow + 1, Plot_Column7 & Plot_HeaderRow + RawNumberCycles_UPb)
            Set ChartDataY_RawSignal8 = .Range(Plot_Column8 & Plot_HeaderRow + 1, Plot_Column8 & Plot_HeaderRow + RawNumberCycles_UPb)
            Set ChartDataY_RawSignal32 = .Range(Plot_Column32 & Plot_HeaderRow + 1, Plot_Column32 & Plot_HeaderRow + RawNumberCycles_UPb)
            Set ChartDataY_RawSignal38 = .Range(Plot_Column38 & Plot_HeaderRow + 1, Plot_Column38 & Plot_HeaderRow + RawNumberCycles_UPb)
                Set ChartShape_RawSignal = .Shapes.AddChart
                    Set NewChart_RawSignal = ChartShape_RawSignal.Chart
        End If
        
        'Ratio 68
        If Plot68 = True Then
            Set ChartDataY_68 = .Range(Plot_Column68 & Plot_HeaderRow + 1, Plot_Column68 & Plot_HeaderRow + RawNumberCycles_UPb)
                Set ChartShape_68 = .Shapes.AddChart
                    Set NewChart_68 = ChartShape_68.Chart
        End If
        
        'Ratio 76
        If Plot76 = True Then
            Set ChartDataY_76 = .Range(Plot_Column76 & Plot_HeaderRow + 1, Plot_Column76 & Plot_HeaderRow + RawNumberCycles_UPb)
                Set ChartShape_76 = .Shapes.AddChart
                    Set NewChart_76 = ChartShape_76.Chart
        End If
        
        'Ratio 64
        If Plot64 = True Then
            Set ChartDataY_64 = .Range(Plot_Column64 & Plot_HeaderRow + 1, Plot_Column64 & Plot_HeaderRow + RawNumberCycles_UPb)
                Set ChartShape_64 = .Shapes.AddChart
                    Set NewChart_64 = ChartShape_64.Chart
        End If
                
        'Ratio 74
        If Plot74 = True Then
            Set ChartDataY_74 = .Range(Plot_Column74 & Plot_HeaderRow + 1, Plot_Column74 & Plot_HeaderRow + RawNumberCycles_UPb)
                Set ChartShape_74 = .Shapes.AddChart
                    Set NewChart_74 = ChartShape_74.Chart
        End If
            
        'Ratio 28
        If Isotope232analyzed = True And Plot28 = True Then
            Set ChartDataY_28 = .Range(Plot_Column28 & Plot_HeaderRow + 1, Plot_Column28 & Plot_HeaderRow + RawNumberCycles_UPb)
                Set ChartShape_28 = .Shapes.AddChart
                    Set NewChart_28 = ChartShape_28.Chart
        End If
        
        'Ratio 75
        If Plot75 = True Then
            Set ChartDataY_75 = .Range(Plot_Column75 & Plot_HeaderRow + 1, Plot_Column75 & Plot_HeaderRow + RawNumberCycles_UPb)
                Set ChartShape_75 = .Shapes.AddChart
                    Set NewChart_75 = ChartShape_75.Chart
        End If
                
                
        'The lines below will delete all series automatically created in the charts created above
        For Each ShapesInSheet In .Shapes
            For Each a In ShapesInSheet.Chart.SeriesCollection
                    a.Delete
            Next
        Next
    
    End With

    If PlotRawSignal = True Then
        With NewChart_RawSignal
        
            .HasTitle = True
            .ChartTitle.Text = "Signal less blank"
            
            .Axes(xlValue).ScaleType = xlLogarithmic 'Logarithmic scale
                
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            .SeriesCollection.NewSeries
            
            For Each a In NewChart_RawSignal.SeriesCollection
                    a.XValues = ChartDataX
            Next
            
            .SeriesCollection(1).Values = ChartDataY_RawSignal2
                .SeriesCollection(1).Name = "Hg 202 (cps)"
            .SeriesCollection(2).Values = ChartDataY_RawSignal4
                .SeriesCollection(2).Name = "Pb 204 (cps)"
            .SeriesCollection(3).Values = ChartDataY_RawSignal6
                .SeriesCollection(3).Name = "Pb 206 (cps)"
            .SeriesCollection(4).Values = ChartDataY_RawSignal7
                .SeriesCollection(4).Name = "Pb 207 (cps)"
            
            If Isotope208analyzed = True Then
                .SeriesCollection(5).Values = ChartDataY_RawSignal8
            End If
                
                .SeriesCollection(5).Name = "Pb 208 (cps)"
             
             If Isotope232analyzed = True Then
                .SeriesCollection(6).Values = ChartDataY_RawSignal32
            End If
                
                .SeriesCollection(6).Name = "Th 232 (cps)"
            .SeriesCollection(7).Values = ChartDataY_RawSignal38
                .SeriesCollection(7).Name = "U 238 (cps)"
            
            .Parent.Top = 100
            
        End With
    End If
    
    If Plot64 = True Then
        With NewChart_64
            
            .Legend.Delete
            
            .SeriesCollection.NewSeries
                .SeriesCollection(1).XValues = ChartDataX
                    .SeriesCollection(1).Values = ChartDataY_64
                        .SeriesCollection(1).Name = "Ratio 64"
        End With
    End If
    
    If Plot74 = True Then
        With NewChart_74
            
            .Legend.Delete
            
            .SeriesCollection.NewSeries
                .SeriesCollection(1).XValues = ChartDataX
                    .SeriesCollection(1).Values = ChartDataY_74
                        .SeriesCollection(1).Name = "Ratio 74"
        End With
    End If

    If Plot28 = True Then
        With NewChart_28
        
            .Legend.Delete
              
            .SeriesCollection.NewSeries
                .SeriesCollection(1).XValues = ChartDataX
                
                    If Isotope232analyzed = True Then
                        .SeriesCollection(1).Values = ChartDataY_28
                    End If
                        .SeriesCollection(1).Name = "Ratio 28"
        End With
    End If
    
    If Plot75 = True Then
        With NewChart_75
            
            .Legend.Delete
            
            .SeriesCollection.NewSeries
                .SeriesCollection(1).XValues = ChartDataX
                    .SeriesCollection(1).Values = ChartDataY_75
                        .SeriesCollection(1).Name = "Ratio 75"
        End With
    End If
    
    If Plot68 = True Then
        With NewChart_68
        
            .Legend.Delete
            
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .XValues = ChartDataX
                    .Values = ChartDataY_68
                        .Name = "Ratio 68"
                        
                'lines to add a trendline to the chart
                .Trendlines.Add
                    With .Trendlines(1)
                        .DisplayEquation = True
                        .DisplayRSquared = True
                        .DataLabel.Left = 250
                        .DataLabel.Top = 130
                        .DataLabel.Format.TextFrame2.TextRange.Font.Bold = msoTrue
                        .DataLabel.Format.TextFrame2.TextRange.Font.Size = 16
                    
                    End With
                
            End With
    
        End With
    End If
    
    If Plot76 = True Then
        With NewChart_76
            
            .Legend.Delete
            
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .XValues = ChartDataX
                    .Values = ChartDataY_76
                        .Name = "Ratio 76"
    
                'lines to add a trendline to the chart
                .Trendlines.Add
                    With .Trendlines(1)
                        .DisplayEquation = True
                        .DisplayRSquared = True
                        .DataLabel.Left = 250
                        .DataLabel.Top = 130
                        
                        .DataLabel.Format.TextFrame2.TextRange.Font.Bold = msoTrue
                        .DataLabel.Format.TextFrame2.TextRange.Font.Size = 16
                    
                    End With
            
            End With
    
        End With
    End If
    
    For Each Cht In Sh.ChartObjects
        Cht.Chart.Type = xlXYScatter
    Next Cht

    'The lines below will make series lines invisible and set minimum and maximum scale
    For Each ShapesInSheet In Sh.Shapes
            
        For Each a In ShapesInSheet.Chart.SeriesCollection
            a.Format.Line.Visible = msoFalse
        Next

        With ShapesInSheet.Chart
            Set MinimumScaleRange = Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + 1)
            Set MaximumScaleRange = Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb)
            
            .Axes(xlValue).Crosses = xlMaximum
            
            If Not IsEmpty(MinimumScaleRange) = True Then
                .Axes(xlCategory).MinimumScale = MinimumScaleRange
            End If
            
            If Not IsEmpty(MaximumScaleRange) = True Then
                .Axes(xlCategory).MaximumScale = MaximumScaleRange
            End If
            
'                .Axes(xlValue).Format.TextFrame2.TextRange.Font.Bold = msoTrue
'                .Axes(xlValue).Format.TextFrame2.TextRange.Font.Bold = msoTrue
        End With
    Next
    

End Sub

Public Sub ResultsPreviewCalculation()
    
    'This program will update the results in the preview box

    Dim ChartDataX As Range
'    Dim ChartDataY_RawSignal2 As Range
'    Dim ChartDataY_RawSignal4 As Range
'    Dim ChartDataY_RawSignal6 As Range
'    Dim ChartDataY_RawSignal7 As Range
'    Dim ChartDataY_RawSignal8 As Range
'    Dim ChartDataY_RawSignal32 As Range
'    Dim ChartDataY_RawSignal38 As Range
    Dim ChartDataY_68 As Range
    Dim ChartDataY_76 As Range
'    Dim ChartDataY_64 As Range
'    Dim ChartDataY_74 As Range
'    Dim ChartDataY_28 As Range
'    Dim ChartDataY_75 As Range
    
    Dim Preview_Ratio68 As Range
    Dim Preview_Ratio68ErrorAbs As Range
    Dim Preview_Ratio68ErrorRelative As Range
    Dim Preview_Ratio68R As Range
    Dim Preview_Ratio68R2 As Range
    
    Dim Preview_Ratio76 As Range
    Dim Preview_Ratio76ErrorAbs As Range
    Dim Preview_Ratio76ErrorRelative As Range
    Dim Preview_Ratio76R As Range
    Dim Preview_Ratio76R2 As Range
    
    If SpotRaster_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    If Plot_Sh Is Nothing Then
        Set Plot_Sh = ActiveSheet
            Call CheckPlotSheet(Plot_Sh)
    End If
        
    With Plot_Sh
    
        Set ChartDataX = .Range(Plot_ColumnCyclesTime & Plot_HeaderRow + 1, Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb)

'        Set ChartDataY_RawSignal2 = .Range(Plot_Column2 & Plot_HeaderRow + 1, Plot_Column2 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_RawSignal4 = .Range(Plot_Column4 & Plot_HeaderRow + 1, Plot_Column4 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_RawSignal6 = .Range(Plot_Column6 & Plot_HeaderRow + 1, Plot_Column6 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_RawSignal7 = .Range(Plot_Column7 & Plot_HeaderRow + 1, Plot_Column7 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_RawSignal8 = .Range(Plot_Column8 & Plot_HeaderRow + 1, Plot_Column8 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_RawSignal32 = .Range(Plot_Column32 & Plot_HeaderRow + 1, Plot_Column32 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_RawSignal38 = .Range(Plot_Column38 & Plot_HeaderRow + 1, Plot_Column38 & Plot_HeaderRow + RawNumberCycles_UPb)
        
        Set ChartDataY_68 = .Range(Plot_Column68 & Plot_HeaderRow + 1, Plot_Column68 & Plot_HeaderRow + RawNumberCycles_UPb)
        Set ChartDataY_76 = .Range(Plot_Column76 & Plot_HeaderRow + 1, Plot_Column76 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_64 = .Range(Plot_Column64 & Plot_HeaderRow + 1, Plot_Column64 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_74 = .Range(Plot_Column74 & Plot_HeaderRow + 1, Plot_Column74 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_28 = .Range(Plot_Column28 & Plot_HeaderRow + 1, Plot_Column28 & Plot_HeaderRow + RawNumberCycles_UPb)
'        Set ChartDataY_75 = .Range(Plot_Column75 & Plot_HeaderRow + 1, Plot_Column75 & Plot_HeaderRow + RawNumberCycles_UPb)


        'The following lines must be changed if the Plot_ResultsPreview constant be changed
        Set Preview_Ratio68 = .Range("T4")
        Set Preview_Ratio68ErrorAbs = .Range("U4")
        Set Preview_Ratio68ErrorRelative = .Range("V4")
        Set Preview_Ratio68R = .Range("W4")
        Set Preview_Ratio68R2 = .Range("X4")
    
        Set Preview_Ratio76 = .Range("T5")
        Set Preview_Ratio76ErrorAbs = .Range("U5")
        Set Preview_Ratio76ErrorRelative = .Range("V5")
        Set Preview_Ratio76R = .Range("W5")
        Set Preview_Ratio76R2 = .Range("X5")
        ''''''''''''''''''''''''''''''''''''''''''''''''
    End With

    Select Case SpotRaster_UPb
        Case "Spot"
            'Intercept of 68 trend
            Preview_Ratio68 = WorksheetFunction.Intercept(ChartDataY_68, ChartDataX)
                '68 intercept error multiplied by student´s t factor for 68% confidence
                Preview_Ratio68ErrorAbs = LineFitInterceptError(ChartDataY_68, ChartDataX) * _
                    WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(ChartDataY_68) - 2)
            
            'R
            Preview_Ratio68R = WorksheetFunction.Pearson(ChartDataY_68, ChartDataX)
            'R2
            Preview_Ratio68R2 = WorksheetFunction.Power(Preview_Ratio68R, 2)
                
        Case "Raster"
            '68 average
            Preview_Ratio68 = WorksheetFunction.Average(ChartDataY_68)
            
                '68 average error propagation multiplied by student´s t factor for 68% confidence
                Preview_Ratio68ErrorAbs = (WorksheetFunction.StDev_S(ChartDataY_68) / Sqr(WorksheetFunction.count(ChartDataY_68)) * _
                    WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(ChartDataY_68) - 1))
                                                            
            'R
            Preview_Ratio68R = WorksheetFunction.Pearson(ChartDataY_68, ChartDataX)
            'R2
            Preview_Ratio68R2 = WorksheetFunction.Power(Preview_Ratio68R, 2)
    End Select
         
        '68 relative error
        Preview_Ratio68ErrorRelative = 100 * Preview_Ratio68ErrorAbs / Preview_Ratio68

    
    'Intercept of 76 trend
    Preview_Ratio76 = WorksheetFunction.Intercept(ChartDataY_76, ChartDataX)
                                                          
    '76 error
    Preview_Ratio76ErrorAbs = LineFitInterceptError(ChartDataY_76, ChartDataX) * WorksheetFunction.T_Inv_2T(ConfLevel, WorksheetFunction.count(ChartDataY_76) - 2)
    'R
    Preview_Ratio76R = WorksheetFunction.Pearson(ChartDataY_76, ChartDataX)
    'R2
    Preview_Ratio76R2 = WorksheetFunction.Power(Preview_Ratio76R, 2)
    
        '76 relative error
        Preview_Ratio76ErrorRelative = 100 * Preview_Ratio76ErrorAbs / Preview_Ratio76

End Sub

Sub LineUpMyCharts(Sh As Worksheet, Optional MainChart As Integer)

    'Modified from http://peltiertech.com/Excel/ChartsHowTo/ResizeAndMoveAChart.html
    
    Dim MyWidth As Single, MyHeight As Single
    Dim MyWidthMainChart As Single, MyHeightMainChart As Single
    Dim NumWide As Long
    Dim iChtIx As Long, iChtCt As Long
    Dim a As Integer
    Dim B As Integer

    If mwbk Is Nothing Then
        Call PublicVariables
    End If
    
    
    MyWidth = Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow).End(xlToRight).Offset(, 1).Left / 2
    MyHeight = 200
    
    MyWidthMainChart = Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow).End(xlToRight).Offset(, 1).Left - _
        Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow).Offset(, 1).Left '2 * MyWidth
    
    MyHeightMainChart = 2 * MyHeight
    
    If MainChart < 0 Then
        MsgBox ("MainChart must not be an integer < 0.")
            End
    End If

    NumWide = 2
    
        a = 1
        B = 0

    iChtCt = Sh.ChartObjects.count
    For iChtIx = 1 To iChtCt
    
        Sh.ChartObjects(iChtIx).Placement = xlFreeFloating
        
        If iChtIx = MainChart Then
            With Sh.ChartObjects(iChtIx)
                .Width = MyWidthMainChart
                .Height = MyHeightMainChart
                .Left = Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb).Offset(, 1).Left
                .Top = Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb + 2).Top
            End With
            
            a = 0
            B = 1
        Else
            With Sh.ChartObjects(iChtIx)
                .Width = MyWidth
                .Height = MyHeight
                
                    .Left = ((iChtIx - a) Mod NumWide) * MyWidth + Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow).Offset(, 1).Left * 2
                    '.Left = ((iChtIx - a) Mod NumWide) * MyWidth + MyWidthMainChart + Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow).Offset(, 1).Left
                    .Top = Int(((iChtIx - a) / NumWide - B)) * MyHeight + Sh.Range(Plot_LastColumn & Plot_HeaderRow + 1).Top * 2 ' The second part of this formula is the top of the second cell
                    
            End With
        End If
    Next
    
    On Error GoTo 0
    On Error Resume Next
        Application.GoTo Plot_Sh.Range("A1")
        If Err.Number = 0 Then
            On Error GoTo 0
                Call SampleNameTxtBox
                    Call AddIgnoreSplButton
                        Call ResultsPreviewCalculation
                            Call FormatPlot(Plot_Sh)
        End If
    On Error GoTo 0
    
End Sub

Sub Plot_CopyData(Source_Sh As Worksheet, Destination_Sh As Worksheet)

    'This procedure copies isotopes signal from the Source_Sh (analyses data files) to the Destination_Sh (Plot_sh and
    'Plot_ShHidden.
    
    Application.GoTo Source_Sh.Range("A1")

    With Source_Sh
            .Range(RawCyclesTimeRange).Copy Destination_Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + 1)
            .Range(RawHg202Range).Copy Destination_Sh.Range(Plot_Column2 & Plot_HeaderRow + 1)
            .Range(RawPb204Range).Copy Destination_Sh.Range(Plot_Column4 & Plot_HeaderRow + 1)
            .Range(RawPb206Range).Copy Destination_Sh.Range(Plot_Column6 & Plot_HeaderRow + 1)
            .Range(RawPb207Range).Copy Destination_Sh.Range(Plot_Column7 & Plot_HeaderRow + 1)
            
            If Isotope208analyzed = True Then
                .Range(RawPb208Range).Copy Destination_Sh.Range(Plot_Column8 & Plot_HeaderRow + 1)
            End If
            
            If Isotope232analyzed = True Then
                .Range(RawTh232Range).Copy Destination_Sh.Range(Plot_Column32 & Plot_HeaderRow + 1)
            End If
            
            .Range(RawU238Range).Copy Destination_Sh.Range(Plot_Column38 & Plot_HeaderRow + 1)
    End With
    
    

End Sub

Sub Plot_OrdinaryCalculations(Sh As Worksheet)

    'This procedure copies and copies dividing isotopes signal to other columns in the same worksheet creating the ratios.
    
    Dim FindCells As Object
    Dim FirstAddress As String
    Dim B As Range 'used to loop through all cells in Plot_Column75 cells

    'Application.Goto Sh.Range("A1")

    With Sh
        
        '64 ratio
        .Range(Plot_Column6 & Plot_HeaderRow + 1, Plot_Column6 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy _
            Destination:=.Range(Plot_Column64 & Plot_HeaderRow + 1)
        
        .Range(Plot_Column4 & Plot_HeaderRow + 1, Plot_Column4 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy
            .Range(Plot_Column64 & Plot_HeaderRow + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide
                    
        '74 ratio
        .Range(Plot_Column7 & Plot_HeaderRow + 1, Plot_Column7 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy _
            Destination:=.Range(Plot_Column74 & Plot_HeaderRow + 1)
        
        .Range(Plot_Column4 & Plot_HeaderRow + 1, Plot_Column4 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy
            .Range(Plot_Column74 & Plot_HeaderRow + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide
        
        '28 ratio
        If Isotope232analyzed = True Then
            .Range(Plot_Column32 & Plot_HeaderRow + 1, Plot_Column32 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy _
                Destination:=.Range(Plot_Column28 & Plot_HeaderRow + 1)
            
            .Range(Plot_Column38 & Plot_HeaderRow + 1, Plot_Column38 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy
                .Range(Plot_Column28 & Plot_HeaderRow + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide
        End If
        
        '76 ratio
        .Range(Plot_Column7 & Plot_HeaderRow + 1, Plot_Column7 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy _
            Destination:=.Range(Plot_Column76 & Plot_HeaderRow + 1)
        
        .Range(Plot_Column6 & Plot_HeaderRow + 1, Plot_Column6 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy
            .Range(Plot_Column76 & Plot_HeaderRow + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide
        
        '68 ratio
        .Range(Plot_Column6 & Plot_HeaderRow + 1, Plot_Column6 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy _
            Destination:=.Range(Plot_Column68 & Plot_HeaderRow + 1)
        
        .Range(Plot_Column38 & Plot_HeaderRow + 1, Plot_Column38 & Plot_HeaderRow + Val(RawNumberCycles_UPb)).Copy
            .Range(Plot_Column68 & Plot_HeaderRow + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlDivide
            
        'Below, 75 ratios will be calculated based on 76 and 68 ratios
        For Each B In .Range(Plot_Column75 & Plot_HeaderRow + 1, Plot_Column75 & Plot_HeaderRow + Val(RawNumberCycles_UPb))
            
            If Not IsEmpty(.Range(Plot_Column68 & B.Row)) = True Or Not IsEmpty(.Range(Plot_Column76 & B.Row)) = True Then
                
                B = .Range(Plot_Column68 & B.Row) * .Range(Plot_Column76 & B.Row) * RatioUranium_UPb
            
            End If
        
        Next
        
        Application.CutCopyMode = False
    
        With Sh.Cells
            'When some of the cycles are ignored, ratios that use these cycles as denomitors will raise #DIV/0!
            Set FindCells = .Find("#DIV/0!", LookIn:=xlValues)
            If Not FindCells Is Nothing Then
                FirstAddress = FindCells.Address
                Do
                    FindCells.ClearContents
    
                    Set FindCells = .FindNext(FindCells)
                Loop While Not FindCells Is Nothing 'Or FindCells.Address <> FirstAddress
            End If
            
        End With

    End With

End Sub

Sub Plot_ClosePlot(Plot_Sh As Worksheet, Optional ShowRecalcMsg As Boolean = True)

    'This procedure calls the WriteCycles procedure to save the selected cycles of the analyis,
    'the FormatSamList and then deletes the plot_sh and the plot_shhidden.

    Dim FindIDObj As Object
    Dim AnalysisID As Integer
    
    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If
    
    Call CheckPlotSheet(Plot_Sh)

    AnalysisID = Plot_Sh.Range(Plot_IDCell)
    
    Set Plot_CyclesTimeRange = Plot_Sh.Range(Plot_ColumnCyclesTime & Plot_HeaderRow + 1, Plot_ColumnCyclesTime & Plot_HeaderRow + RawNumberCycles_UPb)
    
    Set FindIDObj = SamList_Sh.Columns(SamList_Sh.Range(SamList_ID & ":" & SamList_ID).Column).Find(AnalysisID)

    Call WriteCycles(Plot_CyclesTimeRange, SamList_Sh.Range(FindIDObj.Address).Row)
    
    Call FormatSamList
    
    Application.DisplayAlerts = False
        Plot_Sh.Delete
        Plot_ShHidden.Delete
    Application.DisplayAlerts = True
    
    If ShowRecalcMsg = True Then
        If MsgBox("You must run the full data reduction process again to your cycles selection " & _
            "be effective. Would you like to do it know?" & vbNewLine & _
            "A good tip is to select cycles for all samples before calculate all of them " & _
            "again.", vbYesNo) = vbYes Then
            
                If MsgBox("Depending on the number of analysis, this must take a long time to complete. " & _
                "Do you still want to proceed?", vbYesNo) = vbYes Then
                
                    Application.ScreenUpdating = False
                    
                    Call CalcBlank
                        Call CalcAllSlpStd_BlkCorr
                            Call CalcAllSlp_StdCorr
                                Call FormatSamList
                                    Call FormatStartANDOptions
                                        Call FormatSlpStdBlkCorr
                                            Call FormatSlpStdCorr
                                                Call FormatBlkCalc
                                                
                    Application.ScreenUpdating = True
                
                End If
                
        End If
    End If
    
End Sub

Sub RestoreOriginalPlotData(Plot_Sh As Worksheet)
   
    Dim SourceRange As Range
    Dim DestinationRange As Range
   
    Application.ScreenUpdating = False
   
    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If
    
    Call CheckPlotSheet(Plot_Sh) 'check if the activesshet is a valid plot sheet
    
    Set SourceRange = Plot_ShHidden.Range(Plot_FirstColumn & Plot_HeaderRow + 1, Plot_LastColumn & Plot_HeaderRow + RawNumberCycles_UPb)
    Set DestinationRange = Plot_Sh.Range(Plot_FirstColumn & Plot_HeaderRow + 1)

    SourceRange.Copy Destination:=DestinationRange
    
    Call FormatPlot(Plot_Sh)
    
End Sub

Sub CheckPlotSheet(Plot_Sh As Worksheet)

    If SamList_Sh Is Nothing Then
        Call PublicVariables
    End If
    
    'The lines below will check if the activesshet is a valid plot sheet. This is done trying to set Plot_ShHidden to
    'a sheet that has a similar name to a valid plot sheet. The only difference is that it has the string "Hidden".
    On Error Resume Next
        Set Plot_ShHidden = mwbk.Worksheets(Plot_Sh.Name & "Hidden")
        
        If Err.Number = 9 Then
            MsgBox "This is not a valid plot worksheet."
                Call UnloadAll
                    End
        ElseIf Err.Number <> 0 Then
            MsgBox "An error occured!"
                Call UnloadAll
                    End
        End If
    On Error GoTo 0


End Sub

Sub SampleNameTxtBox()

    'Created 18122015
    'This procedure creates a new shape in a plot sheet to highlight the name and id of the analysis being ploted.
    
    Dim SplNameLabel As Shape
    Dim SplNameLen As Long
    Dim AnalysisName As String
    Dim AnalysisID As Long
    Dim ShapeWidthSplName As Long
    
    AnalysisName = Plot_Sh.Range(Plot_AnalysisName)
    SplNameLen = Len(AnalysisName)
    AnalysisID = Plot_Sh.Range(Plot_IDCell)
    
    If Plot_Sh Is Nothing Then
        MsgBox "Unable to add textbox with sample name."
            Exit Sub
    End If

    Set SplNameLabel = Plot_Sh.Shapes.AddTextbox(msoTextOrientationHorizontal, 800, _
        200, 100, 100)
    
    With SplNameLabel.TextFrame2
        .TextRange.Characters.Text = AnalysisName
        .TextRange.Font.Size = 36
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = False
        .AutoSize = msoAutoSizeShapeToFitText
        
    End With

End Sub

Sub AddIgnoreSplButton()

    'Created 18122015
    'To be continued

    Dim IgnoreButtonForm As Button
    Dim CallingProcedure As String
    
    Call PublicVariables
    
    If Plot_Sh Is Nothing Then
        MsgBox "Unable to add Ignore button."
            Exit Sub
    End If

    Set IgnoreButtonForm = Plot_Sh.Buttons.Add(1020, 100, 150, 50)
    
    IgnoreButtonForm.Characters.Text = "Ignore Analysis"
    
    With IgnoreButtonForm.Font
        
        .Name = "Calibri"
        .FontStyle = "Bold"
        .Size = 20
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
        
    End With
    
'    CallingProcedure = "'" & TW.Name & "'!IgnoreAnalysis"
'
'    IgnoreButtonForm.OnAction = CallingProcedure
'    MsgBox IgnoreButtonForm.OnAction
    
End Sub


Sub ChangeChartTitleToSampleName()

    'Procedure used to change the title of the selected chart to the name of the sample stored in a valid Chronus workbook.

    'Created 01092015
    
    Dim Chrt As Chart

    Call PublicVariables 'Always used to check if the activeworkbook is valid Chronus workbook
    
    Set Chrt = Application.ActiveChart
    
    If Chrt Is Nothing Then 'The user must have selected a chart in order to run this procedure
    
        MsgBox "Please, select a chart.", vbOKOnly
            Application.DisplayAlerts = False
                Call UnloadAll
                    Application.DisplayAlerts = True
                        End
                        
    End If
    
    If SampleName_UPb = "" Then 'The name of the sample obviously must not be empty, so it is checked here
    
        If MsgBox("The name of the sample is missing. Would you like to update it?", vbYesNo) = vbYes Then
            MsgBox "After updating the sample name, try changing the chart name again."
                Box1_Start.Show
        Else
            Application.DisplayAlerts = False
                Call UnloadAll
                    Application.DisplayAlerts = False
                        End
        End If
        
    End If

    'If the user selected a chart and this is a valid Chronus workbook, then the code below will
    'change the title of this chart, positionate and format it.
    With Chrt
    
        .SetElement (msoElementChartTitleAboveChart) 'This command add a title to the chart
        
        With .ChartTitle
        
            .Text = SampleName_UPb.Value
            .Left = 207.499
            .Top = 50
            
            With .Format

                .TextFrame2.TextRange.Font.Bold = msoFalse
                .TextFrame2.TextRange.Font.Size = 14

                    With .Fill
                        .Visible = msoTrue
                        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
                        .ForeColor.TintAndShade = 0
                        .ForeColor.Brightness = 0
                        .Transparency = 0
                        .Solid
                    End With
            End With
        End With
    End With

End Sub


