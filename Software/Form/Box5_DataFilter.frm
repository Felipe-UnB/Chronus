VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box5_DataFilter 
   Caption         =   "Data filtering"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11490
   OleObjectBlob   =   "Box5_DataFilter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Box5_DataFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************
'/----------------------------------------------------\
'|  Macro desenvolvida por: Felipe Valença de Oliveira|
'|  Laboratório de Geocronologia - UnB                |
'|  Primeira versão (v1): Janeiro - 2014              |
'\----------------------------------------------------/
'******************************************************


    'Variables used by Box5_DataFilter
    Dim Error75 As Single 'Criteria indicating by to user to ignore grains with 207/235 error ratios too big.
    Dim Rho As Single 'Minimum error concordance between ratios 206/238 and 207/235 indicated by user.
    Dim f206 As Single 'Maximum Common Pb contents acceptable indicated by user.
    Dim MinValue As Single 'Minimum concorcande indicated by user
    Dim MaxValue As Single 'Maximum concordance indicated by user
    Dim Age68Limit As Single 'Age 206/208 limit, used to determine which ages concordance will be
                             'considered (206/238 and 207/206 or 206/238 and 207/235)
    Dim Conc6875 As Double 'Concordance between ages 68 and 75 (68 is the numerator)
    Dim Conc6876 As Double 'Concordance between ages 68 and 76 (68 is the numerator)
            
    Dim Error75Range As Range 'Range with ratio 207/235 errors
    Dim Range75 As Range ' range with ratio 207/235
    Dim Range75LastCell As Range
    Dim RhoRange As Range 'Range with Rho values
    Dim f206Range As Range 'Range with f206 values
    Dim Age68Range As Range
    Dim AutoFilterRange As Range
    Dim AddedSufix As String

Private Sub CheckBox1_Error75_Click()
    
    If CheckBox1_Error75 = True Then
        Error75Entry.Enabled = True
            Error75Entry.Value = 5
                Error75Entry.BackColor = vbWhite
    Else
        Error75Entry.Enabled = False
            Error75Entry.Value = ""
                Error75Entry.BackColor = &H8000000F
    End If
        
End Sub

Private Sub CheckBox2_Rho_Click()

    If CheckBox2_Rho = True Then
        RhoEntry.Enabled = True
            RhoEntry.Value = 0.5
                RhoEntry.BackColor = vbWhite
    Else
        RhoEntry.Enabled = False
            RhoEntry.Value = ""
                RhoEntry.BackColor = &H8000000F
    End If

End Sub

Private Sub CheckBox3_F206_Click()

    If CheckBox3_F206 = True Then
        f206Entry.Enabled = True
            f206Entry.Value = 3
                f206Entry.BackColor = vbWhite
    Else
        f206Entry.Enabled = False
            f206Entry.Value = ""
                f206Entry.BackColor = &H8000000F
    End If

End Sub

Private Sub CheckBox4_AutoSelectAnalyses_Click()

    If CheckBox4_AutoSelectAnalyses = True Then
        
        With Me.TextBox1
            .Enabled = True
            .BackColor = vbWhite
            
            If Not IsEmpty(SelectedBins_UPb) = True Then
                .Value = SelectedBins_UPb.Value
            End If
            
        End With
        
        Me.CheckBox5_IgnoreSecStd.Enabled = False
        Me.AdvancedFilters.Value = True
        Me.TextBox1.SetFocus
        
    Else
        
        With Me.TextBox1
            .Enabled = False
            .Value = ""
            .BackColor = &H8000000F
        End With
        
        Me.CheckBox5_IgnoreSecStd.Enabled = False
    
    End If

End Sub


Private Sub IgneousRock_Click()
    
    TextBoxMin = -5
    TextBoxMax = 5
    TextBoxAge68 = 1000
    
End Sub

Private Sub SedimentaryRock_Click()
    
    TextBoxMin = -10
    TextBoxMax = 10
    TextBoxAge68 = 1000

End Sub

Private Sub TextBoxAge68_Change()
    If AdvancedFilters.Value = True Then
        If Not IsNumeric(TextBoxAge68.Text) Or f206Entry < 0 Then
            MsgBox "Please, enter numbers >1 only.", vbInformation
                TextBoxAge68.SelStart = 0
                    TextBoxAge68.SelLength = Len(TextBoxAge68.Text)
        End If
    End If

        
End Sub

Private Sub UserForm_Initialize()
    
'    If InstalledIsoplot = False Then
'        MsgBox "Isoplot 4.15 must be installed and enabled in this " & _
'            "computer in order to use Data Filter.", vbOKOnly
'            Call UnloadAll
'                End
'    End If
    If ActiveSheet.Name <> SlpStdCorr_Sh_Name Then
        If MsgBox("This tool is only applied to SlpStdCorr sheets created by Chronus. Would you like to proceed?", vbYesNo) = vbNo Then
            Call UnloadAll
                End
        End If
    End If
    
    Call PublicVariables
    
    SlpStdCorr_Sh.AutoFilterMode = False
    
    CheckBox1_Error75.Value = True
        CheckBox2_Rho.Value = True
            CheckBox3_F206.Value = True
    
    Error75Entry.Value = 5 'Default Error75Entry number.
        Error75Entry.Enabled = True
        
    f206Entry.Value = 3 'Default f206EntryEntry number.
        f206Entry.Enabled = True
        
    RhoEntry.Value = 0.5 'Default RhoEntry number.
        RhoEntry.Enabled = True
    
End Sub

Private Sub f206Entry_Change() 'f206Entry validation
    
    If CheckBox3_F206.Value = True Then
        If Not IsNumeric(f206Entry.Text) Or f206Entry < 0 Then
            MsgBox "Please, enter numbers >1 only.", vbInformation
                f206Entry.SelStart = 0
                    f206Entry.SelLength = Len(f206Entry.Text)
        End If
    End If
    
End Sub

Private Sub RhoEntry_Change() 'RhoEntry validation

    If CheckBox2_Rho.Value = True Then
        If Not IsNumeric(RhoEntry.Text) Or RhoEntry < 0 Then
            MsgBox "Please, enter numbers >1 only.", vbInformation
                RhoEntry.SelStart = 0
                    RhoEntry.SelLength = Len(RhoEntry.Text)
        End If
    End If
    
End Sub

Private Sub Error75Entry_Change() 'Error75Entry validation
    
    If CheckBox1_Error75.Value = True Then
        If Not IsNumeric(Error75Entry.Text) Or Error75Entry < 0 Then
            MsgBox "Please, enter numbers >1 only.", vbInformation
                Error75Entry.SelStart = 0
                    Error75Entry.SelLength = Len(Error75Entry.Text)
        End If
    End If
    
End Sub

Private Sub TextBoxMin_Change() 'TextBoxMin validation
    
    If AdvancedFilters.Value = True Then
        If Not IsNumeric(TextBoxMin.Text) Then
            MsgBox "Please, enter numbers only.", vbInformation
                TextBoxMin.SelStart = 0
                    TextBoxMin.SelLength = Len(TextBoxMin.Text)
        End If
    End If
    
End Sub

Private Sub TextBoxMax_Change() 'TextBoxMax validation
    
    If AdvancedFilters.Value = True Then
        If Not IsNumeric(TextBoxMax.Text) Then
            MsgBox "Please, enter numbers only.", vbInformation
                TextBoxMax.SelStart = 0
                    TextBoxMax.SelLength = Len(TextBoxMax.Text)
        End If
    End If
    
End Sub

Private Sub AdvancedFilters_Click() 'Option buttons enabled or not
    If AdvancedFilters.Value = True Then
               
        With IgneousRock
            .Enabled = True
            .Value = True
            
            TextBoxMin = -5
            TextBoxMax = 5
            TextBoxAge68 = 1000
        End With
        
        SedimentaryRock.Enabled = True
            
        TextBoxMin.Enabled = True
            TextBoxMax.Enabled = True
                TextBoxAge68.Enabled = True
    
        TextBoxMin.BackColor = vbWhite
            TextBoxMax.BackColor = vbWhite
                TextBoxAge68.BackColor = vbWhite

    Else
        IgneousRock.Enabled = False
            SedimentaryRock.Enabled = False
            
        With TextBoxMin
            .Value = ""
            .Enabled = False
            .BackColor = &H8000000F
        End With
        
        With TextBoxMax
            .Value = ""
            .Enabled = False
            .BackColor = &H8000000F
        End With
        
        With TextBoxAge68
            .Value = ""
            .Enabled = False
            .BackColor = &H8000000F
            End With
    End If
End Sub

Private Sub Ok_Click()
    
    Dim Count1 As Long
    Dim Count2 As Long
    Dim Count3 As Long
    Dim Count4 As Long
    Dim Count5 As Long
    Dim AnalysesTotal As Long 'Number of analyses
    
    Box5_DataFilter.Hide
    
    If SlpStdCorr_Sh Is Nothing Then
        Call PublicVariables
    End If
                
    SlpStdCorr_Sh.Cells.Font.Strikethrough = False
    
        Error75 = Error75Entry.Value
        Rho = RhoEntry.Value
        f206 = f206Entry.Value
        
        If AdvancedFilters.Value = True Then
            MinValue = TextBoxMin.Value
            MaxValue = TextBoxMax.Value
            Age68Limit = TextBoxAge68.Value
        End If
        
        With SlpStdCorr_Sh
            Set Error75Range = .Range(StdCorr_Column751Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column751Std & StdCorr_HeaderRow + 1).End(xlDown))
            Set Range75 = .Range(StdCorr_Column75 & StdCorr_HeaderRow + 1, .Range(StdCorr_Column75 & StdCorr_HeaderRow + 1).End(xlDown))
            Set Range75LastCell = .Range(StdCorr_Column75 & StdCorr_HeaderRow + 1).End(xlDown)
            Set RhoRange = .Range(StdCorr_Column7568Rho & StdCorr_HeaderRow + 1, .Range(StdCorr_Column7568Rho & StdCorr_HeaderRow + 1).End(xlDown))
            Set f206Range = .Range(StdCorr_ColumnF206 & StdCorr_HeaderRow + 1, .Range(StdCorr_ColumnF206 & StdCorr_HeaderRow + 1).End(xlDown))
            Set Age68Range = .Range(StdCorr_Column68AgeMa & StdCorr_HeaderRow + 1, .Range(StdCorr_Column68AgeMa & StdCorr_HeaderRow + 1).End(xlDown))
        End With
         
        Count1 = Error75Range.count
        Count2 = Range75.count
        Count3 = RhoRange.count
        Count4 = f206Range.count
        Count5 = Age68Range.count
         
        If _
            Count1 <> Count2 Or _
            Count1 <> Count3 Or _
            Count1 <> Count4 Or _
            Count1 <> Count5 _
            Then
                MsgBox ("The filter can not be applied, there are cells missing in SlpStdCorr_Sh.")
                    Call UnloadAll
                        End
        Else
            
            AnalysesTotal = Count1
            
        End If
            'Procedures below will check Ratio Error75, Rho and f206 based on criteria indicated by user.
                 
    
    Call FilterAnalysis(AnalysesTotal)

        Call AutoSelectAnalyses

    Set AutoFilterRange = SlpStdCorr_Sh.Range(StdCorr_FirstColumn & StdCorr_HeaderRow, StdCorr_LastColumn & StdCorr_HeaderRow)
'            AutoFilterRange.Select
        AutoFilterRange.AutoFilter
'            SlpStdCorr_Sh.AutoFilter = True '.AutoFilterMode = True
                 
    SlpStdCorr_Sh.Activate
                 
    Unload Box5_DataFilter
            
End Sub

Sub AutoSelectAnalyses()
    
    'This procedure tkaes the arguments selected by the user in Box5_DataFilter and
    'creates groups of analyses that agree with them, separating the analyses per bin
    
    'FUTURE IMPLEMENTATION - allow the procedure to ignore internal standards analyses
    
    Dim SelectedBins() As String
    Dim Counter1 As Long
    Dim Counter2 As Long
    Dim Counter3 As Long
    Dim CellInRange As Range
    Dim MinLimit As Double
    Dim MaxLimit As Double
    Dim ConsideredAge As Double 'Depending on the 68 age, the 68 or the 76 age will be usaed to compare with the bins
    Dim AgeConcordance As Double
    Dim AnalysesInAgreement As Long
    Dim AnalysisName As Range
    
    If Me.CheckBox4_AutoSelectAnalyses.Value = False Then
        Exit Sub
    End If
    
    AddedSufix = "-GROUP_"
    
    SelectedBins_UPb = Me.TextBox1.Value 'Bins will be stored in StartANDOptions sheet
    
    SelectedBins = Split(Me.TextBox1.Value, ";") 'Splits the string with n bins in a array with n elements

    If IsArrayEmpty(SelectedBins) = True Or UBound(SelectedBins) = 0 Then
        MsgBox "You must select at least 2 bins for the auto select procedure.", vbOKOnly
            Me.Show
    End If
    
    For Counter1 = LBound(SelectedBins) To UBound(SelectedBins)
        SelectedBins(Counter1) = Replace(SelectedBins(Counter1), " ", "") 'Removes any spaces from the array elements (all analyses names)
            If IsNumeric(SelectedBins(Counter1)) = False Then
                MsgBox "Only numbers >= 0 are accepted as bins. Please, check the bins.", vbOKOnly
                    Me.Show
            End If
    Next

    For Counter1 = LBound(SelectedBins) To UBound(SelectedBins)
        For Counter2 = Counter1 + 1 To UBound(SelectedBins)
            If SelectedBins(Counter1) = SelectedBins(Counter2) Then
                MsgBox "Bins are duplicated. Please, check them and then retry."
                    Me.Show
            End If
        Next
    Next
    
    Counter2 = 97
    Counter3 = 48

    Call RemoveSufix

    For Counter1 = LBound(SelectedBins) To UBound(SelectedBins)
        
        AnalysesInAgreement = 0
        
        If Counter1 <> UBound(SelectedBins) Then
            MinLimit = Val(SelectedBins(Counter1))
            MaxLimit = Val(SelectedBins(Counter1 + 1))
        Else
            Exit For
        End If
        
        For Each CellInRange In Age68Range
            
            Set AnalysisName = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellInRange.Row)
            
            If IsNumeric(CellInRange) = True And CellInRange.Font.Strikethrough = False Then
                If CellInRange <= Age68Limit Then
                    ConsideredAge = CellInRange.Value
                    AgeConcordance = SlpStdCorr_Sh.Range(StdCorr_Column6875Conc & CellInRange.Row)
                Else
                    ConsideredAge = SlpStdCorr_Sh.Range(StdCorr_Column76AgeMa & CellInRange.Row)
                    AgeConcordance = SlpStdCorr_Sh.Range(StdCorr_Column6876Conc & CellInRange.Row)
                End If

                If ConsideredAge >= MinLimit And ConsideredAge < MaxLimit Then
                    
                    With AnalysisName
                        
                        .Value = .Value & AddedSufix & Chr(Counter2) & Chr(Counter3) & "/" & MinLimit & "-" & MaxLimit
                            AnalysesInAgreement = AnalysesInAgreement + 1
                    End With
                
                End If
                
            End If
            
        Next
        
        If AnalysesInAgreement <> 0 Then
            Counter3 = Counter3 + 1
                If Counter3 = 58 Then
                    Counter2 = Counter2 + 1
                        If Counter2 = 122 Then
                            MsgBox "The number of bins exceed 225! This program will stop."
                                Exit Sub
                        End If
                        
                    Counter3 = 48
                End If
        End If
    Next

    SlpStdCorr_Sh.Range(StdCorr_SlpName & 1).EntireColumn.AutoFit

End Sub

Sub RemoveSufix()

    Dim AnalysisName As Range
    Dim CellInRange As Range
    Dim SearchStr As Long
    
    For Each CellInRange In Age68Range
        
        Set AnalysisName = SlpStdCorr_Sh.Range(StdCorr_SlpName & CellInRange.Row)
        
        SearchStr = InStr(1, AnalysisName, AddedSufix)
        
        If SearchStr <> 0 Then
            AnalysisName.Value = Left(AnalysisName.Value, SearchStr - 1)
        End If
        
    Next

End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
    'Code modified from http://www.cpearson.com/excel/TextBox.htm
    
    Dim InsertPosition As Long
    
    InsertPosition = Me.TextBox1.SelStart
    
    Select Case KeyAscii
    
        Case Asc("0") To Asc("9")
            
        Case Asc(".")
            
            If InsertPosition = 0 Then
                KeyAscii = 0
            Else
                
                If InsertPosition - 1 = 0 Then 'For the cases when the user is typing in the beginning of the textbox
                    InsertPosition = InsertPosition + 1

                        If IsNumeric(Mid(Me.TextBox1.Text, InsertPosition - 1, 1)) = False Then
                            KeyAscii = 0
                        End If
                
                Else

                    If IsNumeric(Mid(Me.TextBox1.Text, InsertPosition, 1)) = False Then
                        KeyAscii = 0
                    End If
                
                End If
                
            End If
            
        Case Asc(";")
        
            If InsertPosition = 0 Then
                KeyAscii = 0
            Else
                
                If InsertPosition - 1 = 0 Then 'For the cases when the user is typing in the beginning of the textbox
                    InsertPosition = InsertPosition + 1

                        If IsNumeric(Mid(Me.TextBox1.Text, InsertPosition - 1, 1)) = False Then
                            KeyAscii = 0
                        End If
                
                Else

                    If IsNumeric(Mid(Me.TextBox1.Text, InsertPosition, 1)) = False Then
                        KeyAscii = 0
                    End If
                
                End If
                
            End If
            
        Case Else
            KeyAscii = 0
    End Select
    
End Sub

Sub FilterAnalysis(AnalysesTotal As Long)

    Dim SearchStr As Long
    Dim Convert75Percent As Double
    Dim Count1 As Long
    Dim Count2 As Long
    Dim CutCellsRange As Range
    Dim PasteRange As Variant
    Dim CellInRange As Range 'Cells inside range that will be checked
    Dim RemoveCells() As Long
    Dim RemoveCellsRange As Range
    
    Dim Countf206 As Long
    Dim CountError75 As Long
    Dim CountRho As Long
    Dim CountConcordance As Long
    Dim CountBadResults As Long
    
    Dim FailedAnalyses As Long
    
        SearchStr = InStr(SlpStdCorr_Sh.Range(StdCorr_Column751Std & StdCorr_HeaderRow), "%")
        
        For Each CellInRange In Error75Range
            
            On Error Resume Next
            
                If SearchStr = 0 Then
                    Convert75Percent = 100 * (CellInRange / CellInRange.Offset(, -1))
                Else
                    Convert75Percent = CellInRange
                End If
                
                If IsNumeric(CellInRange) = False Then
                
                    CountBadResults = CountBadResults + 1
                    
                ElseIf Convert75Percent > Error75 Or CellInRange < 0 Then
                
                    Call FailAnalysisFilter(CellInRange)
                    'Call HighlightBorders(CellInRange, "BiggerSmaller", ">" & Trim(Str(Error75)))
                        CountError75 = CountError75 + 1
                End If
            
            On Error GoTo 0
            
        Next
        
            For Each CellInRange In RhoRange
            
                If IsNumeric(CellInRange) = False Then
                
                    CountBadResults = CountBadResults + 1
                    
                ElseIf CellInRange < Rho Or CellInRange < 0 Then
                
                    Call FailAnalysisFilter(CellInRange)
                        CountRho = CountRho + 1
                        
                End If

            Next
            
                 For Each CellInRange In f206Range
                 
                    If IsNumeric(CellInRange) = False Then
                    
                        CountBadResults = CountBadResults + 1
                    
                    ElseIf CellInRange > f206 Or CellInRange < 0 Then
                        
                        Call FailAnalysisFilter(CellInRange)
                            Countf206 = Countf206 + 1
                    
                    End If
                
                Next
                            
        'Considering that the type of rock is known, additional criteria will be used to filter the data.
        'The procecures below are resposinble for this.
            
        If IgneousRock = True Then 'If it's an igneous rock.
            
            On Error Resume Next
            
                For Each CellInRange In Age68Range
                        
                        If IsNumeric(CellInRange) = True Then
                        
                            If CellInRange <= Age68Limit Then
                                    
                                Conc6875 = SlpStdCorr_Sh.Range(StdCorr_Column6875Conc & CellInRange.Row)
                            
                                If Err.Number = 0 Then
                                
                                    Select Case Conc6875 '68-75 concordance was chosen
                                        
                                        Case Is < MinValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                        
                                        Case Is >= MaxValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                                
                                    End Select
                                    
                                Else
                                
                                    Call FailAnalysisFilter(CellInRange)
                                        CountBadResults = CountBadResults + 1
                                        
                                End If
                                                                        
                            Else
                                
                                Conc6876 = SlpStdCorr_Sh.Range(StdCorr_Column6876Conc & CellInRange.Row)
                                
                                If Err.Number = 0 Then
                                
                                    Select Case Conc6876 '68/76 concordance was chosen
                                        
                                        Case Is < MinValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                        
                                        Case Is >= MaxValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                    
                                    End Select
                                    
                                Else
                                
                                    Call FailAnalysisFilter(CellInRange)
                                        CountBadResults = CountBadResults + 1
                                
                                End If
                            
                            End If
                            
                        Else
                        
                            Call FailAnalysisFilter(CellInRange)
                                CountBadResults = CountBadResults + 1
                                    
                        End If
                    
                    Err.Clear
                    
                Next
                
            On Error GoTo 0
        
    ElseIf SedimentaryRock = True Then
            
            On Error Resume Next
            
                For Each CellInRange In Age68Range
                    
                        If IsNumeric(CellInRange) = True Then
                        
                            If CellInRange <= Age68Limit Then
                                
                                Conc6875 = SlpStdCorr_Sh.Range(StdCorr_Column6875Conc & CellInRange.Row)
                                
                                If Err.Number = 0 Then
                                
                                    Select Case Conc6875 '68-75 concordance was chosen
                                        
                                        Case Is < MinValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                    
                                        Case Is >= MaxValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                                
                                    End Select
                                                  
                                Else
                                
                                    CountBadResults = CountBadResults + 1
                                        
                                End If
                                                  
                            Else
        
                                Conc6876 = SlpStdCorr_Sh.Range(StdCorr_Column6876Conc & CellInRange.Row)
                                
                                If Err.Number = 0 Then
                                    
                                    Select Case Conc6876 '68/76 concordance was chosen
                                        
                                        Case Is < MinValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                        
                                        Case Is >= MaxValue
                                            Call FailAnalysisFilter(CellInRange)
                                                CountConcordance = CountConcordance + 1
                                                
                                    End Select
                                    
                                Else
                                    
                                    CountBadResults = CountBadResults + 1
                                        
                                End If
                                    
                            End If
                            
                        Else
                        
                            CountBadResults = CountBadResults + 1
                        
                        End If
                    
                    Err.Clear
                    
                Next
        
            On Error GoTo 0
            
        End If
        
        ReDim RemoveCells(1 To 1) As Long
        Count1 = 1
        
        'These following two blocks of code will send the bad analyses to the end of the list (first block) and
        'then delete the empty rows
        
        'Block1
        For Each CellInRange In Range75
            If CellInRange.Font.Strikethrough = True Then
                
                RemoveCells(Count1) = CellInRange.Row
                    ReDim Preserve RemoveCells(1 To UBound(RemoveCells) + 1) As Long
                        Count1 = Count1 + 1
                
                Set CutCellsRange = SlpStdCorr_Sh.Range(StdCorr_FirstColumn & CellInRange.Row, StdCorr_LastColumn & CellInRange.Row)
                                
                If CutCellsRange.Row = Range75LastCell.Row Then
                    Set PasteRange = CutCellsRange.Offset(1)
'                        CutCellsRange.Cut (PasteRange)
                            Exit For
                Else
                    Set PasteRange = CutCellsRange.End(xlDown).Offset(1)
                    PasteRange.Select
                        CutCellsRange.Cut (PasteRange)
                End If
            End If
        Next
        
        'Block 2
        Count1 = 1
        Count2 = 0
        
        If Not LBound(RemoveCells) = UBound(RemoveCells) Then
            For Count1 = 1 To UBound(RemoveCells) - 1
                    
                With SlpStdCorr_Sh
                    Set RemoveCellsRange = .Range(StdCorr_FirstColumn & RemoveCells(Count1) - Count2, StdCorr_LastColumn & RemoveCells(Count1) - Count2)
                End With

                If IsEmpty(RemoveCellsRange.Item(1)) = True Then
                    RemoveCellsRange.Delete (xlShiftUp)
                End If
                
                Count2 = Count2 + 1
                
            Next
        End If
        
        FailedAnalyses = UBound(RemoveCells) - 1
        
        Call ShowDataFilterResult(AnalysesTotal, FailedAnalyses, Countf206, CountError75, CountRho, CountConcordance)

End Sub

Private Sub FailAnalysisFilter(CellRange As Range)

    With CellRange.EntireRow.Font
        '.Color = -16776961
        '.TintAndShade = 0
        .Strikethrough = True
    End With
    
    With CellRange.Font
        '.Bold = True
        .Name = "Franklin Gothic Heavy"
        .Size = 11
        .Strikethrough = True
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

End Sub

Sub HighlightBorders(CellRange As Range, ConditionType As String, Condition As String, Optional CondValue1 = 0, Optional CondValue2 = 0)

    'Created 31012017
    
    'NOT USED
    
    'This program will change the borders of the indicated cell. The primary application of this
    'is to highlight those cells that failed the test applied by chronus on SlpStdCorr sheet.
    
    Select Case ConditionType
    
        Case "Between"
            CellRange.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                Formula1:="=" & Str(CondValue1), Formula2:="=" & Str(CondValue2)
        
        Case "BiggerSmaller"
            CellRange.FormatConditions.Add Type:=xlExpression, Formula1:="=" & CellRange.Address & Condition
    
    End Select
    
    CellRange.FormatConditions(CellRange.FormatConditions.count).SetFirstPriority
    CellRange.FormatConditions(1).StopIfTrue = False
    
    With CellRange.FormatConditions(1).Borders(xlLeft)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With CellRange.FormatConditions(1).Borders(xlRight)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With CellRange.FormatConditions(1).Borders(xlTop)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    With CellRange.FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub

Private Sub ShowDataFilterResult( _
    AnalysesTotal As Long, _
    FailedAnalyses As Long, _
    Countf206 As Long, _
    CountError75 As Long, _
    CountRho As Long, _
    CountConcordance As Long)
    
    'The arguments of this procedure are the number of analysis that failed the respective tests
    
    Dim Countf206Percent As Long
    Dim CountError75Percent As Long
    Dim CountRhoPercent As Long
    Dim CountConcordancePercent As Long
    
    
    If FailedAnalyses <> 0 Then
        Countf206Percent = 100 * Countf206 / FailedAnalyses
        CountError75Percent = 100 * CountError75 / FailedAnalyses
        CountRhoPercent = 100 * CountRho / FailedAnalyses
        CountConcordancePercent = 100 * CountConcordance / FailedAnalyses
    Else
        Countf206Percent = 0
        CountError75Percent = 0
        CountRhoPercent = 0
        CountConcordancePercent = 0
    End If
    
    Load Box8_DataFilterResult
    
    With Box8_DataFilterResult
        .Label2_Concordance = CountConcordance & " (" & CountConcordancePercent & "%)"
        .Label2_Error75 = CountError75 & " (" & CountError75Percent & "%)"
        .Label2_F206 = Countf206 & " (" & Countf206Percent & "%)"
        .Label2_Rho = CountRho & " (" & CountRhoPercent & "%)"
        .Label2_Total = AnalysesTotal & " (" & FailedAnalyses & " failed)"
        
        .Show
    End With
            
    

End Sub

Private Sub Cancel_Click()

    Call UnloadAll
        End
        
End Sub
