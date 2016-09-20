Attribute VB_Name = "AdditionalTools"
Option Explicit

Public FolderAddress As String
Public Const Comp_SampleNameColumn As String = "A"
Public Const Comp_AnalysisDateColumn As String = "B"
Public Const Comp_AnalysisID As String = "C"
Public Comp_HeaderRow As Long
Public CounterRow As Long

Public Comp_AnalysesName As String    'AnalysesName is the common name of all the analyses that should be copied. This is a user input from the
                                 'Box9_CompileResuults

Public Comp_TargetSheet As String   'Name of the sheet where the analyses should be searched for.


Public Comp_SlpName As String
Public Comp_ColumnID As String

Public Comp_NewSheet As Worksheet

Sub TestBox9()
    
    FolderAddress = "D:\UnB\Geocronologia\Chronus development\Data Compilation - Test Files\"
    Box9_CompileResults.ComboBox1_Sheets = FinalReport_Sh_Name
    Box9_CompileResults.TextBox1_AnalysesNames = "91500"

'    BlkCalc_Sh_Name
'    SlpStdBlkCorr_Sh_Name
'    SlpStdCorr_Sh_Name

End Sub

Sub CompileAnalyses()

    'This procedure will allow the user to open multiple results (.xlsx) and copy the
    'selected analyses to a single file.

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
        Call CreateWorkbookForAnalyses
        Call AnalysesCompilation
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub CreateWorkbookForAnalyses()

    'Creates a new workbook, with the SplStdCorr sheet to store the compiled results.
    'Updated 19022015 - Different sheets are created depending on the user choice for
    'the type of sheet to be examined.
    
    Dim NewWorkbook As Workbook
    
    Set TW = ThisWorkbook
    Set NewWorkbook = Application.Workbooks.Add
    Set Comp_NewSheet = NewWorkbook.Worksheets(1)
    
    With Comp_NewSheet
        
        Select Case Comp_TargetSheet

            Case BlkCalc_Sh_Name
                .Name = BlkCalc_Sh_Name
                
                    Comp_SlpName = BlkSlpName
                    Comp_ColumnID = BlkColumnID
                    Comp_HeaderRow = BlkCalc_HeaderLine
                    
                    Call FormatBlkCalc(True, True)

            Case SlpStdBlkCorr_Sh_Name
                .Name = SlpStdBlkCorr_Sh_Name
                
                    Comp_SlpName = ColumnSlpName
                    Comp_ColumnID = ColumnID
                    Comp_HeaderRow = HeaderRow
                    
                    Set SlpStdBlkCorr_Sh = Comp_NewSheet
                    
                    Call SetSlpStdBlkCorr_Sh_Variables
                    
                    Call FormatSlpStdBlkCorr(True, True)

            Case SlpStdCorr_Sh_Name
                .Name = SlpStdCorr_Sh_Name
                
                    Comp_SlpName = StdCorr_SlpName
                    Comp_ColumnID = StdCorr_ColumnID
                    Comp_HeaderRow = StdCorr_HeaderRow
 
                    Call FormatSlpStdCorr(True, False, True)
                    
            Case FinalReport_Sh_Name
            
                'NewWorkbook.Close (False)
                
                On Error Resume Next
                        
                    TW.Sheets(FinalReport_Sh_Name).Copy _
                        After:=NewWorkbook.Sheets(1)
                    
                    If Err.Number <> 0 Then
                        MsgBox "Chronus not instaled correctly"
                        Call UnloadAll
                        End
                    End If
                    
                    NewWorkbook.Worksheets(1).Delete
                    Set Comp_NewSheet = NewWorkbook.Worksheets(1)
                    Set FinalReport_Sh = Comp_NewSheet
                    
                    Comp_NewSheet.Rows("5:11").EntireRow.Delete
                    
                On Error GoTo 0
                                    
                    Comp_SlpName = FR_ColumnSlpName
                    Comp_ColumnID = FR_ColumnSlpName
                    Comp_HeaderRow = FR_HeaderRow
                    
                    Call FormatFinalReport(True)

        End Select
    End With
    
    With Comp_NewSheet
    
        'The lines below will add two columns (for samples' names and analyses' date)
        .Range("A1").Columns.EntireColumn.Insert (xlShiftToRight)
        .Range("A1").Columns.EntireColumn.Insert (xlShiftToRight)
        
        With .Range(Comp_SampleNameColumn & Comp_HeaderRow)
            .Value = "Sample"
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
        With .Range(Comp_AnalysisDateColumn & Comp_HeaderRow)
            .Value = "Analysis date"
            .Font.Bold = True
            .EntireColumn.AutoFit
        End With
        
    End With
    
End Sub

Sub AnalysesCompilation()

    'This procedure checks all files in the selected folder. For those with the
    'DesiredExtension, they are opened and then the CopyStandard procedure is
    'called.
    
    Dim FSO As Scripting.FileSystemObject
    Dim WorkbooksFolder As Object 'Scripting.Folder
    Dim File As Object 'Scripting.File
    Dim a As Long 'Number of files with the specified extension found
    Dim DesiredExtension1 As String 'Chronus' files extension
    Dim DesiredExtension2 As String 'Chronus' files extension
    Dim OpenedWorkbook As Workbook
    Dim StandardName As String
    Dim CellToPaste As Range
    Dim FoundExtension As Boolean 'Indicates if a single with one of the desired extension was found.
    
    DesiredExtension1 = "xlsx" 'Old Chronus files extension
    DesiredExtension2 = "xlsm" 'New Chronus files extension
    FoundExtension = False
    
    Set FSO = CreateObject("Scripting.FileSystemObject") 'If the variable FSO is already declared as Scripting.Filesystem why do I have to set it like this?
       
    On Error Resume Next
        Set WorkbooksFolder = FSO.GetFolder(FolderAddress)

        If Err.Number <> 0 Then
            MsgBox "Invalid folder."
                Exit Sub
        End If
    On Error GoTo 0
    
    Select Case Comp_TargetSheet

        Case BlkCalc_Sh_Name
            CounterRow = BlkCalc_HeaderLine + 1

        Case SlpStdBlkCorr_Sh_Name
            CounterRow = HeaderRow + 1
            
        Case SlpStdCorr_Sh_Name
            CounterRow = StdCorr_HeaderRow + 1
            
        Case FinalReport_Sh_Name
            CounterRow = FR_HeaderRow + 2

    End Select

    For Each File In WorkbooksFolder.Files
        
        If FSO.getExtensionName(File.path) = DesiredExtension1 Or FSO.getExtensionName(File.path) = DesiredExtension2 Then
            
            FoundExtension = True
            
            Set OpenedWorkbook = Workbooks.Open(File.path)
                Set CellToPaste = Comp_NewSheet.Range(Comp_AnalysisID & CounterRow)

            Call CopyAnalysis(OpenedWorkbook, CellToPaste)

            OpenedWorkbook.Close (False)

        End If

    Next

    If FoundExtension = False Then
        MsgBox "No files with the extensions " & DesiredExtension1 & " or " & DesiredExtension2 & " found!", vbOKOnly
    End If

    If CounterRow = StdCorr_HeaderRow + 1 Then
        MsgBox "No analysis with the name " & Comp_AnalysesName & " was found.", vbOKOnly
    End If
    
    If FoundExtension = True And CounterRow <> StdCorr_HeaderRow + 1 Then
        MsgBox "You will have to fill the data of analysis manually!", vbOKOnly
    End If
    
End Sub

Function SelectFolderCompilation() As String

    'A slightly different version of the original SelectFolder procedure to let the user just select the folder where the workbooks with
    'standard results are

    'Created 13112015 - By Felipe Valença
    
    Dim strButtonCaption As String
    Dim strDialogTitle As String
    Dim SelectDialog As FileDialog
    Dim SelectionDone As Integer
        
    Set SelectDialog = Application.FileDialog(msoFileDialogFolderPicker)
    
    'Captions of the SelectDialog
    strButtonCaption = "Select a Folder"
    strDialogTitle = "Folder Selection Dialog"
    

    With SelectDialog
        .ButtonName = strButtonCaption
        .InitialView = msoFileDialogViewDetails     'Detailed View
        .Title = strDialogTitle
        .AllowMultiSelect = False 'Let user just select only one folder
        'SelectDialog.Show displays a file dialog box and returns a Long indicating whether
        'the user pressed the Action button (-1) or the Cancel button (0).
        SelectionDone = .Show
        
        On Error Resume Next
            FolderAddress = .SelectedItems(1) & "\"

            If SelectionDone <> -1 Then 'The user has clicked on "Cancel" button
                End
            End If
        
        On Error GoTo 0
    End With
    
    SelectFolderCompilation = FolderAddress
  
End Function

Sub CopyAnalysis(WB As Workbook, CellToPaste As Range)

    Dim Ws As Worksheet
    Dim FindAnalysis As Object
    Dim FindDate As Object
    Dim RangeToCopy As Range
    Dim FirstAddress As String
    Dim AnalysisDateRange As Range
    Dim Cell1 As Range
    Dim Counter As Long
    
    On Error Resume Next
        Set Ws = WB.Worksheets(Comp_TargetSheet)
        If Err.Number <> 0 Then
            On Error GoTo 0
            Exit Sub
        End If
    On Error GoTo 0
    
    With Ws.Range(Comp_SlpName & Comp_HeaderRow, Ws.Range(Comp_SlpName & Comp_HeaderRow).End(xlDown))
        Set FindAnalysis = .Find(Comp_AnalysesName)
    End With
    
        If Not FindAnalysis Is Nothing Then
        
            FirstAddress = FindAnalysis.Address
            
            Counter = 0
            
            Do
            
                Set Cell1 = Ws.Cells(FindAnalysis.Row, FindAnalysis.Column)
            
                If Comp_NewSheet.Name = FinalReport_Sh_Name Then
                    Set RangeToCopy = Ws.Range(Comp_ColumnID & Cell1.Row, Ws.Range(FR_LastColumn & Cell1.Row))
                Else
                    Set RangeToCopy = Ws.Range(Comp_ColumnID & Cell1.Row, Ws.Range(Comp_ColumnID & Cell1.Row).End(xlToRight))
                End If
                    RangeToCopy.Copy
                        
                        CellToPaste.Offset(Counter).PasteSpecial (xlPasteValues)
                        Comp_NewSheet.Range(Comp_SampleNameColumn & CellToPaste.Row + Counter) = WB.Name
                                                
                Counter = Counter + 1
    
                With Ws.Range(Comp_SlpName & Comp_HeaderRow, Ws.Range(Comp_SlpName & Comp_HeaderRow).End(xlDown))
                    Set FindAnalysis = .FindNext(FindAnalysis)
                End With

            Loop While Not FindAnalysis Is Nothing And FindAnalysis.Address <> FirstAddress
                
        End If
        
        CounterRow = CounterRow + Counter
    
End Sub

Sub ComboBoxSheetsNames()

    'This program will populate Box8_CompileResults

    Dim StandardsNamesHeader As Range 'Cell with standard names header.
    Dim Counter As Integer 'Used to add itens to External Standard ComboBox
    Dim SheetsNames(1 To 4) As String
    
    'The lines below will add values to the array SheetsNames
    SheetsNames(1) = BlkCalc_Sh_Name
    SheetsNames(2) = SlpStdBlkCorr_Sh_Name
    SheetsNames(3) = SlpStdCorr_Sh_Name
    SheetsNames(4) = FinalReport_Sh_Name
    
    For Counter = 1 To UBound(SheetsNames)
        Box9_CompileResults.ComboBox1_Sheets.AddItem (SheetsNames(Counter))
    Next

End Sub

Sub CopyStandardCompilation()
    'This will copy the standard compilation to all opened spreadsheets.
    
    Dim CompilationWB As Workbook
    Dim CompilationWS As Worksheet
    Dim WB As Workbook
    Dim CompilationSheetName As String
    Dim InputStdDataTable As Variant
    Dim InputStdDiagram As Variant
    
    Dim SSDataTable As Worksheet
    Dim SSDiagram As Chart
    
    Set CompilationWB = ActiveWorkbook

    InputStdDataTable = InputBox("Please, write the name of the secondary standard data sheet.")
    If MsgBox("Is there another sheet you would like to copy?", vbYesNo) = vbYes Then
        InputStdDiagram = InputBox("Write the diagram's sheet name.")
    End If
       
    On Error Resume Next
        Set SSDataTable = CompilationWB.Worksheets(InputStdDataTable)
        Set SSDiagram = CompilationWB.Charts(InputStdDiagram)
    
        If Err.Number <> 0 Then

            MsgBox "At least one of the names is not correct.", vbOKOnly
            On Error GoTo 0
            Exit Sub

        End If
    On Error GoTo 0
         
    For Each WB In Workbooks
        
        If Not WB.Name = CompilationWB.Name And Not UCase(WB.Name) = "PERSONAL.XLSB" Then
            If MsgBox("Do you want to copy the sheet(s) to " & WB.Name & " ?", vbYesNo) = vbYes Then
                SSDataTable.Copy After:=WB.Worksheets(WB.Worksheets.count)
                SSDiagram.Copy After:=WB.Worksheets(WB.Worksheets.count)
            End If
        End If
        
    Next
    
End Sub
