Attribute VB_Name = "AdditionalTools"
Option Explicit

Public FolderAddress As String
Public Const Comp_SampleNameColumn As String = "A"
Public Const Comp_AnalysisDateColumn As String = "B"
Public Const Comp_AnalysisID As String = "C"
Public Const Comp_HeaderRow As Long = 1
Public CounterRow As Long
Dim NewSheet As Worksheet

Sub CompileAnalyses()

    'This procedure will allow the user to open multiple results (.xlsx) and copy the
    'selected analyses to a single file.
    
    'Only results at SlpStdCorr sheets can be copied.

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
        Call CreateWorkbookForAnalyses
        Call StandardCompilation
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub CreateWorkbookForAnalyses()

    'Creates a new workbook, with the SplStdCorr sheet to store the compiled results.
    'Updated 19022015 - Different sheets are created depending on the user choice for
    'the type of sheet to be examined.
    
    Dim NewWorkbook As Workbook
    
    Set NewWorkbook = Application.Workbooks.Add
    Set NewSheet = NewWorkbook.Worksheets(1)
    
    With NewSheet
        
        Select Case ComboBox1_Sheets
            
            Case BlkCalc_Sh_Name
                .Name = BlkCalc_Sh_Name
                    Call FormatBlkCalc(True)
            
            Case SlpStdBlkCorr_Sh_Name
                .Name = SlpStdBlkCorr_Sh_Name
                    Call FormatSlpStdBlkCorr(True)
            
            Case SlpStdCorr_Sh_Name
                .Name = SlpStdCorr_Sh_Name
                    Call FormatSlpStdCorr(True, False)
                    
        End Select
                
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

Sub StandardCompilation()

    'This procedure checks all files in the selected folder. For those with the
    'DesiredExtension, they are opened and then the CopyStandard procedure is
    'called.

    Dim FSO As Scripting.FileSystemObject
    Dim WorkbooksFolder As Object 'Scripting.Folder
    Dim File As Object 'Scripting.File
    Dim a As Long 'Number of files with the specified extension found
    Dim DesiredExtension As String 'String to be removed from the name of the sample
    Dim OpenedWorkbook As Workbook
    Dim StandardName As String
    Dim CellToPaste As Range
 
    DesiredExtension = "xlsx"
       
    Set FSO = CreateObject("Scripting.FileSystemObject") 'If the variable FSO is already declared as Scripting.Filesystem why do I have to set it like this?
       
    'On Error Resume Next
        Set WorkbooksFolder = FSO.GetFolder(FolderAddress)

        If Err.Number <> 0 Then
            MsgBox "Invalid folder."
                Exit Sub
        End If
    On Error GoTo 0

    CounterRow = StdCorr_HeaderRow + 1

    For Each File In WorkbooksFolder.Files
        
        If FSO.getExtensionName(File.path) = DesiredExtension Then
            
            Set OpenedWorkbook = Workbooks.Open(File.path)
                Set CellToPaste = NewSheet.Range(Comp_AnalysisID & CounterRow)
            
            Call CopyStandard(InputBoxStandardName, OpenedWorkbook, CellToPaste)
            
            OpenedWorkbook.Close (False)
            
        End If

    Next

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

Sub CopyStandard(StandardName As String, WB As Workbook, CellToPaste As Range)

    Dim Ws As Worksheet
    Dim FindStandard As Object
    Dim FindDate As Object
    Dim RangeToCopy As Range
    Dim FirstAddress As String
    Dim AnalysisDateRange As Range
    Dim Cell1 As Range
    Dim Counter As Long
    
    For Each Ws In WB.Worksheets
        If Ws.Name = SlpStdCorr_Sh_Name Then
            
            With Ws.Range(StdCorr_SlpName & 1, Ws.Range(StdCorr_SlpName & 1).End(xlDown))
                Set FindStandard = .Find(StandardName)
            End With
            
                If Not FindStandard Is Nothing Then
                
                    FirstAddress = FindStandard.Address
                    
                    Counter = 0
                    SlpStdCorr_Sh.Activate
                    Do
                    
                        Set Cell1 = Ws.Cells(FindStandard.Row, FindStandard.Column)
                    
                        Set RangeToCopy = Ws.Range(StdCorr_ColumnID & Cell1.Row, Ws.Range(StdCorr_ColumnID & Cell1.Row).End(xlToRight))
                            RangeToCopy.Copy
                                CellToPaste.Offset(Counter).PasteSpecial (xlPasteValues)
                                SlpStdCorr_Sh.Range(Comp_SampleNameColumn & CellToPaste.Row + Counter) = WB.Name
                                
                        Counter = Counter + 1
            
                        With Ws.Range(StdCorr_SlpName & 1, Ws.Range(StdCorr_SlpName & 1).End(xlDown))
                            Set FindStandard = .FindNext(FindStandard)
                        End With

                    Loop While Not FindStandard Is Nothing And FindStandard.Address <> FirstAddress
                        
                End If
        End If
    Next
    
    CounterRow = CounterRow + Counter
    
End Sub

Sub ComboBoxSheetsNames()

    'This program will populate comboxes from Box1_Start
    'and Box2_UPb_Options will standards informations
    'stored in add-in workbook.

    Dim StandardsNamesHeader As Range 'Cell with standard names header.
    Dim Counter As Integer 'Used to add itens to External Standard ComboBox
    Dim SheetsNames(1 To 3) As String
    
    'The lines below will add values to the array SheetsNames
    SheetsNames(1) = BlkCalc_Sh_Name
    SheetsNames(2) = SlpStdBlkCorr_Sh_Name
    SheetsNames(3) = SlpStdCorr_Sh_Name

    For Counter = 1 To UBound(SheetsNames)
        Box9_CompileResults.ComboBox1_Sheets.AddItem (SheetsNames(Counter))
    Next

End Sub

