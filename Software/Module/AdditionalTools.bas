Attribute VB_Name = "AdditionalTools"
Option Explicit

Public FolderAddress As String
Public Const Comp_SampleNameColumn As String = "A"
Public Const Comp_AnalysisDateColumn As String = "B"
Public Const Comp_AnalysisID As String = "C"
Public Const Comp_HeaderRow As Long = 1
Public CounterRow As Long

Sub MainProgram()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
        
        Call SelectFolderCompilation
        Call CreateWorkbookForStandards
        Call StandardCompilation
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub CreateWorkbookForStandards()
    
    Dim NewWorkbook As Workbook
    
    Set NewWorkbook = Application.Workbooks.Add
    Set SlpStdCorr_Sh = NewWorkbook.Worksheets(1)
    
    With SlpStdCorr_Sh
        
        .Name = SlpStdCorr_Sh_Name
        
        Call FormatSlpStdCorr(True, False)
        
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

    Dim FSO As Scripting.FileSystemObject
    Dim WorkbooksFolder As Object 'Scripting.Folder
    Dim File As Object 'Scripting.File
    Dim a As Long 'Number of files with the specified extension found
    Dim DesiredExtension As String 'String to be removed from the name of the sample
    Dim OpenedWorkbook As Workbook
    Dim StandardName As String
    Dim InputBoxStandardName As String
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

    InputBoxStandardName = InputBox("What is the name of the standard?", "Standard Name")
        
        If InputBoxStandardName = False Or Len(InputBoxStandardName) = 0 Then
            MsgBox "The program will stop."
                End
'        Else
'            StandardName = InputBoxStandardName
        End If

    CounterRow = StdCorr_HeaderRow + 1

    For Each File In WorkbooksFolder.Files
        
        If FSO.getExtensionName(File.path) = DesiredExtension Then
            
            Set OpenedWorkbook = Workbooks.Open(File.path)
                Set CellToPaste = SlpStdCorr_Sh.Range(Comp_AnalysisID & CounterRow)
            
            Call CopyStandard(InputBoxStandardName, OpenedWorkbook, CellToPaste)
                
                'CounterRow = CounterRow + 1
            
            OpenedWorkbook.Close (False)
            
        End If

    Next

End Sub

Sub SelectFolderCompilation()

    'A slightly different version of the original SelectFolder procedure to let the user just select the folder where the workbooks with
    'standard results are

    'Created 13112015 - By Felipe Valença
    
    Dim strButtonCaption As String
    Dim strDialogTitle As String
    Dim SelectDialog As FileDialog
    Dim SelectionDone As Integer
    'Dim StandardFolderPath As String  - I pretend to use this variable to check if the usar has chosen some folder and not just hit the "Select a Folder" button
        
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
  
End Sub

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
'                        On Error Resume Next
'                            With WB.Worksheets(SamList_Sh_Name)
'
'                                Set FindDate = .Range(SamList_FileName & 1).EntireColumn.Find(FindStandard.Value)
'
'                                    If Err.Number = 0 Then
'                                        Set AnalysisDateRange = .Range(SamList_FirstCycleTime & FindDate.Row)
'                                            AnalysisDateRange.Copy
'                                                SlpStdCorr_Sh.Range(Comp_AnalysisDateColumn & CellToPaste.Row).PasteSpecial (xlPasteValuesAndNumberFormats)
'                                    End If
'                            End With
'                        On Error GoTo 0
            
                        With Ws.Range(StdCorr_SlpName & 1, Ws.Range(StdCorr_SlpName & 1).End(xlDown))
                            Set FindStandard = .FindNext(FindStandard)
                        End With

                    Loop While Not FindStandard Is Nothing And FindStandard.Address <> FirstAddress
                        
                End If
        End If
    Next
    
    CounterRow = CounterRow + Counter
    
End Sub

