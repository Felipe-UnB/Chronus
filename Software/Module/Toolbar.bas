Attribute VB_Name = "Toolbar"
Option Explicit

    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
    Const SE_ERR_FNF = 2&
    Const SE_ERR_PNF = 3&
    Const SE_ERR_ACCESSDENIED = 5&
    Const SE_ERR_OOM = 8&
    Const SE_ERR_DLLNOTFOUND = 32&
    Const SE_ERR_SHARE = 26&
    Const SE_ERR_ASSOCINCOMPLETE = 27&
    Const SE_ERR_DDETIMEOUT = 28&
    Const SE_ERR_DDEFAIL = 29&
    Const SE_ERR_DDEBUSY = 30&
    Const SE_ERR_NOASSOC = 31&
    Const ERROR_BAD_FORMAT = 11&
    
    Public oToolbar As CommandBar
    Public StartButton As CommandBarButton
    Public Bt As CommandBarButton 'generic CommandBarButton
    Public OpenFiles As CommandBarButton
    Public SlpStdBlkCorr_Calc As CommandBarButton
    Public SlpStdCorr_Calc As CommandBarButton
    Public ConvertToPercent As CommandBarButton
    Public ConvertToAbsolute As CommandBarButton
    Public StartOptions As CommandBarButton
    Public FormatSheets As CommandBarButton
    Public OpenAnalysisByID As CommandBarButton
    Public CloseAnalysisByID As CommandBarButton
    Public StdDevTest As CommandBarButton
    Public FilterData As CommandBarButton
    Public NextID As CommandBarButton
    Public PreviousID As CommandBarButton
    Public RestoreData As CommandBarButton
    Public CreateaFinalReport As CommandBarButton
    Public ChartTitleAsSampleName As CommandBarButton
    Public QuestionHelp As CommandBarButton
    
    Const MyToolbar As String = "UPb Data Reduction" ' Give the toolbar a name

Sub AddToolbar()
    'Code mostly based on http://www.pptfaq.com/FAQ00031_Create_an_ADD-IN_with_TOOLBARS_that_run_macros.htm

    On Error Resume Next
    ' so that it doesn't stop on the next line if the toolbar's already there

    ' Create the toolbar; PowerPoint will error if it already exists
    Set oToolbar = CommandBars.Add(Name:=MyToolbar, _
        Position:=msoBarFloating, Temporary:=True)

    If Err.Number <> 0 Then

        Exit Sub 'The toolbar's already there, so we have nothing to do

    End If

    On Error GoTo ErrorHandler

    ' Now add a button to the new toolbar
    Set StartOptions = oToolbar.Controls.Add(Type:=msoControlButton)
    Set StartButton = oToolbar.Controls.Add(Type:=msoControlButton)
    Set SlpStdBlkCorr_Calc = oToolbar.Controls.Add(Type:=msoControlButton)
    Set SlpStdCorr_Calc = oToolbar.Controls.Add(Type:=msoControlButton)
    Set ConvertToPercent = oToolbar.Controls.Add(Type:=msoControlButton)
    Set ConvertToAbsolute = oToolbar.Controls.Add(Type:=msoControlButton)
    Set FormatSheets = oToolbar.Controls.Add(Type:=msoControlButton)
    Set OpenFiles = oToolbar.Controls.Add(Type:=msoControlButton)
    Set OpenAnalysisByID = oToolbar.Controls.Add(Type:=msoControlButton)
    Set CloseAnalysisByID = oToolbar.Controls.Add(Type:=msoControlButton)
    Set RestoreData = oToolbar.Controls.Add(Type:=msoControlButton)
    Set NextID = oToolbar.Controls.Add(Type:=msoControlButton)
    Set PreviousID = oToolbar.Controls.Add(Type:=msoControlButton)
    Set StdDevTest = oToolbar.Controls.Add(Type:=msoControlButton)
    Set FilterData = oToolbar.Controls.Add(Type:=msoControlButton)
    Set CreateaFinalReport = oToolbar.Controls.Add(Type:=msoControlButton)
    Set ChartTitleAsSampleName = oToolbar.Controls.Add(Type:=msoControlButton)
    Set QuestionHelp = oToolbar.Controls.Add(Type:=msoControlButton)

    ' And set some of the button's properties
        With StartButton

            'Tooltip text when mouse if placed over button
            .DescriptionText = "Start a complete process of data reduction."

            'Text if Text in Icon is chosen
            .Caption = "Complete data reduction"

            'Runs the Sub Button1() code when clicked
            .OnAction = "Button1_FullDataReduction" 'Procedure that will be executed when this button is clicked

            ' Button displays as icon, not text or both
            .Style = msoButtonIcon

            ' chooses icon #52 from the available Office icons
            .FaceId = 610

        End With

        With OpenFiles

            .DescriptionText = "Select the ID of the analysis to open it."
            .Caption = "Open analysis"
            .OnAction = "Button2_OpenFilesByIDs" 'Procedure that will be executed when this button is clicked
            .Style = msoButtonIcon
            .FaceId = 733

        End With

        With SlpStdBlkCorr_Calc

            .DescriptionText = "Calculates samples and standards (internal and external) blank corrected ratios and errors."
            .Caption = "Correct data for blank"
            .OnAction = "Button3_CalcAllSlpStd_BlkCorr" 'Procedure that will be executed when this button is clicked
            .Style = msoButtonIcon
            .FaceId = 1771

        End With

        With SlpStdCorr_Calc

            .DescriptionText = "Correct all samples and internal standards by external standard."
            .Caption = "Correct all sample by standard"
            .OnAction = "Button4_CalcAllSlp_StdCorr"
            .Style = msoButtonIcon
            .FaceId = 2112

        End With

        With ConvertToPercent

            .DescriptionText = "Convert all uncertanties in BlkCalc_Sh, SlpStdBlkCorr_Sh e SlpStdCorr_Sh to percentage."
            .Caption = "Relative uncertanties"
            .OnAction = "Button_ConvertToPercent"
            .Style = msoButtonIcon
            .FaceId = 6238

        End With

        With ConvertToAbsolute

            .DescriptionText = "Convert all uncertanties in BlkCalc_Sh, SlpStdBlkCorr_Sh e SlpStdCorr_Sh to absolute."
            .Caption = "Absolute uncertanties"
            .OnAction = "Button_ConvertToAbsolute"
            .Style = msoButtonIcon
            .FaceId = 6237

        End With

        With StartOptions 'UPDATE

            .DescriptionText = "Start a new reduction or change of the options."
            .Caption = "Option userforms"
            .OnAction = "Button_StartOptions"
            .Style = msoButtonIcon
            .FaceId = 2102

        End With

        With FormatSheets 'UPDATE

            .DescriptionText = "Apply the default format to all worksheets."
            .Caption = "Format worksheets"
            .OnAction = "Button_FormatSheets"
            .Style = msoButtonIcon
            .FaceId = 3249

        End With

        With OpenAnalysisByID

            .DescriptionText = "Open and plot an analysis based on number selected by the user."
            .Caption = "Open analyses by IDs"
            .OnAction = "Button_OpenAnalysisByID"
            .Style = msoButtonIcon
            .FaceId = 1561

        End With

        With CloseAnalysisByID

            .DescriptionText = "Close an analysis plot saving the cycles selected by the user."
            .Caption = "Close analyses plots"
            .OnAction = "Button_CloseAnalysisByID"
            .Style = msoButtonIcon
            .FaceId = 1087

        End With

        With StdDevTest

            .DescriptionText = "Let the user select isotopes and ratios to run a standard deviation test."
            .Caption = "Standard deviation test"
            .OnAction = "Button_StdDevTest"
            .Style = msoButtonIcon
            .FaceId = 2146

        End With

        With FilterData

            .DescriptionText = "Based on some criteria set by the user, " & _
                                "this program strikethrough cells that fails some tests."
            .Caption = "Filter data"
            .OnAction = "Button_FilterData"
            .Style = msoButtonIcon
            .FaceId = 601

        End With

        With NextID

            .DescriptionText = "Opens and plots the next analysis based on its ID."
            .Caption = "Next ID"
            .OnAction = "Button_NextID"
            .Style = msoButtonIcon
            .FaceId = 39

        End With

        With PreviousID

            .DescriptionText = "Opens and plots the previous analysis based on its ID."
            .Caption = "Previous ID"
            .OnAction = "Button_PreviousID"
            .Style = msoButtonIcon
            .FaceId = 41

        End With

        With RestoreData

            .DescriptionText = "Restores the plot to the state it has in the last time it was opened."
            .Caption = "Restore original plot"
            .OnAction = "Button_RestoreData"
            .Style = msoButtonIcon
            .FaceId = 37

        End With

        With CreateaFinalReport

            .DescriptionText = "Creates a formatted report ready to be published."
            .Caption = "Final Report"
            .OnAction = "Button_FinalReport"
            .Style = msoButtonIcon
            .FaceId = 161

        End With
        
        With ChartTitleAsSampleName
            
            .DescriptionText = "Change the title of the selected chart to the name of the sample."
            .Caption = "Chart title"
            .OnAction = "Button_ChartTitleAsSampleName"
            .Style = msoButtonIcon
            .FaceId = 1058
                    
        End With
        
        With QuestionHelp
            
            .DescriptionText = "Opens the support website."
            .Caption = "Chronus support"
            .OnAction = "Button_QuestionHelp"
            .Style = msoButtonIcon
            .FaceId = 926
                    
        End With

    ' Repeat the above for as many more buttons as you need to add
    ' Be sure to change the .OnAction property at least for each new button

    ' You can set the toolbar position and visibility here if you like
    ' By default, it'll be visible when created. Position will be ignored in PPT 2007 and later
    oToolbar.Top = 150
    oToolbar.Left = 150
    oToolbar.Visible = True

NormalExit:
    Exit Sub   ' so it doesn't go on to run the errorhandler code

ErrorHandler:
    'Just in case there is an error
    MsgBox Err.Number & vbCrLf & Err.Description
    Resume NormalExit:
End Sub

Sub Button1_FullDataReduction()
 'This code will run when you click Button 1 added above
 'Add a similar subroutine for each additional button you create on the toolbar
     'This is just some silly example code.
     'You 'd put your real working code here to do whatever
     'it is that you want to do

    Call FullDataReduction

    Call UnloadAll: End

End Sub

Sub Button2_OpenFilesByIDs()

    Call OpenFilesByIDs

    Call UnloadAll: End

End Sub

Sub Button3_CalcAllSlpStd_BlkCorr()

    Application.ScreenUpdating = False

    Call ConvertAbsolute
        Call CalcAllSlpStd_BlkCorr
            Call FormatMainSh

    Call UnloadAll: End

    Application.ScreenUpdating = True

End Sub

Sub Button4_CalcAllSlp_StdCorr()

    Application.ScreenUpdating = False

    Call ConvertAbsolute
        Call CalcAllSlp_StdCorr
            Call FormatMainSh

    Call UnloadAll: End

    Application.ScreenUpdating = True

End Sub

Sub Button_ConvertToPercent()

    Call ConvertPercentage

    Call UnloadAll: End

End Sub

Sub Button_FormatSheets()

    Call FormatMainSh

    Call UnloadAll: End

End Sub

Sub Button_ConvertToAbsolute()

    Call ConvertAbsolute

    Call UnloadAll: End

End Sub

Sub Button_StartOptions()

    Application.ScreenUpdating = False

    Box1_Start.Show

    Application.ScreenUpdating = True

    Call UnloadAll: End

End Sub

Sub Button_OpenAnalysisByID()

    'Based on user selection, this program will check is the cell contents is a number and then
    'open the analysis based on this number, which should be its ID.

    Dim IDnumber As String
    Dim CellInRange

    Application.ScreenUpdating = False

    'The code below is an modification to let the user select many IDs to open at once
'    For Each CellInRange In Selection
'        If IsNumeric(CellInRange) = False Then
'            MsgBox "All cells must be an integer!"

    If IsNumeric(ActiveCell) = False Then
1        IDnumber = InputBox("What is the analysis ID?", "Analysis ID")
            If IsNumeric(IDnumber) = False Then
                If MsgBox("You must choose an integer. Would you like to try again?", vbYesNo) = vbYes Then
                    GoTo 1
                Else
                    End
                End If
            End If
    Else
        IDnumber = ActiveCell
    End If

    FailToOpen = False 'It is necessary to set this because if in one loop thos variables be changed to true, on the next loop
    'it has to come back to its initial state (false)

    Call OpenAnalysisToPlot_ByIDs(Val(IDnumber), False)

        If FailToOpen = False Then
            Call Plot_PlotAnalysis(Plot_Sh, True, False, True, False, True, True, True)
                Call LineUpMyCharts(Plot_Sh, 1)

        Else
            Application.DisplayAlerts = False
                On Error Resume Next
                    Plot_Sh.Delete
                    Plot_ShHidden.Delete
                On Error GoTo 0
            Application.DisplayAlerts = True

        End If
    Call UnloadAll: End

    Application.ScreenUpdating = True

End Sub

Sub Button_NextID()

    Dim ActualID As Integer
    Dim NextID As Integer

    Application.ScreenUpdating = False

    Call PublicVariables
    
    Set Plot_Sh = ActiveSheet 'This line is necessary because Plot_sh is not set before this procedure

    Call CheckPlotSheet(Plot_Sh)

    'Plot_Sh.Range(Plot_IDCell) is the range in Plot_Sh with the ID of the plotted analysis.
    If Not IsEmpty(Plot_Sh.Range(Plot_IDCell)) = True And IsNumeric(Plot_Sh.Range(Plot_IDCell)) = True Then

        ActualID = Plot_Sh.Range(Plot_IDCell)
        NextID = Plot_Sh.Range(Plot_IDCell) + 1

        Call Plot_ClosePlot(ActiveSheet, False)
            Call OpenAnalysisToPlot_ByIDs(Val(NextID), False)

                If FailToOpen = False Then
                    Call Plot_PlotAnalysis(Plot_Sh, True, False, True, False, True, True, True)
                        Call LineUpMyCharts(Plot_Sh, 1)

                Else
                    Application.DisplayAlerts = False
                        Plot_Sh.Delete
                        Plot_ShHidden.Delete
                    Application.DisplayAlerts = True

                End If

    Else

        MsgBox "Please, check the cell " & Plot_Sh.Range(Plot_IDCell).Address & ", it should contain the ID of the analysis plotted.", vbOKOnly

    End If

    Call UnloadAll: End

    Application.ScreenUpdating = True

End Sub

Sub Button_PreviousID()

    Dim ActualID As Integer
    Dim PreviousID As Integer

    Application.ScreenUpdating = False

    Call PublicVariables
    
    Set Plot_Sh = ActiveSheet 'This line is necessary because Plot_sh is not set before this procedure
    
    Call CheckPlotSheet(Plot_Sh)

    'Plot_Sh.Range(Plot_IDCell) is the range in Plot_Sh with the ID of the plotted analysis.
    If Not IsEmpty(Plot_Sh.Range(Plot_IDCell)) = True And IsNumeric(Plot_Sh.Range(Plot_IDCell)) = True Then

        ActualID = Plot_Sh.Range(Plot_IDCell)
        PreviousID = Plot_Sh.Range(Plot_IDCell) - 1

        Call Plot_ClosePlot(ActiveSheet, False)
            Call OpenAnalysisToPlot_ByIDs(Val(PreviousID), False)

                If FailToOpen = False Then
                    Call Plot_PlotAnalysis(Plot_Sh, True, False, True, False, True, True, True)
                        Call LineUpMyCharts(Plot_Sh, 1)

                Else
                    Application.DisplayAlerts = False
                        On Error Resume Next
                            Plot_Sh.Delete
                            Plot_ShHidden.Delete
                        On Error GoTo 0
                    Application.DisplayAlerts = True

                End If

    Else

        MsgBox "Please, check the cell " & Plot_Sh.Range(Plot_IDCell).Address & ", it should contain the ID of the analysis plotted.", vbOKOnly

    End If

    Call UnloadAll: End

    Application.ScreenUpdating = True

End Sub


Sub Button_CloseAnalysisByID()

    Call Plot_ClosePlot(ActiveSheet)

    Call UnloadAll: End

End Sub

Sub Button_StdDevTest()

    Box6_StdDevTest.Show

    Call UnloadAll: End

End Sub

Sub Button_FilterData()

    Application.ScreenUpdating = False
    
        Box5_DataFilter.Show
    
        Call UnloadAll
        
    Application.ScreenUpdating = True

End Sub

Sub Button_RestoreData()

    Call RestoreOriginalPlotData(ActiveSheet)

    Call UnloadAll: End

End Sub

Sub DeleteFromShortcut()

    On Error Resume Next
        For Each Bt In CommandBars(MyToolbar).Controls
            CommandBars(MyToolbar).Controls(Bt).Delete
        Next
        
        CommandBars(MyToolbar).Delete
        
        If Err.Number <> 0 Then
            MsgBox "It was not possible to remove Chronus toolbar."
        End If
        
    On Error GoTo 0
    
End Sub

Sub teste()

    On Error Resume Next
          
        For Each Bt In CommandBars("Isoplot 3 Worksheet Tools").Controls
            CommandBars("Isoplot 3 Worksheet Tools").Controls(Bt).Delete
        Next
        
        CommandBars("Isoplot 3 Worksheet Tools").Delete

    On Error GoTo 0
    
End Sub

Sub Button_FinalReport()

    Call CreateFinalReport

    Call UnloadAll: End

End Sub

Sub Button_ChartTitleAsSampleName()

    Call ChangeChartTitleToSampleName

    Call UnloadAll: End
    
End Sub

Sub Button_QuestionHelp()

    'Created 21122015
    'Based on Walkenbach (2010) OpenURL function, at page 681, and https://support.microsoft.com/en-us/kb/170918
    'It uses the ShellExecute API function declared at the beginning of this module (based on the microsoft support
    'page previously cited
    
    'This program will try to open the Chronus support website on GitHub.
    
    Dim WebAddress As String
    Dim URL As String
    Dim Result As Long
    Dim ErrMsg As String

    WebAddress = "https://github.com/Felipe-UnB/Chronus/issues/new"

    Result = ShellExecute(0&, vbNullString, WebAddress, _
    vbNullString, vbNullString, vbNormalFocus)
    
    If Result <= 32 Then
        Select Case Result
            Case SE_ERR_FNF
                ErrMsg = "File not found"
            Case SE_ERR_PNF
                ErrMsg = "Path not found"
            Case SE_ERR_ACCESSDENIED
                ErrMsg = "Access denied"
            Case SE_ERR_OOM
                ErrMsg = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                ErrMsg = "DLL not found"
            Case SE_ERR_SHARE
                ErrMsg = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                ErrMsg = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                ErrMsg = "DDE Time out"
            Case SE_ERR_DDEFAIL
                ErrMsg = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                ErrMsg = "DDE busy"
            Case SE_ERR_NOASSOC
                ErrMsg = "No association for file extension"
            Case ERROR_BAD_FORMAT
                ErrMsg = "Invalid EXE file or error in EXE image"
            Case Else
                ErrMsg = "Unknown error"
        End Select
        
        MsgBox "It was not possible to open the support website. (" & ErrMsg & ")"
        
    End If
    
End Sub

'Private Sub Workbook_Open()
'
'    Call Auto_Open
'
'End Sub
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Call DeleteFromShortcut

End Sub


