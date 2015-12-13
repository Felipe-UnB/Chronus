VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box6_StdDevTest 
   Caption         =   "Standard Deviation Test"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3570
   OleObjectBlob   =   "Box6_StdDevTest.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Box6_StdDevTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Ok_Click()

'    If Not _
'        CheckBox1_Test68.Value = False And _
'        CheckBox2_Test76.Value = False And _
'        CheckBox3_Test28.Value = False And _
'        CheckBox4_Test74.Value = False And _
'        CheckBox5_Test64.Value = False Then
        
        Box6_StdDevTest.Hide
        
        Application.DisplayAlerts = False
        
        Call StandardDeviationTest(Plot_Sh, CheckBox1_Test68.Value, CheckBox6_Line68.Value, CheckBox2_Test76.Value, TextBox1, _
            CheckBox3_Test28.Value, CheckBox4_Test74.Value, CheckBox5_Test64.Value)
        
        Application.DisplayAlerts = True
        
'    End If
    
    Call UnloadAll

End Sub

Private Sub CommandButton2_TestAll_Click()

    Dim counter As Long
    Dim TotalTime As Double
    Dim NumAnalyses As Long
    
    TotalTime = Timer
    
    If MsgBox("This process might take a long time if more than 100 analyses should be processed. " & _
    "Would you like to continue?", vbYesNo) = vbNo Then
        Call UnloadAll
            Exit Sub
    End If
        
    ScreenUpd = Application.ScreenUpdating
    
    Application.ScreenUpdating = False
    
    Box6_StdDevTest.Hide
    
    Call SetPathsNamesIDsTimesCycles
    
    NumAnalyses = NumElements(PathsNamesIDsTimesCycles, 2) - 1
    
    For counter = 1 To NumAnalyses
        
        FailToOpen = False 'It is necessary to set this because if in one loop thos variables be changed to true, on the next loop
        'it has to come back to its initial state (false)
            
            Call OpenAnalysisToPlot_ByIDs(Val(counter), True)
        
        If FailToOpen = False Then
        
'            Call Plot_PlotAnalysis(Plot_Sh)
'                Call LineUpMyCharts(Plot_Sh, 1)
        
            Call StandardDeviationTest(Plot_Sh, CheckBox1_Test68.Value, CheckBox6_Line68.Value, CheckBox2_Test76.Value, TextBox1, _
                CheckBox3_Test28.Value, CheckBox4_Test74.Value, CheckBox5_Test64.Value, False, True)
                
            Call Plot_ClosePlot(Plot_Sh, False)
                
        Else
            Application.DisplayAlerts = False
                On Error Resume Next
                    Plot_Sh.Delete
                    Plot_ShHidden.Delete
                On Error GoTo 0
            Application.DisplayAlerts = True

        End If
        
    Next
    
    TotalTime = Timer - TotalTime
    
    MsgBox "Standard deviation test: " & NumAnalyses & " analysis(es) in " & Round(TotalTime, 4) & _
                                    " s (" & Round(NumAnalyses / TotalTime, 3) & " s per analysis)"
    
    If MsgBox("Would you like to start the complete data reduction process, in order to the " & _
    "standard deviation test be effective? ", vbYesNo) = vbYes Then
            Call FullDataReduction
    End If

    Call UnloadAll

    Application.ScreenUpdating = ScreenUpd
End Sub

Private Sub TextBox1_Change()
    
    If Val(TextBox1.Value) = False Then
        MsgBox "You must enter a number bigger than 0."
            TextBox1.Value = ""
                TextBox1.SetFocus
    End If
    
End Sub

Private Sub UserForm_Initialize()
    
    Dim Sh As Worksheet
    
    If mwbk Is Nothing Then
        Call PublicVariables
    End If
    
    Set Plot_Sh = ActiveSheet
    
        On Error Resume Next
            Set Plot_ShHidden = mwbk.Worksheets(Plot_Sh.Name & "Hidden")
            
            If Err.Number = 9 Then
                MsgBox "This is not a valid plot worksheet."
                    End
            ElseIf Err.Number <> 0 Then
                MsgBox "An error occured!"
                    End
            End If
        On Error GoTo 0

End Sub


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
