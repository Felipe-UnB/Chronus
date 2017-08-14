VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box10_Presentation 
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7200
   OleObjectBlob   =   "Box10_Presentation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Box10_Presentation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()

    Button_QuestionHelp ("https://github.com/Felipe-UnB/Chronus/wiki")

End Sub


Private Sub userForm_Initialize()
    
    SetPresentationVariable
    
    Application.EnableEvents = False
    Box10_Presentation.CheckBox1.Value = Not (ShowPresentation)

    Label4 = ChronusVersion

End Sub

Private Sub CheckBox1_Click()
    
    SetPresentationVariable
    
    If Box10_Presentation.CheckBox1.Value = True Then
        ShowPresentation = False
    Else
        ShowPresentation = True
    End If
    
    TW.Save

End Sub

Sub SetPresentationVariable()

    On Error Resume Next
        If ShowPresentation Is Nothing Then
            Set TW = ThisWorkbook
                Set StartANDOptions_TW_Sh = TW.Worksheets("Start-AND-Options")
                    Set ShowPresentation = StartANDOptions_TW_Sh.Range("B58")
        End If
        
        If Err.Number <> 0 Then
            Exit Sub
        End If
    On Error GoTo 0

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
        
    If CloseMode = vbFormControlMenu Then
                
        Unload Box10_Presentation
                
    End If
    
End Sub

