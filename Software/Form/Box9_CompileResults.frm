VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box9_CompileResults 
   Caption         =   "Compile results"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6735
   OleObjectBlob   =   "Box9_CompileResults.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Box9_CompileResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton3_Ok_Click()

        
    If Len(TextBox1_AnalysesNames) = 0 Then
        MsgBox "No name provided."
            TextBox1_AnalysesNames.SetFocus
                Exit Sub
    End If
    
    If ComboBox1_Sheets = "" Then
        MsgBox "The sheet to search in was not selected.", vbOKOnly
            ComboBox1_Sheets.SetFocus
                Exit Sub
    End If
    
    Call CompileAnalyses

End Sub

Private Sub CommandButton4_Click()

    FolderAddress = SelectFolderCompilation
    
End Sub

Private Sub Label14_Click()

End Sub

Private Sub TextBox8_BlankName_Change()

End Sub

Private Sub TextBox1_AnalysesNames_Change()

End Sub

Private Sub UserForm_Initialize()

    Call ComboBoxSheetsNames
    
End Sub


