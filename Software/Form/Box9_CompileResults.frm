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

Private Sub CommandButton4_Click()

    FolderAddress = SelectFolderCompilation
    
End Sub

Private Sub Label14_Click()

End Sub

Private Sub UserForm_Initialize()

    Call ComboBoxSheetsNames
    
End Sub


