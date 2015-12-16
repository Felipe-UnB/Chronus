VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box3_ProblemsList 
   Caption         =   "Missing Information"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   OleObjectBlob   =   "Box3_ProblemsList.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Box3_ProblemsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Ok_Click()
    
    Box3_ProblemsList.Hide

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim Response As Integer
    
    If CloseMode = vbFormControlMenu Then
    
        Box3_ProblemsList.Hide
        
'        Response = MsgBox("Do you really want to end the program execution?", vbYesNo)
'            If Response = vbNo Then
'                Cancel = True
'            ElseIf Response = vbYes Then
'                Call UnloadAll
'                    End
'            End If
    End If
    
End Sub
