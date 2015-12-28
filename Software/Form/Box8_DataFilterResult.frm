VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box8_DataFilterResult 
   Caption         =   "Data filter results"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3540
   OleObjectBlob   =   "Box8_DataFilterResult.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Box8_DataFilterResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Ok_Click()
    
    Box8_DataFilterResult.Hide
        Unload Box8_DataFilterResult
    
End Sub

