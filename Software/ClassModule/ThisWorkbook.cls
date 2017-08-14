VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()

    'This program is executed as soon as the Chronus workbook is opened, which happens every time excel is started.

'    Dim i As Long
    Dim IsoplotFound As Boolean
'    Dim ChronusProject As Object
                
    Set TW = ThisWorkbook
        Set StartANDOptions_TW_Sh = TW.Worksheets("Start-AND-Options")
            Set ShowPresentation = StartANDOptions_TW_Sh.Range("B58")
            
    Call CheckAccessVBPM
    
'    Call CheckIsoplotReference
        
'    If Application.Version <> "14.0" Then 'Chronus was developed in this version
'
'        If MsgBox("This version of excel was not tested with Chronus. " & _
'                "Would you like to try installing it anyway?", vbYesNo) = vbNo Then
'            End
'        Else
'            Set ChronusProject = Application.AddIns2(ProgramName) 'An object is associated to the Chronus VBA project
'                If ChronusProject.Installed = False Then 'If Chronus is not loaded (for excel, to load is the same as to install)
'                    ChronusProject.Installed = True
'                End If
'        End If
'
'    Else
'
'        Set ChronusProject = Application.AddIns2(ProgramName) 'An object is associated to the Chronus VBA project
'            If ChronusProject.Installed = False Then 'If Chronus is not loaded (for excel, to load is the same as to install)
'                Application.AddIns2(ProgramName).Installed = True
'            End If
'    End If

    On Error Resume Next
        If ShowPresentation = True Then
            Box10_Presentation.Show
        End If
    On Error GoTo 0
    
    Call AddToolbar

End Sub



