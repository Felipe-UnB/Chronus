Attribute VB_Name = "ExportProject"
Option Explicit

Public Sub ExportVisualBasicCode()

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
'A little modification in the original code was made in order to export the different components to different folders.

'From https://gist.github.com/steve-jansen/7589478 with modifications

    Const Module = 1
    Const ClassModule = 2
    Const Form = 3
    Const Document = 100
    Const Padding = 24

    Dim VBProj As VBIDE.VBProject
    Dim VBComponent As Object
    Dim count As Integer
    Dim path As String
    Dim directory As String
    Dim DirectoryByType As String
    Dim extension As String
    Dim fso As New FileSystemObject
    
    Set VBProj = Application.VBE.ActiveVBProject
    
    directory = "D:\UnB\Chronus"
    count = 0

    If Not fso.folderexists(directory) Then
        Call fso.createfolder(directory)
        
            Call fso.createfolder(directory & "\" & "ClassModule")
                Call fso.createfolder(directory & "\" & "Form")
                    Call fso.createfolder(directory & "\" & "Module")
                        Call fso.createfolder(directory & "\" & "Other")
    End If
    
        If Not fso.folderexists(directory & "\" & "ClassModule") Then
            Call fso.createfolder(directory & "\" & "ClassModule")
        End If
    
            If Not fso.folderexists(directory & "\" & "Form") Then
                Call fso.createfolder(directory & "\" & "Form")
            End If
    
                If Not fso.folderexists(directory & "\" & "Module") Then
                    Call fso.createfolder(directory & "\" & "Module")
                End If
            
                    If Not fso.folderexists(directory & "\" & "Other") Then
                        Call fso.createfolder(directory & "\" & "Other")
                    End If
    
    Set fso = Nothing
    
    Application.SendKeys ("^g")
    Debug.Print
    
    For Each VBComponent In VBProj.VBComponents
        Select Case VBComponent.Type
            Case ClassModule, Document
                extension = ".cls"
                DirectoryByType = "ClassModule"
            Case Form
                extension = ".frm"
                DirectoryByType = "Form"
            Case Module
                extension = ".bas"
                DirectoryByType = "Module"
            Case Else
                extension = ".txt"
                DirectoryByType = "Other"
        End Select
            
                
        On Error Resume Next
        Err.Clear
        
        path = directory & "\" & DirectoryByType & "\" & VBComponent.Name & extension
            Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
End Sub
