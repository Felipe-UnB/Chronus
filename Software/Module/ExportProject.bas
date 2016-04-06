Attribute VB_Name = "ExportProject"
Option Explicit

' Excel macro to export all VBA source code in this project to text files for proper source control versioning
' Requires enabling the Excel setting in Options/Trust Center/Trust Center Settings/Macro Settings/Trust access to the VBA project object model
'A little modification in the original code was made in order to export the different components to different folders.

'https://gist.github.com/steve-jansen/7589478

Public Sub ExportVisualBasicCode()
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
    Dim FSO As New FileSystemObject
    
    Set VBProj = Application.VBE.ActiveVBProject
    
    directory = "D:\UnB\Projetos-Software\Chronus\Software" & "\"
    count = 0

    If Not FSO.folderexists(directory) Then
        Call FSO.createfolder(directory)
        
            Call FSO.createfolder(directory & "ClassModule")
                Call FSO.createfolder(directory & "Form")
                    Call FSO.createfolder(directory & "Module")
                        Call FSO.createfolder(directory & "Other")
    End If
    
        If Not FSO.folderexists(directory & "ClassModule") Then
            Call FSO.createfolder(directory & "ClassModule")
        End If
    
            If Not FSO.folderexists(directory & "Form") Then
                Call FSO.createfolder(directory & "Form")
            End If
    
                If Not FSO.folderexists(directory & "Module") Then
                    Call FSO.createfolder(directory & "Module")
                End If
            
                    If Not FSO.folderexists(directory & "Other") Then
                        Call FSO.createfolder(directory & "Other")
                    End If
    
    Set FSO = Nothing
    
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
        
        path = directory & DirectoryByType & "\" & VBComponent.Name & extension
            Call VBComponent.Export(path)
        
        If Err.Number <> 0 Then
            Call MsgBox("Failed to export " & VBComponent.Name & " to " & path, vbCritical)
        Else
            count = count + 1
            Debug.Print "Exported " & Left$(VBComponent.Name & ":" & Space(Padding), Padding) & path
        End If

        On Error GoTo 0
    Next
    
    If MsgBox("Also export Chronus.xlam?", vbYesNo, "Chronus.xlam") = vbYes Then
        Call ExportChronusXlam
    End If
    
End Sub

Sub ExportChronusXlam()
    
    Dim VBProj As VBIDE.VBProject

'    Set VBProj = Application.VBE.ActiveVBProject
'
'    VBProj.SaveAs ("D:\UnB\Projetos Software\Chronus\Software\Chronus.xlam")

    'Declare Variables
    Dim FSO As FileSystemObject
    Dim sFile As String
    Dim sSFolder As String
    Dim sDFolder As String
    Dim CopiedMsg As String
    Dim DuplicateMsg As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    CopiedMsg = "Specified file copied successfully!"
    DuplicateMsg = "Specified file already exists in the destination folder. Shoul it be overwritten?"
    
    'This is Your File Name which you want to Copy
    sFile = ChronusNameVersion
    
    'Change to match the source folder path
    sSFolder = "C:\Users\Felipe V\AppData\Roaming\Microsoft\AddIns\"
    
    'Change to match the destination folder path
    sDFolder = "D:\UnB\Projetos-Software\Chronus\Software\"
  
    'Checking If File Is Located in the Source Folder
    If Not FSO.FileExists(sSFolder & sFile) Then
        MsgBox "Specified File Not Found", vbInformation, "Not Found"
        
        'Copying If the Same File is Not Located in the Destination Folder
        ElseIf Not FSO.FileExists(sDFolder & sFile) Then
            FSO.CopyFile (sSFolder & sFile), sDFolder, True
            MsgBox CopiedMsg, vbInformation, "Done!"
            
            Else
                If MsgBox(DuplicateMsg, _
                vbYesNo, "File already exists") = vbYes Then
                    
                    FSO.CopyFile (sSFolder & sFile), sDFolder, True
                        MsgBox CopiedMsg, vbInformation, "Done!"
                        
                End If
    End If

End Sub
