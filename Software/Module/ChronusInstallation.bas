Attribute VB_Name = "ChronusInstallation"
Option Explicit

Sub CheckAccessVBPM()

    'Created 21122015
    'This procedure checks if the trust Access to Visual Basic Project
    'Object Model option is enabled
    
    Dim VBPrj As VBProject
    
    On Error Resume Next
        
        Set VBPrj = ThisWorkbook.VBProject
        
        If Err <> 0 Then
            MsgBox "You must enable Chronus to access the VBA project object model."
                On Error GoTo 0
                    Call UnloadAll
                        End
        End If

    On Error GoTo 0
    
End Sub
Sub CheckIsoplotReference()

    'Check if Isoplot is installed and loaded.

    Dim i As Long
    Dim IsoplotFound As Boolean
    
    IsoplotFound = False
    
    For i = 1 To Application.VBE.ActiveVBProject.References.count
        
        If Application.VBE.ActiveVBProject.References(i).Name = "Isoplot4" Then
            IsoplotFound = True
                
                i = Application.VBE.ActiveVBProject.References.count
        End If
        
    Next i
        
    If IsoplotFound = False Then
        MsgBox "Chronus - U-Pb data reduction. " & _
        "Please, install isoplot before Chronus.", vbOKOnly
            Call UnloadAll
                End
    End If
    
End Sub

Sub AddReference()
    'Macro purpose:  To add a reference to the project using the GUID for the
    'reference library
    
    'This procedure is a modification of the original code by Ken Puls available
    'at the address below.
    'http://www.vbaexpress.com/kb/getarticle.php?kb_id=267
    
    Dim strGUID(1 To 8, 1 To 3) As Variant
    Dim theRef As Reference
    Dim Counter As Long
    
    'Update the GUID you need below.
    strGUID(1, 1) = "{000204EF-0000-0000-C000-000000000046}" 'VBA- V4.1
    strGUID(1, 2) = 4
    strGUID(1, 3) = 1
    
    strGUID(2, 1) = "{00020813-0000-0000-C000-000000000046}" 'Excel-V1.7
    strGUID(2, 2) = 1
    strGUID(2, 3) = 7
    
    strGUID(3, 1) = "{00020430-0000-0000-C000-000000000046}" 'stdole-V2.0
    strGUID(3, 2) = 2
    strGUID(3, 3) = 0
    
    strGUID(4, 1) = "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}" 'Office-V2.5
    strGUID(4, 2) = 2
    strGUID(4, 3) = 5
    
    strGUID(5, 1) = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}" 'MSForms-V2.0
    strGUID(5, 2) = 2
    strGUID(5, 3) = 0
    
    strGUID(6, 1) = "{00024517-0000-0000-C000-000000000046}" 'RefEdit-V1.2
    strGUID(6, 2) = 1
    strGUID(6, 3) = 2
    
    strGUID(7, 1) = "{0002E157-0000-0000-C000-000000000046}" 'VBIDE-V5.3
    strGUID(7, 2) = 5
    strGUID(7, 3) = 3
    
    strGUID(8, 1) = "{420B2830-E718-11CF-893D-00A0C9054228}" 'Scripting-V1.0
    strGUID(8, 2) = 1
    strGUID(8, 3) = 0
    
    'Set to continue in case of error
    On Error Resume Next
    
        'Remove any missing references
        For Counter = ThisWorkbook.VBProject.References.count To 1 Step -1
            Set theRef = ThisWorkbook.VBProject.References.Item(Counter)
                If theRef.IsBroken = True Then
                    ThisWorkbook.VBProject.References.Remove theRef
                End If
        Next Counter
        
        'Clear any errors so that error trapping for GUID additions can be evaluated
        Err.Clear
        
        'Add the references
        For Counter = 1 To UBound(strGUID)
            ThisWorkbook.VBProject.References.AddFromGuid _
            GUID:=strGUID(Counter, 1), Major:=strGUID(Counter, 2), Minor:=strGUID(Counter, 3)
        Next Counter
        
        'If an error was encountered, inform the user
        Select Case Err.Number
            Case Is = 32813 'Reference already in use.  No action necessary
            
            Case Is = vbNullString 'Reference added without issue
            
            Case Else 'An unknown error was encountered, so alert the user
            
                MsgBox "A problem was encountered trying to" & vbNewLine _
                & "add or remove a reference in this file" & vbNewLine & "Please check the " _
                & "references in your VBA project!", vbCritical + vbOKOnly, "Error!"
                
                    End
        
        End Select
    
    On Error GoTo 0
    
End Sub

Sub Display_GUID_Info()
    
    'PURPOSE: Displays GUID information for each active _
    Object Library reference in the VBA project
    'SOURCE: www.TheSpreadsheetGuru.com
    
    'Code from http://www.thespreadsheetguru.com/the-code-vault/2014/3/16/display-object-library-reference-guid-information
    'This procedure is used to find the GUID information os the references that are added by the AddReference procedure.
    
    Dim ref As Reference
         
    'Loop Through Each Active Reference (Displays in Immediate Window [ctrl + g])
    For Each ref In ThisWorkbook.VBProject.References
      Debug.Print "Reference Name: ", ref.Name
      Debug.Print "Path: ", ref.FullPath
      Debug.Print "GUID: " & ref.GUID
      Debug.Print "Version: " & ref.Major & "." & ref.Minor
      Debug.Print " "
    Next ref
  
End Sub

