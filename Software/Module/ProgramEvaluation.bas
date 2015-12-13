Attribute VB_Name = "ProgramEvaluation"
Option Explicit

Sub GetWMIComputerInfo()
' http://msdn.microsoft.com/en-us/library/aa394102(VS.85).aspx
Dim objWMI As Object
Dim CompStatus As Object
Dim objProcess As Object
Dim sngProcessTime As Long
Dim objRefresher As Object
 
  Set objWMI = GetWMIService
    
  Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")

    Set CompStatus = objWMI.ExecQuery("Select * from Win32_Process WHERE Name = 'EXCEL.EXE'")

        For Each objProcess In CompStatus

            If objProcess.Name = "EXCEL.EXE" Then

                sngProcessTime = (CSng(objProcess.KernelModeTime) + CSng(objProcess.UserModeTime)) / 10000000

                Debug.Print "Working Set Size: " & objProcess.WorkingSetSize / 1000000
                        Exit For

            End If

        Next
        
    Set CompStatus = objRefresher.AddEnum _
    (objWMI, "Win32_PerfFormattedData_PerfProc_Process").objectSet
    
        For Each objProcess In CompStatus
        
            If objProcess.Name = "EXCEL.EXE" Then
            
                objRefresher.Refresh
            
                Debug.Print "Percent Processor Time: " & CompStatus.PercentProcessorTime
                Debug.Print "Memory: " & CompStatus.WorkingSet
                Debug.Print "Private Memory: " & CompStatus.PrivateBytes
                    Debug.Print
                        Exit For

            End If
        
        Next
 
End Sub

Function GetWMIService() As Object
' http://msdn.microsoft.com/en-us/library/aa394586(VS.85).aspx
Dim strComputer As String
 
  strComputer = "."
 
  Set GetWMIService = GetObject("winmgmts:" _
                              & "{impersonationLevel=impersonate}!\\" _
                              & strComputer & "\root\cimv2")
End Function

Sub testttttttsets()

    Dim strComputer As String
    Dim objWMI As Object
Dim CompStatus As Object
Dim objProcess As Object
Dim sngProcessTime As Long
Dim objRefresher As Object
Dim i As Long
    
    Set objWMI = GetWMIService
    
    Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
    Set CompStatus = objRefresher.AddEnum _
        (objWMI, "Win32_PerfFormattedData_PerfProc_Process").objectSet
    
    For i = 1 To 30
        objRefresher.Refresh
        For Each objProcess In CompStatus
            If objProcess.Name = "EXCEL" Then
            
                Debug.Print "PercentProcessorTime" & " -- " & objProcess.PercentProcessorTime
                Debug.Print "WorkingSet" & " -- " & objProcess.WorkingSet / 1000000
                Debug.Print "PrivateBytes" & " -- " & objProcess.PrivateBytes / 1000000
                
            End If
        Next
        Debug.Print
    
        Application.Wait (Now + TimeValue("0:00:01"))
        
    Next

End Sub
