VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box4_Addresses 
   Caption         =   "Addresses"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3840
   OleObjectBlob   =   "Box4_Addresses.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Box4_Addresses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox2_Change()

    If CheckBox2.Value = False Then
        RefEdit5_232.Enabled = False
        RefEdit5_232.Value = ""
            RefEdit11_232Header.Enabled = False
            RefEdit11_232Header.Value = ""
                Isotope232Analyzed_UPb = False
    Else
        RefEdit5_232.Enabled = True
            RefEdit11_232Header.Enabled = True
                Isotope232Analyzed_UPb = True
    End If
    
End Sub

Private Sub CheckBox3_Change()
    
    If CheckBox3.Value = False Then
        RefEdit23_Num_Cycles.Enabled = False
        RefEdit23_Num_Cycles.Value = ""
            RefEdit21_208Header.Enabled = False
            RefEdit21_208Header.Value = ""
                Isotope208Analyzed_UPb = False
    Else
        RefEdit20_208.Enabled = True
            RefEdit21_208Header.Enabled = True
                Isotope208Analyzed_UPb = True
    End If
    
End Sub

Private Sub CheckBox4_Change()

    If CheckBox4.Value = False Then
        RefEdit23_Num_Cycles.Enabled = False
        RefEdit23_Num_Cycles.Value = ""
        EachSampleNumberCycles_UPb = False
    Else
        RefEdit23_Num_Cycles.Enabled = True
        EachSampleNumberCycles_UPb = True
    End If

'    Public RawCyclesTimeRange As Range
'    Public RawAnalysisDateRange As Range
'    Public RawNumCyclesRange As Range
'
'    Set Box4_Addresses_RawCyclesTime = Box4_Addresses.RefEdit15_CyclesTime
'    Set Box4_Addresses_RawAnalysisDate = Box4_Addresses.RefEdit22_AnalysisDate
'    Set Box4_Addresses_RawNumCycles_Each_Sample = Box4_Addresses.RefEdit23_Num_Cycles

End Sub

Private Sub UserForm_Initialize()

    If SampleName_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    'Page Address controls
    Set Box4_Addresses_RawHg202 = Box4_Addresses.RefEdit1_202
    Set Box4_Addresses_RawPb204 = Box4_Addresses.RefEdit2_204
    Set Box4_Addresses_RawPb206 = Box4_Addresses.RefEdit3_206
    Set Box4_Addresses_RawPb207 = Box4_Addresses.RefEdit4_207
    Set Box4_Addresses_RawPb208 = Box4_Addresses.RefEdit20_208
    Set Box4_Addresses_RawTh232 = Box4_Addresses.RefEdit5_232
    Set Box4_Addresses_RawU238 = Box4_Addresses.RefEdit6_238
    Set Box4_Addresses_RawHg202Header = Box4_Addresses.RefEdit7_202Header
    Set Box4_Addresses_RawPb204Header = Box4_Addresses.RefEdit8_204Header
    Set Box4_Addresses_RawPb206Header = Box4_Addresses.RefEdit9_206Header
    Set Box4_Addresses_RawPb207Header = Box4_Addresses.RefEdit10_207Header
    Set Box4_Addresses_RawPb208Header = Box4_Addresses.RefEdit21_208Header
    Set Box4_Addresses_RawTh232Header = Box4_Addresses.RefEdit11_232Header
    Set Box4_Addresses_RawU238Header = Box4_Addresses.RefEdit12_238Header
    Set Box4_Addresses_RawCyclesTime = Box4_Addresses.RefEdit15_CyclesTime
    Set Box4_Addresses_RawAnalysisDate = Box4_Addresses.RefEdit22_AnalysisDate
    Set Box4_Addresses_RawNumCycles_Each_Sample = Box4_Addresses.RefEdit23_Num_Cycles

End Sub


Private Sub CommandButton1_Ok_Click()
    
    Dim MsgBoxAlert As Variant 'Message box for for many checks done below
    Dim a As Integer, C As Variant
    Dim AddressRawDataFile As Variant 'Array of variables with address in Box2_UPb_Options
    Dim CellsPopulated() As Single 'Array of the number of cells with values in each variable
    'of AddressRawDataFile (Below)
    ReDim CellsPopulated(1 To 1)
    Dim MinValue As Single
    Dim MaxValue As Single
    Dim RawRanges(1 To 7) As Range
    
    'The conditional clauses below are necessary because not all isotopes must have been analyzed
    If Isotope208Analyzed_UPb = True And Isotope232Analyzed_UPb = True Then
        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawPb208, _
        Box4_Addresses_RawTh232, Box4_Addresses_RawU238, Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, _
        Box4_Addresses_RawPb208Header, Box4_Addresses_RawTh232Header, Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
    ElseIf Isotope208Analyzed_UPb = True And Isotope232Analyzed_UPb = False Then
        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawPb208, Box4_Addresses_RawU238, _
        Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, Box4_Addresses_RawPb208Header, _
        Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
    ElseIf Isotope208Analyzed_UPb = False And Isotope232Analyzed_UPb = True Then
        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawTh232, Box4_Addresses_RawU238, _
        Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, Box4_Addresses_RawTh232Header, _
        Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
    ElseIf Isotope208Analyzed_UPb = False And Isotope232Analyzed_UPb = False Then
        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawU238, _
        Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, _
        Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
    End If

    'Check if all the refedit controls were used to select some address
    For Each C In AddressRawDataFile
        If C = "" Then
            MsgBoxAlert = MsgBox("Please, set all the addresses in Address tab.", vbOKOnly)
                
            On Error Resume Next
                Workbooks.Open FileName:=SamList_Sh.Range("A3")
                    If Err.Number <> 0 Then
                        MsgBox MissingFile1 & SamList_Sh.Range("A3") & MissingFile2
                            Call UpdateFilesAddresses
                                Call UnloadAll
                                    End
                    End If
            On Error GoTo 0
                    C.SetFocus
                        Exit Sub
        End If
    Next
        
    If EachSampleNumberCycles_UPb = True And Box4_Addresses_RawNumCycles_Each_Sample = "" Then
        MsgBoxAlert = MsgBox("Please, set all the addresses in Address tab.", vbOKOnly)
        Box4_Addresses_RawNumCycles_Each_Sample.SetFocus
        Exit Sub
    End If
    
    'The commands below copy the values from the refedit controls to the correct ranges in Start-AND-Option sheet.
    'Pay attention that the program do not copy the same value, it copies only the cell reference and not all the
    'cell address (including sheet name)
    
    RawHg202Range = Right(Box4_Addresses_RawHg202.Value, Len(Box4_Addresses_RawHg202.Value) - InStr(Box4_Addresses_RawHg202.Value, "!"))
    RawPb204Range = Right(Box4_Addresses_RawPb204.Value, Len(Box4_Addresses_RawPb204.Value) - InStr(Box4_Addresses_RawPb204.Value, "!"))
    RawPb206Range = Right(Box4_Addresses_RawPb206.Value, Len(Box4_Addresses_RawPb206.Value) - InStr(Box4_Addresses_RawPb206.Value, "!"))
    RawPb207Range = Right(Box4_Addresses_RawPb207.Value, Len(Box4_Addresses_RawPb207.Value) - InStr(Box4_Addresses_RawPb207.Value, "!"))
    RawPb208Range = Right(Box4_Addresses_RawPb208.Value, Len(Box4_Addresses_RawPb208.Value) - InStr(Box4_Addresses_RawPb208.Value, "!"))
    RawTh232Range = Right(Box4_Addresses_RawTh232.Value, Len(Box4_Addresses_RawTh232.Value) - InStr(Box4_Addresses_RawTh232.Value, "!"))
    RawU238Range = Right(Box4_Addresses_RawU238.Value, Len(Box4_Addresses_RawU238.Value) - InStr(Box4_Addresses_RawU238.Value, "!"))
    
    RawHg202HeaderRange = Right(Box4_Addresses_RawHg202Header.Value, Len(Box4_Addresses_RawHg202Header.Value) - InStr(Box4_Addresses_RawHg202Header.Value, "!"))
    RawPb204HeaderRange = Right(Box4_Addresses_RawPb204Header.Value, Len(Box4_Addresses_RawPb204Header.Value) - InStr(Box4_Addresses_RawPb204Header.Value, "!"))
    RawPb206HeaderRange = Right(Box4_Addresses_RawPb206Header.Value, Len(Box4_Addresses_RawPb206Header.Value) - InStr(Box4_Addresses_RawPb206Header.Value, "!"))
    RawPb207HeaderRange = Right(Box4_Addresses_RawPb207Header.Value, Len(Box4_Addresses_RawPb207Header.Value) - InStr(Box4_Addresses_RawPb207Header.Value, "!"))
    RawPb208HeaderRange = Right(Box4_Addresses_RawPb208Header.Value, Len(Box4_Addresses_RawPb208Header.Value) - InStr(Box4_Addresses_RawPb208Header.Value, "!"))
    RawTh232HeaderRange = Right(Box4_Addresses_RawTh232Header.Value, Len(Box4_Addresses_RawTh232Header.Value) - InStr(Box4_Addresses_RawTh232Header.Value, "!"))
    RawU238HeaderRange = Right(Box4_Addresses_RawU238Header.Value, Len(Box4_Addresses_RawU238Header.Value) - InStr(Box4_Addresses_RawU238Header.Value, "!"))
   
    RawNumCyclesRange = Right(Box4_Addresses_RawNumCycles_Each_Sample.Value, Len(Box4_Addresses_RawNumCycles_Each_Sample.Value) - InStr(Box4_Addresses_RawNumCycles_Each_Sample.Value, "!"))
    RawCyclesTimeRange = Right(Box4_Addresses_RawCyclesTime.Value, Len(Box4_Addresses_RawCyclesTime.Value) - InStr(Box4_Addresses_RawCyclesTime.Value, "!"))
    RawAnalysisDateRange = Right(Box4_Addresses_RawAnalysisDate.Value, Len(Box4_Addresses_RawAnalysisDate.Value) - InStr(Box4_Addresses_RawAnalysisDate.Value, "!"))
    
    If CheckBox3.Value = False Then
        Isotope208Analyzed_UPb = False
    Else
        Isotope208Analyzed_UPb = True
    End If

    If CheckBox2.Value = False Then
        Isotope232Analyzed_UPb = False
    Else
        Isotope232Analyzed_UPb = True
    End If
    
'    If CheckBox4 = False Then
'        RawNumberCycles_UPb = "" 'meaning that all analyses have the same number of cycles
'    End If

    Box4_Addresses.Hide
            
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim Response As Integer
    
    If CloseMode = vbFormControlMenu Then

        Box4_Addresses.Hide
    
    End If
    
End Sub
