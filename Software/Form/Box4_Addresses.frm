VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box4_Addresses 
   Caption         =   "Addresses"
   ClientHeight    =   4785
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
                Isotope232analyzed = False
    Else
        RefEdit5_232.Enabled = True
            RefEdit11_232Header.Enabled = True
                Isotope232analyzed = True
    End If
    
End Sub

Private Sub CheckBox3_Change()
    
    If CheckBox3.Value = False Then
        RefEdit20_208.Enabled = False
        RefEdit20_208.Value = ""
            RefEdit21_208Header.Enabled = False
            RefEdit21_208Header.Value = ""
                Isotope208analyzed = False
    Else
        RefEdit20_208.Enabled = True
            RefEdit21_208Header.Enabled = True
                Isotope208analyzed = True
    End If
    
End Sub

Private Sub UserForm_Initialize()

    If SampleName_UPb Is Nothing Then
        Call PublicVariables
    End If
    
    'Page Address controls
    Set RawHg202 = Box4_Addresses.RefEdit1_202
    Set RawPb204 = Box4_Addresses.RefEdit2_204
    Set RawPb206 = Box4_Addresses.RefEdit3_206
    Set RawPb207 = Box4_Addresses.RefEdit4_207
    Set RawPb208 = Box4_Addresses.RefEdit20_208
    Set RawTh232 = Box4_Addresses.RefEdit5_232
    Set RawU238 = Box4_Addresses.RefEdit6_238
    Set RawHg202Header = Box4_Addresses.RefEdit7_202Header
    Set RawPb204Header = Box4_Addresses.RefEdit8_204Header
    Set RawPb206Header = Box4_Addresses.RefEdit9_206Header
    Set RawPb207Header = Box4_Addresses.RefEdit10_207Header
    Set RawPb208Header = Box4_Addresses.RefEdit21_208Header
    Set RawTh232Header = Box4_Addresses.RefEdit11_232Header
    Set RawU238Header = Box4_Addresses.RefEdit12_238Header
    Set RawCyclesTime = Box4_Addresses.RefEdit15_CyclesTime
    Set AnalysisDate = Box4_Addresses.RefEdit22_AnalysisDate

    'Code to set the ranges for the address of each isotope signal in raw data file based on Start-AND-Options sheet
'    RawHg202.Value = RawHg202Range.Value
'    RawPb204.Value = RawPb204Range.Value
'    RawPb206.Value = RawPb206Range.Value
'    RawPb207.Value = RawPb207Range.Value
'    RawPb208.Value = RawPb208Range.Value
'    RawTh232.Value = RawTh232Range.Value
'    RawU238.Value = RawU238Range.Value
'    RawCyclesTime.Value = RawCyclesTimeRange.Value
'    AnalysisDate.Value = AnalysisDateRange.Value
'    RawHg202Header.Value = RawHg202HeaderRange.Value
'    RawPb204Header.Value = RawPb204HeaderRange.Value
'    RawPb206Header.Value = RawPb206HeaderRange.Value
'    RawPb207Header.Value = RawPb207HeaderRange.Value
'    RawPb208Header.Value = RawPb208HeaderRange.Value
'    RawTh232Header.Value = RawTh232HeaderRange.Value
'    RawU238Header.Value = RawU238HeaderRange.Value

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
    If Isotope208analyzed = True And Isotope232analyzed = True Then
        AddressRawDataFile = Array(RawHg202, RawPb204, RawPb206, RawPb207, RawPb208, _
        RawTh232, RawU238, RawHg202Header, RawPb204Header, RawPb206Header, RawPb207Header, _
        RawPb208Header, RawTh232Header, RawU238Header)
    ElseIf Isotope208analyzed = True And Isotope232analyzed = False Then
        AddressRawDataFile = Array(RawHg202, RawPb204, RawPb206, RawPb207, RawPb208, RawU238, _
        RawHg202Header, RawPb204Header, RawPb206Header, RawPb207Header, RawPb208Header, _
        RawU238Header)
    ElseIf Isotope208analyzed = False And Isotope232analyzed = True Then
        AddressRawDataFile = Array(RawHg202, RawPb204, RawPb206, RawPb207, RawTh232, RawU238, _
        RawHg202Header, RawPb204Header, RawPb206Header, RawPb207Header, RawTh232Header, _
        RawU238Header)
    ElseIf Isotope208analyzed = False And Isotope232analyzed = False Then
        AddressRawDataFile = Array(RawHg202, RawPb204, RawPb206, RawPb207, RawU238, _
        RawHg202Header, RawPb204Header, RawPb206Header, RawPb207Header, _
        RawU238Header)
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
        
    'The commands below copy the values from the refedit controls to the correct ranges in Start-AND-Option sheet.
    'Pay attention that the program do not copy the same value, it copies only the cell reference and not all the
    'cell address (including sheet name)
    
    RawHg202Range = Right(RawHg202.Value, Len(RawHg202.Value) - InStr(RawHg202.Value, "!"))
    RawPb204Range = Right(RawPb204.Value, Len(RawPb204.Value) - InStr(RawPb204.Value, "!"))
    RawPb206Range = Right(RawPb206.Value, Len(RawPb206.Value) - InStr(RawPb206.Value, "!"))
    RawPb207Range = Right(RawPb207.Value, Len(RawPb207.Value) - InStr(RawPb207.Value, "!"))
    RawPb208Range = Right(RawPb208.Value, Len(RawPb208.Value) - InStr(RawPb208.Value, "!"))
    RawTh232Range = Right(RawTh232.Value, Len(RawTh232.Value) - InStr(RawTh232.Value, "!"))
    RawU238Range = Right(RawU238.Value, Len(RawU238.Value) - InStr(RawU238.Value, "!"))
    RawCyclesTimeRange = Right(RawCyclesTime.Value, Len(RawCyclesTime.Value) - InStr(RawCyclesTime.Value, "!"))
    AnalysisDateRange = Right(AnalysisDate.Value, Len(AnalysisDate.Value) - InStr(AnalysisDate.Value, "!"))
    RawHg202HeaderRange = Right(RawHg202Header.Value, Len(RawHg202Header.Value) - InStr(RawHg202Header.Value, "!"))
    RawPb204HeaderRange = Right(RawPb204Header.Value, Len(RawPb204Header.Value) - InStr(RawPb204Header.Value, "!"))
    RawPb206HeaderRange = Right(RawPb206Header.Value, Len(RawPb206Header.Value) - InStr(RawPb206Header.Value, "!"))
    RawPb207HeaderRange = Right(RawPb207Header.Value, Len(RawPb207Header.Value) - InStr(RawPb207Header.Value, "!"))
    RawPb208HeaderRange = Right(RawPb208Header.Value, Len(RawPb208Header.Value) - InStr(RawPb208Header.Value, "!"))
    RawTh232HeaderRange = Right(RawTh232Header.Value, Len(RawTh232Header.Value) - InStr(RawTh232Header.Value, "!"))
    RawU238HeaderRange = Right(RawU238Header.Value, Len(RawU238Header.Value) - InStr(RawU238Header.Value, "!"))

    If CheckBox3.Value = False Then
        Isotope208analyzed = False
    Else
        Isotope208analyzed = True
    End If

    If CheckBox2.Value = False Then
        Isotope232analyzed = False
    Else
        Isotope232analyzed = True
    End If

    Box4_Addresses.Hide
            
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim Response As Integer
    
    If CloseMode = vbFormControlMenu Then

        Box4_Addresses.Hide
    
    End If
    
End Sub
