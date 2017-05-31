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


Private m_blnKeepOnTopOfAll As Boolean
Private m_blnKeepOnTopOfApplication As Boolean
'API functions
Private Declare Function SetWindowPos Lib "user32" _
                                      (ByVal hwnd As Long, _
                                       ByVal hWndInsertAfter As Long, _
                                       ByVal x As Long, _
                                       ByVal y As Long, _
                                       ByVal cx As Long, _
                                       ByVal cy As Long, _
                                       ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" _
                                    Alias "FindWindowA" _
                                    (ByVal lpClassName As String, _
                                     ByVal lpWindowName As String) As Long
'Constants
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_SHOWWINDOW = &H40
Private Const WS_EX_APPWINDOW = &H40000

Private Sub UserForm_Activate()
   If m_blnKeepOnTopOfAll Then
      KeepMeOnTopOfAll
   ElseIf m_blnKeepOnTopOfApplication Then
      KeepMeOnTopOfApp
   End If
End Sub

Private Sub KeepMeOnTopOfAll()
'Keep this userform on top of all other windows
    Dim WStyle As Long
    Dim Result As Long
    Dim hwnd As Long

    hwnd = FindWindow(vbNullString, Me.Caption)
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
End Sub
Private Sub KeepMeOnTopOfApp()
'Keep this userform on top of all other windows
    Dim WStyle As Long
    Dim Result As Long
    Dim hwnd As Long

    hwnd = FindWindow(vbNullString, Me.Caption)
    SetWindowPos hwnd, HWND_TOP, 0, 0, 0, 0, _
        SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    
End Sub

Public Property Get KeepOnTopOfAll() As Boolean

   KeepOnTopOfAll = m_blnKeepOnTopOfAll
   
End Property

Public Property Let KeepOnTopOfAll(ByVal blnKeepOnTopOfAll As Boolean)

   m_blnKeepOnTopOfAll = blnKeepOnTopOfAll
   m_blnKeepOnTopOfApplication = False
   If Me.Visible Then UserForm_Activate

End Property

Public Property Get KeepOnTopOfApplication() As Boolean

   KeepOnTopOfApplication = m_blnKeepOnTopOfApplication
End Property

Public Property Let KeepOnTopOfApplication(ByVal blnKeepOnTopOfApplication As Boolean)
   
   m_blnKeepOnTopOfAll = False
   m_blnKeepOnTopOfApplication = blnKeepOnTopOfApplication
   If Me.Visible Then UserForm_Activate

End Property



'----------------------------------------


Private Sub CheckBox_232Analyzed_Change()

    If CheckBox_232Analyzed.Value = False Then
        RefEdit5_232.Enabled = False
        RefEdit5_232.Value = ""
            RefEdit11_232Header.Enabled = False
            RefEdit11_232Header.Value = ""
                Isotope232Analyzed_UPb = False
                    Box1_Start.CheckBox_232MIC = False
                    Box1_Start.CheckBox_232Faraday = False
    Else
        RefEdit5_232.Enabled = True
            RefEdit11_232Header.Enabled = True
                Isotope232Analyzed_UPb = True
                    Box1_Start.CheckBox_232MIC = True
                    Box1_Start.CheckBox_232Faraday = True
                    
    End If
    
End Sub

Private Sub CheckBox_208Analyzed_Change()
    
    If CheckBox_208Analyzed.Value = False Then
        RefEdit20_208.Enabled = False
        RefEdit20_208.Value = ""
            RefEdit21_208Header.Enabled = False
            RefEdit21_208Header.Value = ""
                Isotope208Analyzed_UPb = False
                    Box1_Start.CheckBox_208MIC = False
                    Box1_Start.CheckBox_208Faraday = False
    Else
        RefEdit20_208.Enabled = True
            RefEdit21_208Header.Enabled = True
                Isotope208Analyzed_UPb = True
                    Box1_Start.CheckBox_208MIC = True
                    Box1_Start.CheckBox_208Faraday = True
    End If
    
End Sub

Private Sub CheckBox_NumberCycles_Change()

    If CheckBox_NumberCycles.Value = False Then
        RefEdit23_Num_Cycles.Enabled = False
        RefEdit23_Num_Cycles.Value = ""
        EachSampleNumberCycles_UPb = False
        Box1_Start.TextBox11_HowMany.Enabled = True
    Else
        RefEdit23_Num_Cycles.Enabled = True
        EachSampleNumberCycles_UPb = True
        Box1_Start.TextBox11_HowMany.Enabled = False
    End If

'    Public RawCyclesTimeRange As Range
'    Public RawAnalysisDateRange As Range
'    Public RawNumCyclesRange As Range
'
'    Set Box4_Addresses_RawCyclesTime = Box4_Addresses.RefEdit15_CyclesTime
'    Set Box4_Addresses_RawAnalysisDate = Box4_Addresses.RefEdit22_AnalysisDate
'    Set Box4_Addresses_RawNumCycles_Each_Sample = Box4_Addresses.RefEdit23_Num_Cycles

End Sub

Private Sub CheckBox_202Analyzed_Change()
    
    If CheckBox_202Analyzed.Value = False Then
        RefEdit1_202.Enabled = False
        RefEdit1_202.Value = ""
            RefEdit7_202Header.Enabled = False
            RefEdit7_202Header.Value = ""
                Isotope202Analyzed_UPb = False
                    Box1_Start.CheckBox_202MIC = False
                    Box1_Start.CheckBox_202Faraday = False
    Else
        RefEdit1_202.Enabled = True
            RefEdit7_202Header.Enabled = True
                Isotope202Analyzed_UPb = True
                    Box1_Start.CheckBox_202MIC = True
                    Box1_Start.CheckBox_202Faraday = True
    End If

End Sub

Private Sub CheckBox_204Analyzed_Change()

    If CheckBox_204Analyzed.Value = False Then
        RefEdit2_204.Enabled = False
        RefEdit2_204.Value = ""
            RefEdit8_204Header.Enabled = False
            RefEdit8_204Header.Value = ""
                Isotope204Analyzed_UPb = False
                    Box1_Start.CheckBox_204MIC = False
                    Box1_Start.CheckBox_204Faraday = False
    Else
        RefEdit2_204.Enabled = True
            RefEdit8_204Header.Enabled = True
                Isotope204Analyzed_UPb = True
                    Box1_Start.CheckBox_204MIC = True
                    Box1_Start.CheckBox_204Faraday = True
    End If

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
    Dim temp As Variant
    
    'The conditional clauses below are necessary because not all isotopes must have been analyzed
'    If Isotope208Analyzed_UPb = True And Isotope232Analyzed_UPb = True Then
'        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawPb208, _
'        Box4_Addresses_RawTh232, Box4_Addresses_RawU238, Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, _
'        Box4_Addresses_RawPb208Header, Box4_Addresses_RawTh232Header, Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
'    ElseIf Isotope208Analyzed_UPb = True And Isotope232Analyzed_UPb = False Then
'        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawPb208, Box4_Addresses_RawU238, _
'        Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, Box4_Addresses_RawPb208Header, _
'        Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
'    ElseIf Isotope208Analyzed_UPb = False And Isotope232Analyzed_UPb = True Then
'        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawTh232, Box4_Addresses_RawU238, _
'        Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, Box4_Addresses_RawTh232Header, _
'        Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
'    ElseIf Isotope208Analyzed_UPb = False And Isotope232Analyzed_UPb = False Then
'        AddressRawDataFile = Array(Box4_Addresses_RawHg202, Box4_Addresses_RawPb204, Box4_Addresses_RawPb206, Box4_Addresses_RawPb207, Box4_Addresses_RawU238, _
'        Box4_Addresses_RawHg202Header, Box4_Addresses_RawPb204Header, Box4_Addresses_RawPb206Header, Box4_Addresses_RawPb207Header, _
'        Box4_Addresses_RawU238Header, Box4_Addresses_RawCyclesTime, Box4_Addresses_RawAnalysisDate)
'    End If

'    AddressRawDataFile = Array( _
'        Box4_Addresses_RawPb206, _
'        Box4_Addresses_RawPb207, _
'        Box4_Addresses_RawU238, _
'        Box4_Addresses_RawPb206Header, _
'        Box4_Addresses_RawPb207Header, _
'        Box4_Addresses_RawU238Header, _
'        Box4_Addresses_RawCyclesTime, _
'        Box4_Addresses_RawAnalysisDate)
'
'    If Isotope202Analyzed_UPb = True Then
'        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawHg202, Box4_Addresses_RawHg202Header))
'    End If
'
'    If Isotope204Analyzed_UPb = True Then
'        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawPb204, Box4_Addresses_RawPb204Header))
'    End If
'
'    If Isotope208Analyzed_UPb = True Then
'        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawPb208, Box4_Addresses_RawPb208Header))
'    End If
'
'    If Isotope232Analyzed_UPb = True Then
'        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawTh232, Box4_Addresses_RawTh232Header))
'    End If

    AddressRawDataFile = ArrayFilledAddresses()

    'Check if all the refedit controls were used to select some address
    For Each C In AddressRawDataFile
        If C = "" Then
            MsgBoxAlert = MsgBox("Please, set all the addresses in Address tab.", vbOKOnly)
                
'            On Error Resume Next
'                Workbooks.Open FileName:=SamList_Sh.Range("A3")
'                    If Err.Number <> 0 Then
'                        MsgBox MissingFile1 & SamList_Sh.Range("A3") & MissingFile2
'                            Call UpdateFilesAddresses
'                                Call UnloadAll
'                                    End
'                    End If
'            On Error GoTo 0
                    C.SetFocus
                        Exit Sub
        End If
    Next
        
    If EachSampleNumberCycles_UPb = True And Box4_Addresses_RawNumCycles_Each_Sample = "" Then
        MsgBoxAlert = MsgBox("Please, set the addresses of the number of cycles of each sample.", vbOKOnly)
        
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
    
    If CheckBox_208Analyzed.Value = False Then
        Isotope208Analyzed_UPb = False
    Else
        Isotope208Analyzed_UPb = True
    End If

    If CheckBox_232Analyzed.Value = False Then
        Isotope232Analyzed_UPb = False
    Else
        Isotope232Analyzed_UPb = True
    End If
    
    If CheckBox_202Analyzed.Value = False Then
        Isotope202Analyzed_UPb = False
    Else
        Isotope202Analyzed_UPb = True
    End If
    
    If CheckBox_204Analyzed.Value = False Then
        Isotope204Analyzed_UPb = False
    Else
        Isotope204Analyzed_UPb = True
    End If
    
'    If CheckBox4 = False Then
'        RawNumberCycles_UPb = "" 'meaning that all analyses have the same number of cycles
'    End If

    Box4_Addresses.Hide
    Box1_Start.Show
            
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    Dim Response As Integer
    
    If CloseMode = vbFormControlMenu Then
                
        Response = MsgBox("Do you really want to end the program execution?", vbYesNo)
            If Response = vbNo Then
                Cancel = True
            ElseIf Response = vbYes Then
                Call UnloadAll
                    End
            End If
            
    End If
    
End Sub

Function ArrayFilledAddresses()

    Dim AddressRawDataFile As Variant
    Dim temp As Variant

    AddressRawDataFile = Array( _
        Box4_Addresses_RawPb206, _
        Box4_Addresses_RawPb207, _
        Box4_Addresses_RawU238, _
        Box4_Addresses_RawPb206Header, _
        Box4_Addresses_RawPb207Header, _
        Box4_Addresses_RawU238Header, _
        Box4_Addresses_RawCyclesTime, _
        Box4_Addresses_RawAnalysisDate)

    If Isotope202Analyzed_UPb = True Then
        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawHg202, Box4_Addresses_RawHg202Header))
    End If
    
    If Isotope204Analyzed_UPb = True Then
        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawPb204, Box4_Addresses_RawPb204Header))
    End If
    
    If Isotope208Analyzed_UPb = True Then
        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawPb208, Box4_Addresses_RawPb208Header))
    End If
    
    If Isotope232Analyzed_UPb = True Then
        temp = ConcatenateArrays(AddressRawDataFile, Array(Box4_Addresses_RawTh232, Box4_Addresses_RawTh232Header))
    End If
    
    ArrayFilledAddresses = AddressRawDataFile

End Function
