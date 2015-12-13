VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box5_DataFilter 
   Caption         =   "Data filtering"
   ClientHeight    =   8430.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4890
   OleObjectBlob   =   "Box5_DataFilter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Box5_DataFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'******************************************************
'/----------------------------------------------------\
'|  Macro desenvolvida por: Felipe Valença de Oliveira|
'|  Laboratório de Geocronologia - UnB                |
'|  Primeira versão (v1): Janeiro - 2014              |
'\----------------------------------------------------/
'******************************************************


Private Sub CheckBox1_Error75_Click()
    
    If CheckBox1_Error75 = True Then
        Error75Entry.Enabled = True
            Error75Entry.Value = 5
                Error75Entry.BackColor = vbWhite
    Else
        Error75Entry.Enabled = False
            Error75Entry.Value = ""
                Error75Entry.BackColor = &H8000000F
    End If
        
End Sub

Private Sub CheckBox2_Rho_Click()

    If CheckBox2_Rho = True Then
        RhoEntry.Enabled = True
            RhoEntry.Value = 0.5
                RhoEntry.BackColor = vbWhite
    Else
        RhoEntry.Enabled = False
            RhoEntry.Value = ""
                RhoEntry.BackColor = &H8000000F
    End If

End Sub

Private Sub CheckBox3_F206_Click()

    If CheckBox3_F206 = True Then
        f206Entry.Enabled = True
            f206Entry.Value = 3
                f206Entry.BackColor = vbWhite
    Else
        f206Entry.Enabled = False
            f206Entry.Value = ""
                f206Entry.BackColor = &H8000000F
    End If

End Sub

Private Sub IgneousRock_Click()
    
    TextBoxMin = 95
    TextBoxMax = 105
    TextBoxAge68 = 1000
    
End Sub

Private Sub SedimentaryRock_Click()
    
    TextBoxMin = 90
    TextBoxMax = 110
    TextBoxAge68 = 1000

End Sub

Private Sub TextBoxAge68_Change()
    If AdvancedFilters.Value = True Then
        If Not IsNumeric(TextBoxAge68.Text) Then
            MsgBox "Please enter numbers only.", vbInformation
                TextBoxAge68.SelStart = 0
                    TextBoxAge68.SelLength = Len(TextBoxAge68.Text)
        End If
    End If

        
End Sub

Private Sub UserForm_Initialize()
    
    If InstalledIsoplot = False Then
        MsgBox "Isoplot 4.15 must be installed and enabled in this " & _
            "computer in order to use Data Filter.", vbOKOnly
            Call UnloadAll
                End
    End If
    
    If SlpStdCorr_Sh Is Nothing Then
        Call PublicVariables
    End If
    
    If ActiveSheet.Name <> SlpStdCorr_Sh.Name Then
        MsgBox "This tool is only available for SlpStdCorr sheet."
            Call UnloadAll
                End
    End If
    
    MsgBox "207Pb/235U ratio error must be in percentage."
        Call ConvertUncertantiesTo("Percentage", SlpStdCorr_Sh)
            Call FormatSlpStdCorr(False)
    
    CheckBox1_Error75.Value = True
        CheckBox2_Rho.Value = True
            CheckBox3_F206.Value = True
    
    Error75Entry.Value = 5 'Default Error75Entry number.
        Error75Entry.Enabled = True
        
    f206Entry.Value = 3 'Default f206EntryEntry number.
        f206Entry.Enabled = True
        
    RhoEntry.Value = 0.5 'Default RhoEntry number.
        RhoEntry.Enabled = True
        
End Sub

Private Sub f206Entry_Change() 'f206Entry validation
    
    If CheckBox3_F206.Value = True Then
        If Not IsNumeric(f206Entry.Text) Then
            MsgBox "Please enter numbers only.", vbInformation
                f206Entry.SelStart = 0
                    f206Entry.SelLength = Len(f206Entry.Text)
        End If
    End If
    
End Sub

Private Sub RhoEntry_Change() 'RhoEntry validation

    If CheckBox2_Rho.Value = True Then
        If Not IsNumeric(RhoEntry.Text) Then
            MsgBox "Please enter numbers only.", vbInformation
                RhoEntry.SelStart = 0
                    RhoEntry.SelLength = Len(RhoEntry.Text)
        End If
    End If
    
End Sub

Private Sub Error75Entry_Change() 'Error75Entry validation
    
    If CheckBox1_Error75.Value = True Then
        If Not IsNumeric(Error75Entry.Text) Then
            MsgBox "Please enter numbers only.", vbInformation
                Error75Entry.SelStart = 0
                    Error75Entry.SelLength = Len(Error75Entry.Text)
        End If
    End If
    
End Sub

Private Sub TextBoxMin_Change() 'TextBoxMin validation
    
    If AdvancedFilters.Value = True Then
        If Not IsNumeric(TextBoxMin.Text) Then
            MsgBox "Please enter numbers only.", vbInformation
                TextBoxMin.SelStart = 0
                    TextBoxMin.SelLength = Len(TextBoxMin.Text)
        End If
    End If
    
End Sub

Private Sub TextBoxMax_Change() 'TextBoxMax validation
    
    If AdvancedFilters.Value = True Then
        If Not IsNumeric(TextBoxMax.Text) Then
            MsgBox "Please enter numbers only.", vbInformation
                TextBoxMax.SelStart = 0
                    TextBoxMax.SelLength = Len(TextBoxMax.Text)
        End If
    End If
    
End Sub

Private Sub AdvancedFilters_Click() 'Option buttons enabled or not
    If AdvancedFilters.Value = True Then
               
        With IgneousRock
            .Enabled = True
            .Value = True
            
            TextBoxMin = 95
            TextBoxMax = 105
            TextBoxAge68 = 1000
        End With
        
        SedimentaryRock.Enabled = True
            
        TextBoxMin.Enabled = True
            TextBoxMax.Enabled = True
                TextBoxAge68.Enabled = True
    
        TextBoxMin.BackColor = vbWhite
            TextBoxMax.BackColor = vbWhite
                TextBoxAge68.BackColor = vbWhite

    Else
        IgneousRock.Enabled = False
            SedimentaryRock.Enabled = False
            
        With TextBoxMin
            .Value = ""
            .Enabled = False
            .BackColor = &H8000000F
        End With
        
        With TextBoxMax
            .Value = ""
            .Enabled = False
            .BackColor = &H8000000F
        End With
        
        With TextBoxAge68
            .Value = ""
            .Enabled = False
            .BackColor = &H8000000F
            End With
    End If
End Sub

Private Sub Ok_Click()
    Dim Error75 As Single 'Criteria indicating by to user to ignore grains with 207/235 error ratios too big.
    Dim Rho As Single 'Minimum error concordance between ratios 206/238 and 207/235 indicated by user.
    Dim f206 As Single 'Maximum Common Pb contents acceptable indicated by user.
    Dim MinValue As Single 'Minimum concorcande indicated by user
    Dim MaxValue As Single 'Maximum concordance indicated by user
    Dim Age68Limit As Single 'Age 206/208 limit, used to determine which ages concordance will be
                             'considered (206/238 and 207/206 or 206/238 and 207/235)
    Dim Conc6875 As Double 'Concordance between ages 68 and 75 (68 is the numerator)
    Dim Conc6876 As Double 'Concordance between ages 68 and 76 (68 is the numerator)
            
    Dim Error75Range As Range 'Range with ratio 207/235 errors
    Dim Range75 ' range with ratio 207/235
    Dim RhoRange As Range 'Range with Rho values
    Dim f206Range As Range 'Range with f206 values
    Dim Age68 As Range
    
    Dim CellInRange As Range 'Cells inside range that will be checked
                
    SlpStdCorr_Sh.Cells.Font.Strikethrough = False
    
        Error75 = Error75Entry.Value
        Rho = RhoEntry.Value
        f206 = f206Entry.Value
        
        If AdvancedFilters.Value = True Then
            MinValue = TextBoxMin.Value
            MaxValue = TextBoxMax.Value
            Age68Limit = TextBoxAge68.Value
        End If
        
        With SlpStdCorr_Sh
            Set Error75Range = .Range(StdCorr_Column751Std & StdCorr_HeaderRow + 1, .Range(StdCorr_Column751Std & StdCorr_HeaderRow + 1).End(xlDown))
            Set Range75 = .Range(StdCorr_Column75 & StdCorr_HeaderRow + 1, .Range(StdCorr_Column75 & StdCorr_HeaderRow + 1).End(xlDown))
            Set RhoRange = .Range(StdCorr_Column7568Rho & StdCorr_HeaderRow + 1, .Range(StdCorr_Column7568Rho & StdCorr_HeaderRow + 1).End(xlDown))
            Set f206Range = .Range(StdCorr_ColumnF206 & StdCorr_HeaderRow + 1, .Range(StdCorr_ColumnF206 & StdCorr_HeaderRow + 1).End(xlDown))
            Set Age68 = .Range(StdCorr_Column68AgeMa & StdCorr_HeaderRow + 1, .Range(StdCorr_Column68AgeMa & StdCorr_HeaderRow + 1).End(xlDown))
        End With
               
            'Procedures below will check Ratio Error75, Rho and f206 based on criteria indicated by user.
                            
        For Each CellInRange In Error75Range
        
            If IsNumeric(CellInRange) = True And CellInRange > Error75 Then
                CellInRange.EntireRow.Font.Strikethrough = True
            End If
            
        Next
        
            For Each CellInRange In RhoRange
            
                If IsNumeric(CellInRange) = True And CellInRange < Rho Then
                    CellInRange.EntireRow.Font.Strikethrough = True
                End If

            Next
            
                 For Each CellInRange In f206Range
                 
                    If IsNumeric(CellInRange) = True And CellInRange > f206 Then
                        CellInRange.EntireRow.Font.Strikethrough = True
                    End If
                
                Next
                            
        'Considering that the type of rock is known, additional criteria will be used to filter the data.
        'The procecures below are resposinble for this.
        
    
'    Public Const StdCorr_Column68AgeMa As String = "X"
'    Public Const StdCorr_Column68AgeMa1std As String = "Y"
'    Public Const StdCorr_Column75AgeMa As String = "Z"
'    Public Const StdCorr_Column75AgeMa1std As String = "AA"
'    Public Const StdCorr_Column76AgeMa As String = "AB"
'    Public Const StdCorr_Column76AgeMa1std As String = "AC"

    
        If IgneousRock = True Then 'If it's an igneous rock.
            For Each CellInRange In Age68
                    
                    If IsNumeric(CellInRange) = True And CellInRange <= Age68Limit Then
                        
                        With SlpStdCorr_Sh
                            Conc6875 = 100 * _
                                      (.Range(StdCorr_Column68AgeMa & CellInRange.Row) / _
                                       .Range(StdCorr_Column75AgeMa & CellInRange.Row))
                        End With
                        
                        Select Case Conc6875 '68-75 concordance was chosen
                            
                            Case Is < MinValue
                                CellInRange.EntireRow.Font.Strikethrough = True
                            
                                Case Is >= MaxValue
                                    CellInRange.EntireRow.Font.Strikethrough = True
                        End Select
                    End If
                                    
                    If IsNumeric(CellInRange) = True And CellInRange > Age68Limit Then
                        
                        With SlpStdCorr_Sh
                            Conc6876 = 100 * _
                                      (.Range(StdCorr_Column68AgeMa & CellInRange.Row) / _
                                       .Range(StdCorr_Column76AgeMa & CellInRange.Row))
                        End With
                        
                        Select Case Conc6876 '68/76 concordance was chosen
                            
                            Case Is < MinValue
                                CellInRange.EntireRow.Font.Strikethrough = True
                            
                                Case Is >= MaxValue
                                    CellInRange.EntireRow.Font.Strikethrough = True
                        End Select
                    End If
10          Next
        
    ElseIf SedimentaryRock = True Then
            For Each CellInRange In Age68
                    If IsNumeric(CellInRange) = True And CellInRange <= Age68Limit Then
                        
                        With SlpStdCorr_Sh
                            Conc6875 = 100 * _
                                      (.Range(StdCorr_Column68AgeMa & CellInRange.Row) / _
                                       .Range(StdCorr_Column75AgeMa & CellInRange.Row))
                        End With

                        Select Case Conc6875 '68-75 concordance was chosen
                            
                            Case Is < MinValue
                                CellInRange.EntireRow.Font.Strikethrough = True
                            
                                Case Is >= MaxValue
                                    CellInRange.EntireRow.Font.Strikethrough = True
                        End Select
                    End If
                    
                    If IsNumeric(CellInRange) = True And CellInRange > Age68Limit Then

                        With SlpStdCorr_Sh
                            Conc6876 = 100 * _
                                      (.Range(StdCorr_Column68AgeMa & CellInRange.Row) / _
                                       .Range(StdCorr_Column76AgeMa & CellInRange.Row))
                        End With

                        Select Case Conc6876 '68/76 concordance was chosen
                            
                            Case Is < MinValue
                                CellInRange.EntireRow.Font.Strikethrough = True
                            
                                Case Is >= MaxValue
                                    CellInRange.EntireRow.Font.Strikethrough = True
                        End Select
                    End If
20          Next
            
        End If
                
    Unload Box5_DataFilter
            
End Sub

Private Sub Cancel_Click()

    Unload Box5_DataFilter
    
End Sub
