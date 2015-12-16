VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Box7_FullReduction 
   Caption         =   "Chronus"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6750
   OleObjectBlob   =   "Box7_FullReduction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Box7_FullReduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CommandButton1_Click()
    
    With Me
    
        If _
            .Program0.Value = False And _
            .Program1.Value = False And _
            .Program2.Value = False And _
            .Program3.Value = False And _
            .Program4.Value = False And _
            .Program5.Value = False And _
            .Program6.Value = False And _
            .Program7.Value = False _
        Then
            If MsgBox("No procedure was selected. Would you like to select?", vbYesNo) = vbNo Then
                Call UnloadAll
                    End
            Else
                Exit Sub
            End If
        End If

        With .CommandButton1
    
            .Width = 115
            .Enabled = False
            .Left = .Left - (115 - 72) / 2
            .Caption = "Running, please wait..."
    
        End With
    
        If .Program0.Value = True Then
            .TextBox0.BackColor = vbRed
            .TextBox0p.BackColor = vbRed
        End If
        
        If .Program1.Value = True Then
            .TextBox1.BackColor = vbRed
            .TextBox1p.BackColor = vbRed
        Else
            If MsgBox _
            ("You are strongly advised to check the raw data once. Otherwise, " & _
            "unrealistic results may be calculated or even the program crashes. " & _
            "Would you like to check?", vbYesNo) _
            = vbYes Then
                .Program1.Value = True
        End If
            
        End If
        
        If .Program2.Value = True Then
            .TextBox2.BackColor = vbRed
            .TextBox2p.BackColor = vbRed
        End If
        
        If .Program3.Value = True Then
            .TextBox3.BackColor = vbRed
            .TextBox3p.BackColor = vbRed
        End If
        
        If .Program4.Value = True Then
            .TextBox4.BackColor = vbRed
            .TextBox4p.BackColor = vbRed
        End If
        
        If .Program5.Value = True Then
            .TextBox5.BackColor = vbRed
            .TextBox5p.BackColor = vbRed
        End If
        
        If .Program6.Value = True Then
            .TextBox6.BackColor = vbRed
            .TextBox6p.BackColor = vbRed
        End If

        If .Program7.Value = True Then
            .TextBox7.BackColor = vbRed
            .TextBox7p.BackColor = vbRed
        End If
        
        .TextBox8.BackColor = vbRed
        .TextBox8p.BackColor = vbRed
        
    End With

    Call FullDataReductionNew(Program0.Value, Program1.Value, Program2.Value, Program3.Value, Program4.Value, Program5.Value, Program6.Value, Program7.Value)
    
    
End Sub

Private Sub CommandButton2_Click()
    
    Call UnloadAll
    
End Sub

Private Sub Program0_Click()
    
    If Me.Program0 = False Then
        Me.Program1 = False
        Me.Program2 = False
        Me.Program3 = False
        Me.Program4 = False
        Me.Program5 = False
        Me.Program6 = False
        Me.Program7 = False
    End If
        
End Sub

Private Sub Program1_Click()
    
    If Me.Program1 = True Then
        Me.Program0 = True
    End If

End Sub

Private Sub Program2_Click()
    
    If Me.Program2 = True Then
        Me.Program0 = True
    Else
        Me.Program3 = False
        Me.Program4 = False
        Me.Program5 = False
        Me.Program6 = False
        Me.Program7 = False
    End If
        
End Sub

Private Sub Program3_Click()
    
    If Me.Program3 = True Then
        Me.Program0 = True
        Me.Program2 = True
    Else
        Me.Program4 = False
        Me.Program5 = False
        Me.Program6 = False
        Me.Program7 = False
    End If
        
End Sub

Private Sub Program4_Click()
    
    If Me.Program4 = True Then
        Me.Program0 = True
        Me.Program2 = True
        Me.Program3 = True
    Else
        Me.Program5 = False
        Me.Program6 = False
        Me.Program7 = False
    End If
    
End Sub

Private Sub Program5_Click()

    If Me.Program5 = True Then
        Me.Program0 = True
        Me.Program2 = True
        Me.Program3 = True
        Me.Program4 = True
    Else
        Me.Program6 = False
        Me.Program7 = False
    End If
    
End Sub

Private Sub Program6_Click()
    
    If Me.Program6 = True Then
        Me.Program0 = True
        Me.Program2 = True
        Me.Program3 = True
        Me.Program4 = True
        Me.Program5 = True
    Else
        Me.Program7 = False
    End If
    
End Sub

Private Sub Program7_Click()
    
    If Me.Program7 = True Then
        Me.Program0 = True
        Me.Program2 = True
        Me.Program3 = True
        Me.Program4 = True
        Me.Program5 = True
        Me.Program6 = True
    End If
    
End Sub

Private Sub UserForm_Initialize()

    Dim Ctl As Control

    For Each Ctl In Box7_FullReduction.Controls
        If TypeName(Ctl) = "TextBox" Then
            Ctl.Locked = True
        End If
    Next
    Me.CommandButton2.Visible = False
    
    Me.Program0 = CheckBoxProgram0
    Me.Program1 = CheckBoxProgram1
    Me.Program2 = CheckBoxProgram2
    Me.Program3 = CheckBoxProgram3
    Me.Program4 = CheckBoxProgram4
    Me.Program5 = CheckBoxProgram5
    Me.Program6 = CheckBoxProgram6
    Me.Program7 = CheckBoxProgram7

End Sub
