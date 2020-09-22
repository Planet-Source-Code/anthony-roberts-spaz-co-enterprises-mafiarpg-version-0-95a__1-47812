VERSION 5.00
Begin VB.Form frmArmory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Register 
      Caption         =   "Register Gun"
      Height          =   495
      Left            =   2520
      TabIndex        =   15
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Leave"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy This Weapon"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "Sell This Weapon"
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Weapons"
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.OptionButton Option8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Punisher .77mm"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   2880
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Assault Rifle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2520
         Width           =   1695
      End
      Begin VB.OptionButton Option6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Uzi"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Pistol"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Switchblade"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Steel Chains"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Gold Knuckles"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         Caption         =   "Brass Knuckles"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label W4 
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label W3 
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   255
      End
      Begin VB.Label W2 
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label W1 
         BackStyle       =   0  'Transparent
         Caption         =   "W"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   255
      End
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash: $"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   4095
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4935
   End
End
Attribute VB_Name = "frmArmory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuy_Click()
    If Option1.Value = True Then
        If bolBrass = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 100 Then
                    bolBrass = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 100
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 50 Then
                    bolBrass = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 50
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    ElseIf Option2.Value = True Then
        If bolGold = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 500 Then
                    bolGold = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 500
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 250 Then
                    bolGold = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 250
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    ElseIf Option3.Value = True Then
        If bolChain = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 1000 Then
                    bolChain = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 1000
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 500 Then
                    bolChain = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 500
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    ElseIf Option4.Value = True Then
        If bolSwitch = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 2500 Then
                    bolSwitch = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 2500
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 1750 Then
                    bolSwitch = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 1750
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    ElseIf Option5.Value = True Then
        If bolPistol = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 5000 Then
                    bolPistol = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 5000
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 2500 Then
                    bolPistol = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 2500
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    ElseIf Option6.Value = True Then
        If bolUzi = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 15000 Then
                    bolUzi = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 15000
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 7500 Then
                    bolUzi = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 7500
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    ElseIf Option7.Value = True Then
        If bolAssault = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 40000 Then
                    bolAssault = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 40000
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 20000 Then
                    bolAssault = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 20000
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    ElseIf Option8.Value = True Then
        If bolPunisher = False Then
            If strCompany <> "TWM" Then
                If intMoney >= 100000 Then
                    bolPunisher = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 100000
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                If intMoney >= 50000 Then
                    bolPunisher = True
                    lblCash.Caption = "Cash: $" & intMoney
                    intMoney = intMoney - 50000
                Else
                    MsgBox "You don't have enough money!"
                End If
            End If
        Else
            MsgBox "You already own this weapon!"
        End If
    Else
        MsgBox "Select a weapon!"
    End If
    lblCash.Caption = "Cash: $" & intMoney
    If bolBrass = True Then
        Option1.ForeColor = &H80FF&
    End If
    If bolGold = True Then
        Option2.ForeColor = &H80FF&
    End If
    If bolChain = True Then
        Option3.ForeColor = &H80FF&
    End If
    If bolSwitch = True Then
        Option4.ForeColor = &H80FF&
    End If
    If bolPistol = True Then
        Option5.ForeColor = &H80FF&
    End If
    If bolUzi = True Then
        Option6.ForeColor = &H80FF&
    End If
    If bolAssault = True Then
        Option7.ForeColor = &H80FF&
    End If
    If bolPunisher = True Then
        Option8.ForeColor = &H80FF&
    End If
End Sub

Private Sub cmdSell_Click()
    If Option1.Value = True Then
        If bolBrass = True Then
            intMoney = intMoney + 20
            bolBrass = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    ElseIf Option2.Value = True Then
        If bolGold = True Then
            intMoney = intMoney + 100
            bolGold = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    ElseIf Option3.Value = True Then
        If bolChain = True Then
            intMoney = intMoney + 250
            bolChain = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    ElseIf Option4.Value = True Then
        If bolSwitch = True Then
            intMoney = intMoney + 500
            bolSwitch = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    ElseIf Option5.Value = True Then
        If bolPistol = True Then
            intMoney = intMoney + 1000
            bolPistol = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    ElseIf Option6.Value = True Then
        If bolUzi = True Then
            intMoney = intMoney + 2500
            bolUzi = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    ElseIf Option7.Value = True Then
        If bolAssault = True Then
            intMoney = intMoney + 5000
            bolAssault = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    ElseIf Option8.Value = True Then
        If bolPunisher = True Then
            intMoney = intMoney + 10000
            bolPunisher = False
        Else
            MsgBox "You don't own a weapon of this type."
        End If
    Else
        MsgBox "Select a weapon!"
    End If
    lblCash.Caption = "Cash: $" & intMoney
    If bolBrass = False Then
        Option1.ForeColor = &HFFFFFF
    End If
    If bolGold = False Then
        Option2.ForeColor = &HFFFFFF
    End If
    If bolChain = False Then
        Option3.ForeColor = &HFFFFFF
    End If
    If bolSwitch = False Then
        Option4.ForeColor = &HFFFFFF
    End If
    If bolPistol = False Then
        Option5.ForeColor = &HFFFFFF
    End If
    If bolUzi = False Then
        Option6.ForeColor = &HFFFFFF
    End If
    If bolAssault = False Then
        Option7.ForeColor = &HFFFFFF
    End If
    If bolPunisher = False Then
        Option8.ForeColor = &HFFFFFF
    End If
End Sub

Private Sub Command1_Click()
    If bolBrass = False And bolGold = False And bolChain = False And bolSwitch = False And bolPistol = False And bolUzi = False And bolAssault = False And bolPunisher = False Then
        strWeap = "Fists"
        intWeap = 1
        intTrain = 0
    ElseIf bolBrass = True And bolGold = False And bolChain = False And bolSwitch = False And bolPistol = False And bolUzi = False And bolAssault = False And bolPunisher = False Then
        strWeap = "Brass Knuckles"
        intWeap = 2
        intTrain = 0
    ElseIf bolBrass = False And bolGold = True And bolChain = False And bolSwitch = False And bolPistol = False And bolUzi = False And bolAssault = False And bolPunisher = False Then
        strWeap = "Gold Knuckles"
        intWeap = 3
        intTrain = 0
    ElseIf bolBrass = False And bolGold = False And bolChain = True And bolSwitch = False And bolPistol = False And bolUzi = False And bolAssault = False And bolPunisher = False Then
        strWeap = "Chains"
        intWeap = 4
        intTrain = 0
    ElseIf bolBrass = False And bolGold = False And bolChain = False And bolSwitch = True And bolPistol = False And bolUzi = False And bolAssault = False And bolPunisher = False Then
        strWeap = "Switchblade"
        intWeap = 5
        intTrain = 0
    ElseIf bolBrass = False And bolGold = False And bolChain = False And bolSwitch = False And bolPistol = True And bolUzi = False And bolAssault = False And bolPunisher = False Then
        strWeap = "Pistol"
        intWeap = 6
        intTrain = 0
    ElseIf bolBrass = False And bolGold = False And bolChain = False And bolSwitch = False And bolPistol = False And bolUzi = True And bolAssault = False And bolPunisher = False Then
        strWeap = "Uzi"
        intWeap = 7
        intTrain = 0
    ElseIf bolBrass = False And bolGold = False And bolChain = False And bolSwitch = False And bolPistol = False And bolUzi = False And bolAssault = True And bolPunisher = False Then
        strWeap = "Assault Rifle"
        intWeap = 8
        intTrain = 0
    ElseIf bolBrass = False And bolGold = False And bolChain = False And bolSwitch = False And bolPistol = False And bolUzi = False And bolAssault = False And bolPunisher = True Then
        strWeap = "Punisher .77mm"
        intWeap = 10
        intTrain = 0
    End If
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    frmRPG.lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    Unload frmArmory
End Sub

Private Sub Form_Load()
    lblCash.Caption = "Cash: $" & intMoney
    If bolBrass = True Then
        Option1.ForeColor = &H80FF&
    End If
    If bolGold = True Then
        Option2.ForeColor = &H80FF&
    End If
    If bolChain = True Then
        Option3.ForeColor = &H80FF&
    End If
    If bolSwitch = True Then
        Option4.ForeColor = &H80FF&
    End If
    If bolPistol = True Then
        Option5.ForeColor = &H80FF&
    End If
    If bolUzi = True Then
        Option6.ForeColor = &H80FF&
    End If
    If bolAssault = True Then
        Option7.ForeColor = &H80FF&
    End If
    If bolPunisher = True Then
        Option8.ForeColor = &H80FF&
    End If
    
    If bolW1 = True Then
        W1.ForeColor = &HFFFF&
    End If
    If bolW2 = True Then
        W2.ForeColor = &HFFFF&
    End If
    If bolW3 = True Then
        W3.ForeColor = &HFFFF&
    End If
    If bolW4 = True Then
        W4.ForeColor = &HFFFF&
    End If
    Register.Caption = "Register Weapon" & vbLf & "$10,000"
    Register.Enabled = False
End Sub

Private Sub Option1_Click()
    lblName.Caption = "Brass Knuckles"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 2" & vbLf & "Buy Cost: $100" & vbLf & "Sell Cost: $20"
    Else
        lblStats.Caption = "Attack: 2" & vbLf & "Buy Cost: $50" & vbLf & "Sell Cost: $20"
    End If
    Register.Enabled = False
End Sub

Private Sub Option2_Click()
    lblName.Caption = "Gold Knuckles"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 3" & vbLf & "Buy Cost: $500" & vbLf & "Sell Cost: $100"
    Else
        lblStats.Caption = "Attack: 3" & vbLf & "Buy Cost: $250" & vbLf & "Sell Cost: $100"
    End If
    Register.Enabled = False
End Sub

Private Sub Option3_Click()
    lblName.Caption = "Steel Chains"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 4" & vbLf & "Buy Cost: $1,000" & vbLf & "Sell Cost: $250"
    Else
        lblStats.Caption = "Attack: 4" & vbLf & "Buy Cost: $500" & vbLf & "Sell Cost: $250"
    End If
    Register.Enabled = False
End Sub

Private Sub Option4_Click()
    lblName.Caption = "Switchblade"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 5" & vbLf & "Buy Cost: $2,500" & vbLf & "Sell Cost: $500"
    Else
        lblStats.Caption = "Attack: 5" & vbLf & "Buy Cost: $1,750" & vbLf & "Sell Cost: $500"
    End If
    Register.Enabled = False
End Sub

Private Sub Option5_Click()
    lblName.Caption = "Pistol"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 6" & vbLf & "Buy Cost: $5,000" & vbLf & "Sell Cost: $1,000"
    Else
        lblStats.Caption = "Attack: 6" & vbLf & "Buy Cost: $2,500" & vbLf & "Sell Cost: $1,000"
    End If
    
    If strCompany <> "TWM" Then
        If bolW1 = False Then
            Register.Enabled = True
        Else
            Register.Enabled = False
        End If
    End If
End Sub

Private Sub Option6_Click()
    lblName.Caption = "Uzi"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 7" & vbLf & "Buy Cost: $15,000" & vbLf & "Sell Cost: $2,500"
    Else
        lblStats.Caption = "Attack: 7" & vbLf & "Buy Cost: $7,500" & vbLf & "Sell Cost: $2,500"
    End If
    
    If strCompany <> "TWM" Then
        If bolW2 = False Then
            Register.Enabled = True
        Else
            Register.Enabled = False
        End If
    End If
End Sub

Private Sub Option7_Click()
    lblName.Caption = "Assault Rifle"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 8" & vbLf & "Buy Cost: $40,000" & vbLf & "Sell Cost: $5,000"
    Else
        lblStats.Caption = "Attack: 8" & vbLf & "Buy Cost: $20,000" & vbLf & "Sell Cost: $5,000"
    End If
    
    If strCompany <> "TWM" Then
        If bolW3 = False Then
            Register.Enabled = True
        Else
            Register.Enabled = False
        End If
    End If
End Sub

Private Sub Option8_Click()
    lblName.Caption = "Punisher .77mm"
    If strCompany <> "TWM" Then
        lblStats.Caption = "Attack: 10" & vbLf & "Buy Cost: $100,000" & vbLf & "Sell Cost: $10,000"
    Else
        lblStats.Caption = "Attack: 10" & vbLf & "Buy Cost: $50,000" & vbLf & "Sell Cost: $10,000"
    End If
    
    If strCompany <> "TWM" Then
        If bolW4 = False Then
            Register.Enabled = True
        Else
            Register.Enabled = False
        End If
    End If
End Sub

Private Sub Register_Click()
    If Option5.Value = True Then
        If intMoney >= 10000 Then
            bolW1 = True
            MsgBox "You purchase a warrant for this weapon"
            W1.ForeColor = &HFFFF&
            Register.Enabled = False
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf Option6.Value = True Then
        If intMoney >= 10000 Then
            bolW2 = True
            MsgBox "You purchase a warrant for this weapon"
            W2.ForeColor = &HFFFF&
            Register.Enabled = False
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf Option7.Value = True Then
        If intMoney >= 10000 Then
            bolW3 = True
            MsgBox "You purchase a warrant for this weapon"
            W3.ForeColor = &HFFFF&
            Register.Enabled = False
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf Option8.Value = True Then
        If intMoney >= 10000 Then
            bolW4 = True
            MsgBox "You purchase a warrant for this weapon"
            W4.ForeColor = &HFFFF&
            Register.Enabled = False
        Else
            MsgBox "You don't have enough money!"
        End If
    Else
        MsgBox "Make a selection!"
    End If
End Sub

