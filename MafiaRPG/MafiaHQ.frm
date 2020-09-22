VERSION 5.00
Begin VB.Form frmHQ 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Leave"
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Top             =   4680
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Heist"
      Height          =   1455
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   4815
      Begin VB.OptionButton Option14 
         Caption         =   "Rival Company"
         Height          =   195
         Left            =   2280
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Local Bank"
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option12 
         Caption         =   "School"
         Height          =   195
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Gas Station"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option10 
         Caption         =   "House"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Old Lady"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblSteal 
         Alignment       =   2  'Center
         Caption         =   "No Information Available"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1080
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Companies"
      Height          =   1695
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   4815
      Begin VB.OptionButton Option8 
         Caption         =   "Blazing Auto's"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option7 
         Caption         =   "The White Market"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Clit N' Sons"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Hutch In Corp."
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Life After Death Co."
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Bones Inc."
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Shadow Hentai Enterprises"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Spaz Co. Enterprises"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblEarned 
         Alignment       =   2  'Center
         Caption         =   "No Information Available"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Withdraw Money from Selected Business"
      Height          =   495
      Left            =   660
      TabIndex        =   1
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton cmdHeist 
      Caption         =   "Prepare a Heist"
      Height          =   495
      Left            =   660
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   5295
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderWidth     =   10
      Height          =   5295
      Left            =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmHQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHeist_Click()
    If Option9.Value = True Then
        If intSpec >= 2 Then
            intChance = Rnd * 100
            intChance = intChance + intSpeed
            If intChance >= 1 And intChance <= 10 Then
                MsgBox "You were caught by the police!  You lose all weapons (Excluding Permitted Weapons), money, and drugs and start this level again."
                bolBrass = False
                strWeap = "Fists"
                intWeap = 1
                bolGold = False
                bolChain = False
                bolSwitch = False
                If bolW1 <> True Then
                    bolPistol = False
                End If
                If bolW2 <> True Then
                    bolUzi = False
                End If
                If bolW3 <> True Then
                    bolAssault = False
                End If
                If bolW4 <> True Then
                    bolPunisher = False
                End If
                intSpoon = 0
                intLazy = 0
                intKoocie = 0
                intLucky = 0
                intHair = 0
                intCoka = 0
                intCalf = 0
                intMoney = 0
                intAmmo = 0
                intExp = intCurLevel / 2
                intCurLevel = intCurLevel / 2
            Else
                intChance = Int(Rnd * 50)
                intMoney = intMoney + intChance
                lblSteal.Caption = "You steal $" & intChance
            End If
        Else
            MsgBox "You don't have enough Special Training to make this heist!"
        End If
    ElseIf Option10.Value = True Then
        If intSpec >= 4 Then
            intChance = Rnd * 100
            intChance = intChance + intSpeed
            If intChance >= 1 And intChance <= 15 Then
                MsgBox "You were caught by the police!  You lose all weapons, money, and drugs and start this level again."
                bolBrass = False
                strWeap = "Fists"
                intWeap = 1
                bolGold = False
                bolChain = False
                bolSwitch = False
                If bolW1 <> True Then
                    bolPistol = False
                End If
                If bolW2 <> True Then
                    bolUzi = False
                End If
                If bolW3 <> True Then
                    bolAssault = False
                End If
                If bolW4 <> True Then
                    bolPunisher = False
                End If
                intSpoon = 0
                intLazy = 0
                intKoocie = 0
                intLucky = 0
                intHair = 0
                intCoka = 0
                intCalf = 0
                intMoney = 0
                intAmmo = 0
                intExp = intCurLevel / 2
                intCurLevel = intCurLevel / 2
            Else
                intChance = Int(Rnd * 75)
                intMoney = intMoney + intChance
                lblSteal.Caption = "You steal $" & intChance
            End If
        Else
            MsgBox "You don't have enough Special Training to make this heist!"
        End If
    ElseIf Option11.Value = True Then
        If intSpec >= 10 Then
            intChance = Rnd * 100
            intChance = intChance + intSpeed
            If intChance >= 1 And intChance <= 20 Then
                MsgBox "You were caught by the police!  You lose all weapons, money, and drugs and start this level again."
                bolBrass = False
                strWeap = "Fists"
                intWeap = 1
                bolGold = False
                bolChain = False
                bolSwitch = False
                If bolW1 <> True Then
                    bolPistol = False
                End If
                If bolW2 <> True Then
                    bolUzi = False
                End If
                If bolW3 <> True Then
                    bolAssault = False
                End If
                If bolW4 <> True Then
                    bolPunisher = False
                End If
                intSpoon = 0
                intLazy = 0
                intKoocie = 0
                intLucky = 0
                intHair = 0
                intCoka = 0
                intCalf = 0
                intMoney = 0
                intAmmo = 0
                intExp = intCurLevel / 2
                intCurLevel = intCurLevel / 2
            Else
                intChance = Int(Rnd * 1000)
                intMoney = intMoney + intChance
                lblSteal.Caption = "You steal $" & intChance
            End If
        Else
            MsgBox "You don't have enough Special Training to make this heist!"
        End If
    ElseIf Option12.Value = True Then
        If intSpec >= 25 Then
            intChance = Rnd * 100
            intChance = intChance + intSpeed
            If intChance >= 1 And intChance <= 25 Then
                MsgBox "You were caught by the police!  You lose all weapons, money, and drugs and start this level again."
                bolBrass = False
                strWeap = "Fists"
                intWeap = 1
                bolGold = False
                bolChain = False
                bolSwitch = False
                If bolW1 <> True Then
                    bolPistol = False
                End If
                If bolW2 <> True Then
                    bolUzi = False
                End If
                If bolW3 <> True Then
                    bolAssault = False
                End If
                If bolW4 <> True Then
                    bolPunisher = False
                End If
                intSpoon = 0
                intLazy = 0
                intKoocie = 0
                intLucky = 0
                intHair = 0
                intCoka = 0
                intCalf = 0
                intMoney = 0
                intAmmo = 0
                intExp = intCurLevel / 2
                intCurLevel = intCurLevel / 2
            Else
                intChance = Int(Rnd * 10000)
                intMoney = intMoney + intChance
                lblSteal.Caption = "You steal $" & intChance
            End If
        Else
            MsgBox "You don't have enough Special Training to make this heist!"
        End If
    ElseIf Option13.Value = True Then
        If intSpec >= 50 Then
            intChance = Rnd * 100
            intChance = intChance + intSpeed
            If intChance >= 1 And intChance <= 75 Then
                MsgBox "You were caught by the police!  You lose all weapons, money, and drugs and start this level again."
                bolBrass = False
                strWeap = "Fists"
                intWeap = 1
                bolGold = False
                bolChain = False
                bolSwitch = False
                If bolW1 <> True Then
                    bolPistol = False
                End If
                If bolW2 <> True Then
                    bolUzi = False
                End If
                If bolW3 <> True Then
                    bolAssault = False
                End If
                If bolW4 <> True Then
                    bolPunisher = False
                End If
                intSpoon = 0
                intLazy = 0
                intKoocie = 0
                intLucky = 0
                intHair = 0
                intCoka = 0
                intCalf = 0
                intMoney = 0
                intAmmo = 0
                intExp = intCurLevel / 2
                intCurLevel = intCurLevel / 2
            Else
                intChance = Int(Rnd * 500000)
                intMoney = intMoney + intChance
                lblSteal.Caption = "You steal $" & intChance
            End If
        Else
            MsgBox "You don't have enough Special Training to make this heist!"
        End If
    ElseIf Option14.Value = True Then
        If intSpec >= 75 Then
            intChance = Rnd * 100
            intChance = intChance + intSpeed
            If intChance >= 1 And intChance <= 90 Then
                MsgBox "You were caught by the police!  You lose all weapons, money, and drugs and start this level again."
                bolBrass = False
                strWeap = "Fists"
                intWeap = 1
                bolGold = False
                bolChain = False
                bolSwitch = False
                If bolW1 <> True Then
                    bolPistol = False
                End If
                If bolW2 <> True Then
                    bolUzi = False
                End If
                If bolW3 <> True Then
                    bolAssault = False
                End If
                If bolW4 <> True Then
                    bolPunisher = False
                End If
                intSpoon = 0
                intLazy = 0
                intKoocie = 0
                intLucky = 0
                intHair = 0
                intCoka = 0
                intCalf = 0
                intMoney = 0
                intAmmo = 0
                intExp = intCurLevel / 2
                intCurLevel = intCurLevel / 2
            Else
                intChance = Int(Rnd * 10000000)
                intMoney = intMoney + intChance
                lblSteal.Caption = "You steal $" & intChance
            End If
        Else
            MsgBox "You don't have enough Special Training to make this heist!"
        End If
    Else
        MsgBox "Make a selection!"
    End If
End Sub

Private Sub Command1_Click()
    If Option1.Value = True Then
        If Int(((intSCE / totSCE) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl1
            intMoney = intMoney + dbl1
            dbl1 = 0
            lblEarned.Caption = "Earned: $" & dbl1
        Else
            MsgBox "You don't completly own this company!"
        End If
    ElseIf Option2.Value = True Then
        If Int(((intSHE / totSHE) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl2
            intMoney = intMoney + dbl2
            dbl2 = 0
            lblEarned.Caption = "Earned: $" & dbl2
        Else
            MsgBox "You don't completly own this company!"
        End If
    ElseIf Option3.Value = True Then
        If Int(((intBI / totBI) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl3
            intMoney = intMoney + dbl3
            dbl3 = 0
            lblEarned.Caption = "Earned: $" & dbl3
        Else
            MsgBox "You don't completly own this company!"
        End If
    ElseIf Option4.Value = True Then
        If Int(((intLADC / totLADC) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl4
            intMoney = intMoney + dbl4
            dbl4 = 0
            lblEarned.Caption = "Earned: $" & dbl4
        Else
            MsgBox "You don't completly own this company!"
        End If
    ElseIf Option5.Value = True Then
        If Int(((intHC / totHC) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl5
            intMoney = intMoney + dbl5
            dbl5 = 0
            lblEarned.Caption = "Earned: $" & dbl5
        Else
            MsgBox "You don't completly own this company!"
        End If
    ElseIf Option6.Value = True Then
        If Int(((intCNS / totCNS) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl6
            intMoney = intMoney + dbl6
            dbl6 = 0
            lblEarned.Caption = "Earned: $" & dbl6
        Else
            MsgBox "You don't completly own this company!"
        End If
    ElseIf Option7.Value = True Then
        If Int(((intTWM / totTWM) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl7
            intMoney = intMoney + dbl7
            dbl7 = 0
            lblEarned.Caption = "Earned: $" & dbl7
        Else
            MsgBox "You don't completly own this company!"
        End If
    ElseIf Option8.Value = True Then
        If Int(((intBA / totBA) * 100)) = 100 Then
            MsgBox "You withdraw $" & dbl8
            intMoney = intMoney + dbl8
            dbl8 = 0
            lblEarned.Caption = "Earned: $" & dbl8
        Else
            MsgBox "You don't completly own this company!"
        End If
    Else
        MsgBox "Make a selection!"
    End If
End Sub

Private Sub Command2_Click()
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    Unload frmHQ
End Sub

Private Sub Option1_Click()
    lblEarned.Caption = "Earned: $" & dbl1
End Sub

Private Sub Option10_Click()
    lblSteal.Caption = "You require [4] Special Attack for this Heist"
End Sub

Private Sub Option11_Click()
    lblSteal.Caption = "You require [10] Special Attack for this Heist"
End Sub

Private Sub Option12_Click()
    lblSteal.Caption = "You require [25] Special Attack for this Heist"
End Sub

Private Sub Option13_Click()
    lblSteal.Caption = "You require [50] Special Attack for this Heist"
End Sub

Private Sub Option14_Click()
    lblSteal.Caption = "You require [75] Special Attack for this Heist"
End Sub

Private Sub Option2_Click()
    lblEarned.Caption = "Earned: $" & dbl2
End Sub

Private Sub Option3_Click()
    lblEarned.Caption = "Earned: $" & dbl3
End Sub

Private Sub Option4_Click()
    lblEarned.Caption = "Earned: $" & dbl4
End Sub

Private Sub Option5_Click()
    lblEarned.Caption = "Earned: $" & dbl5
End Sub

Private Sub Option6_Click()
    lblEarned.Caption = "Earned: $" & dbl6
End Sub

Private Sub Option7_Click()
    lblEarned.Caption = "Earned: $" & dbl7
End Sub

Private Sub Option8_Click()
    lblEarned.Caption = "Earned: $" & dbl8
End Sub

Private Sub Option9_Click()
    lblSteal.Caption = "You require [2] Special Attack for this Heist"
End Sub
