VERSION 5.00
Begin VB.Form frmArena 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBet 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdBet 
      Caption         =   "Make Bet"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000080&
      Caption         =   "Enemy"
      ForeColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdChallenge 
         Caption         =   "Challenge Someone"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label lblEnHP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblEnWeap 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label lblEnName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdFlee 
         Caption         =   "Flee"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   2175
      End
      Begin VB.CommandButton cmdAttack 
         Caption         =   "Attack"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label lblHP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblWeapon 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Leave"
      Height          =   855
      Left            =   4320
      TabIndex        =   0
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Underground Arena"
      ForeColor       =   &H0000C0C0&
      Height          =   855
      Left            =   1320
      TabIndex        =   8
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   3375
      Left            =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmArena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAttack_Click()
    intHit = (Rnd * intWeap) + intSpec
    If intHit > 0 Then
        intArenaHP = intArenaHP - intHit
        strMessage = "You hit " & strArenaName & " for " & intHit & " damage!"
    Else
        strMessage = "You miss the " & strArenaName
    End If
    
    If intArenaHP < 1 Then
        strMessage = "You KO " & strArenaName & "!"
        intArenaEarn = intArenaOdds * intArenaBet
        strMessage2 = "You earn $" & intArenaEarn
        lblMessage.ForeColor = &HC000&
        intMoney = intMoney + intArenaEarn
        cmdBet.Enabled = False
        cmdAttack.Enabled = False
        cmdFlee.Enabled = False
        cmdQuit.Enabled = True
        cmdChallenge.Enabled = False
    End If
    
    If intArenaHP > 0 Then
        intHit = Rnd * intArenaWeapon
        If intHit > 0 Then
            intHP = intHP - intHit
            strMessage2 = "You are hit for " & intHit
        Else
            strMessage2 = "You are missed by the " & strArenaName
        End If
        If intHP < 1 Then
            strMessage = "You are KO'ed!"
            intMoney = intMoney - intArenaBet
            intHP = 1
            cmdBet.Enabled = False
            cmdAttack.Enabled = False
            cmdFlee.Enabled = False
            cmdQuit.Enabled = True
            cmdChallenge.Enabled = False
        End If
    End If
    lblMessage.Caption = strMessage & vbLf & strMessage2
    lblEnHP.Caption = "HP: " & intArenaHP & " / " & intArenaMaxHP
    lblHP.Caption = "HP: " & intHP & " / " & intMaxHP
End Sub

Private Sub cmdBet_Click()
    If txtBet.Text = "" Then
        txtBet.Text = 0
    End If
    If intMoney >= txtBet.Text Then
        cmdChallenge.Enabled = True
        intArenaBet = txtBet.Text
        cmdBet.Enabled = False
    Else
        MsgBox "You've not enough money for this bet!"
    End If
End Sub

Private Sub cmdChallenge_Click()
    intChance = Rnd * 25
    If intChance = 1 Then
        strArenaName = "Buck"
    ElseIf intChance = 2 Then
        strArenaName = "Joe"
    ElseIf intChance = 3 Then
        strArenaName = "Larry"
    ElseIf intChance = 4 Then
        strArenaName = "Smick"
    ElseIf intChance = 5 Then
        strArenaName = "Jakotai"
    ElseIf intChance = 6 Then
        strArenaName = "Sissy"
    ElseIf intChance = 7 Then
        strArenaName = "Mack"
    ElseIf intChance = 8 Then
        strArenaName = "Cass"
    ElseIf intChance = 9 Then
        strArenaName = "Tinkerbell"
    ElseIf intChance = 10 Then
        strArenaName = "Raisin"
    ElseIf intChance = 11 Then
        strArenaName = "Damien"
    ElseIf intChance = 12 Then
        strArenaName = "Carby"
    ElseIf intChance = 13 Then
        strArenaName = "Mike"
    ElseIf intChance = 14 Then
        strArenaName = "Micheal"
    ElseIf intChance = 15 Then
        strArenaName = "Spaz"
    ElseIf intChance = 16 Then
        strArenaName = "Al"
    ElseIf intChance = 17 Then
        strArenaName = "Ali"
    ElseIf intChance = 18 Then
        strArenaName = "King"
    ElseIf intChance = 19 Then
        strArenaName = "Elvis"
    ElseIf intChance = 20 Then
        strArenaName = "Chimmy"
    ElseIf intChance = 21 Then
        strArenaName = "Chongo"
    ElseIf intChance = 22 Then
        strArenaName = "Garbonzo"
    ElseIf intChance = 23 Then
        strArenaName = "Mary"
    ElseIf intChance = 24 Then
        strArenaName = "Siskel"
    Else
        strArenaName = "Unknown"
    End If
    
    intChance = Rnd * 10
    If intChance = 1 Then
        strArenaWeapon = "Fingernails"
    ElseIf intChance = 2 Then
        strArenaWeapon = "Chains"
    ElseIf intChance = 3 Then
        strArenaWeapon = "Steel Knuckles"
    ElseIf intChance = 4 Then
        strArenaWeapon = "Magnum 9mm"
    ElseIf intChance = 5 Then
        strArenaWeapon = "Soap Bar"
    ElseIf intChance = 6 Then
        strArenaWeapon = "Sign Post"
    ElseIf intChance = 7 Then
        strArenaWeapon = "Fold-Up Chair"
    ElseIf intChance = 8 Then
        strArenaWeapon = "Fold-Up Table"
    ElseIf intChance = 9 Then
        strArenaWeapon = "Grabage Can Lid"
    Else
        strArenaWeapon = "Fists"
    End If
    
    intArenaWeapon = (Int(Rnd * 25 + 1)) + intLevel
    intArenaMaxHP = (Int(Rnd * 1000 + 1)) + intLevel
    intArenaHP = intArenaMaxHP
    intArenaOdds = (Int(intArenaWeapon / 2)) + (Int(intArenaMaxHP / 50)) - intLevel
    If intArenaOdds <= 0 Then
        intArenaOdds = 1
    End If
    
    lblMessage.Caption = "$" & intArenaBet & " bet on odds " & intArenaOdds & ":1" & vbLf & "Fighting: " & strArenaName
    lblEnName.Caption = strArenaName
    lblEnWeap.Caption = strArenaWeapon & " (" & intArenaWeapon & ")"
    lblEnHP.Caption = "HP: " & intArenaHP & " / " & intArenaMaxHP
    lblWeapon.Caption = strWeap & " (" & intWeap & "+" & intSpec & ")"
    lblHP.Caption = "HP: " & intHP & " / " & intMaxHP
    
    cmdAttack.Enabled = True
    cmdFlee.Enabled = True
    cmdQuit.Enabled = False
    cmdChallenge.Enabled = False
End Sub

Private Sub cmdFlee_Click()
    intMoney = intMoney - intArenaBet
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    frmRPG.lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    Unload frmArena
End Sub

Private Sub cmdQuit_Click()
    frmRPG.Enabled = True
    frmRPG.Visible = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    frmRPG.lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    Unload frmArena
End Sub

Private Sub Form_Load()
    Randomize
    lblWeapon.Caption = strWeap & " (" & intWeap & "+" & intSpec & ")"
    lblHP.Caption = "HP: " & intHP & " / " & intMaxHP
    Frame1.Caption = strName
End Sub
