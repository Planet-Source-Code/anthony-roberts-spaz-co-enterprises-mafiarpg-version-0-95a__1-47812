VERSION 5.00
Begin VB.Form frmLevel 
   BorderStyle     =   0  'None
   Caption         =   "Level up!"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Weapon"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   3855
   End
   Begin VB.CommandButton cmdHealPotion 
      Caption         =   "Buy more Heal Potions (5 Skill Credits for 3)"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   960
      Width           =   4095
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Return to Game"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   4095
   End
   Begin VB.Frame fraWeapon 
      Caption         =   "New Weapon"
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   4095
      Begin VB.OptionButton opt7 
         Caption         =   "Firesword (250 Skill Points) (Attack: 10)"
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   3855
      End
      Begin VB.OptionButton opt6 
         Caption         =   "Scimitar (100 Skill Points) (Attack: 8)"
         ForeColor       =   &H00008080&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   3855
      End
      Begin VB.OptionButton opt5 
         Caption         =   "Sword (50 Skill Points) (Attack: 7)"
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   3855
      End
      Begin VB.OptionButton opt4 
         Caption         =   "Machete (30 Skill Points) (Attack: 6)"
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   3855
      End
      Begin VB.OptionButton opt3 
         Caption         =   "Butcher Knife (20 Skill Points) (Attack: 5)"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3855
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Steak Knife (10 Skill Points) (Attack: 4)"
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3855
      End
      Begin VB.OptionButton opt1 
         Caption         =   "Butter Knife (5 Skill Points) (Attack: 3)"
         ForeColor       =   &H00000040&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdHP 
      Caption         =   "Increse HP by 15 (5 Skill Credits)"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "Set your skill points."
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape shpBack 
      BackColor       =   &H80000004&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008080&
      BorderWidth     =   5
      Height          =   4335
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHealPotion_Click()
    If intSet >= 5 Then
        intHeal = intHeal + 3
        intSet = intSet - 5
    Else
        MsgBox "You've not enough points to get more Heal Potions."
    End If
    lblMessage.Caption = "You have " & intSet & " skill points to spend on Hit Points or a new Weapon."
End Sub

Private Sub cmdHP_Click()
    If intSet >= intHPIncrese Then
        intMaxHP = intMaxHP + 15
        intHP = intHP + 15
        intSet = intSet - intHPIncrese
        intHPIncrese = intHPIncrese + 5
    Else
        MsgBox "You've not enough Skill Points to spend on more Hit Points."
    End If
    lblMessage.Caption = "You have " & intSet & " skill points to spend on Hit Points or a new Weapon."
    cmdHP.Caption = "Increse HP by 15 (" & intHPIncrese & " Skill Credits)"
End Sub

Private Sub cmdReturn_Click()
    frmRPG.lblName.Caption = strName & ", Level " & intLevel
    frmRPG.lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    frmRPG.lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
    frmRPG.cmdFlee.Caption = "Use Heal (" & intHeal & ")"
    frmRPG.Visible = True
    Unload frmLevel
End Sub

Private Sub cmdSelect_Click()
    If opt1.Value = True And intSet >= 5 Then
        opt1.Enabled = False
        opt2.Enabled = True
        opt3.Enabled = True
        opt4.Enabled = True
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
        strWeap = "Butter Knife"
        intWeap = 3
        intSet = intSet - 5
        intTrain = 0
    ElseIf opt2.Value = True And intSet >= 10 Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = True
        opt4.Enabled = True
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
        strWeap = "Steak Knife"
        intWeap = 4
        intSet = intSet - 10
        intTrain = 0
    ElseIf opt3.Value = True And intSet >= 20 Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = True
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
        strWeap = "Butcher Knife"
        intWeap = 5
        intSet = intSet - 20
        intTrain = 0
    ElseIf opt4.Value = True And intSet >= 30 Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
        strWeap = "Machete"
        intWeap = 6
        intSet = intSet - 30
        intTrain = 0
    ElseIf opt5.Value = True And intSet >= 50 Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = False
        opt6.Enabled = True
        opt7.Enabled = True
        strWeap = "Sword"
        intWeap = 7
        intSet = intSet - 50
        intTrain = 0
    ElseIf opt6.Value = True And intSet >= 100 Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = False
        opt6.Enabled = False
        opt7.Enabled = True
        strWeap = "Scimitar"
        intWeap = 8
        intSet = intSet - 100
        intTrain = 0
    ElseIf opt7.Value = True And intSet >= 250 Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = False
        opt6.Enabled = False
        opt7.Enabled = False
        strWeap = "Firesword"
        intWeap = 10
        intSet = intSet - 250
        intTrain = 0
    Else
        MsgBox "No selection was made!"
    End If
    lblMessage.Caption = "You have " & intSet & " skill points to spend on Hit Points or a new Weapon."
    frmRPG.lblPercent = (intTrain * 4) & "%"
    frmRPG.shpTrain.Height = intTrain * 4
End Sub

Private Sub Form_Load()
    cmdHP.Caption = "Increse HP by 15 (" & intHPIncrese & " Skill Credits)"
    lblMessage.Caption = "You have " & intSet & " skill points to spend on Hit Points or a new Weapon."
    If strWeap = "Sharp Stick" Then
        opt1.Enabled = True
        opt2.Enabled = True
        opt3.Enabled = True
        opt4.Enabled = True
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
    ElseIf strWeap = "Butter Knife" Then
        opt1.Enabled = False
        opt2.Enabled = True
        opt3.Enabled = True
        opt4.Enabled = True
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
    ElseIf strWeap = "Steak Knife" Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = True
        opt4.Enabled = True
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
    ElseIf strWeap = "Butcher Knife" Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = True
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
    ElseIf strWeap = "Machete" Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = True
        opt6.Enabled = True
        opt7.Enabled = True
    ElseIf strWeap = "Sword" Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = False
        opt6.Enabled = True
        opt7.Enabled = True
    ElseIf strWeap = "Scimitar" Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = False
        opt6.Enabled = False
        opt7.Enabled = True
    ElseIf strWeap = "Firesword" Then
        opt1.Enabled = False
        opt2.Enabled = False
        opt3.Enabled = False
        opt4.Enabled = False
        opt5.Enabled = False
        opt6.Enabled = False
        opt7.Enabled = False
    End If
End Sub
