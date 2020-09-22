VERSION 5.00
Begin VB.Form frmCasino 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4335
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Win Values"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Spin"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   2895
   End
   Begin VB.OptionButton Option6 
      Caption         =   "$100000"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   3120
      Width           =   975
   End
   Begin VB.OptionButton Option5 
      Caption         =   "$10000"
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.OptionButton Option4 
      Caption         =   "$1000"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.OptionButton Option3 
      Caption         =   "$100"
      Height          =   195
      Left            =   2040
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.OptionButton Option2 
      Caption         =   "$10"
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "$1"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Leave"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Welcome!"
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   4095
   End
   Begin VB.Label lblNum 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   480
      TabIndex        =   7
      Top             =   120
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmCasino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    Unload frmCasino
End Sub

Private Sub Command2_Click()
    If Option1.Value = True Then
        intBet = 1
    ElseIf Option2.Value = True Then
        intBet = 10
    ElseIf Option3.Value = True Then
        intBet = 100
    ElseIf Option4.Value = True Then
        intBet = 1000
    ElseIf Option5.Value = True Then
        intBet = 10000
    ElseIf Option6.Value = True Then
        intBet = 100000
    Else
        MsgBox "Please make a bet"
        intBet = 0
    End If
    If intMoney < intBet Then
        MsgBox "You've not enough money to make this bet"
        Exit Sub
    End If
    If intBet <> 0 Then
        intNum1 = Rnd * 8 + 1
        intNum2 = Rnd * 8 + 1
    End If
    lblNum.Caption = intNum1 & intNum2
    If intNum1 = intNum2 Then
        If intNum1 = 1 Then
            intMoney = intMoney + (intBet * 2)
            lblMessage.Caption = "You win $" & (intBet * 2)
        ElseIf intNum1 = 2 Then
            intMoney = intMoney + (intBet * 3)
            lblMessage.Caption = "You win $" & (intBet * 3)
        ElseIf intNum1 = 3 Then
            intMoney = intMoney + (intBet * 5)
            lblMessage.Caption = "You win $" & (intBet * 5)
        ElseIf intNum1 = 4 Then
            intMoney = intMoney + (intBet * 10)
            lblMessage.Caption = "You win $" & (intBet * 10)
        ElseIf intNum1 = 5 Then
            intMoney = intMoney + (intBet * 25)
            lblMessage.Caption = "You win $" & (intBet * 25)
        ElseIf intNum1 = 6 Then
            intMoney = intMoney - intBet
            lblMessage.Caption = "You lose $" & intBet
        ElseIf intNum1 = 7 Then
            intMoney = intMoney + (intBet * 100)
            lblMessage.Caption = "You win $" & (intBet * 100)
        ElseIf intNum1 = 8 Then
            intMoney = intMoney + (intBet * 10)
            lblMessage.Caption = "You win $" & (intBet * 10)
        ElseIf intNum1 = 9 Then
            intMoney = intMoney + (intBet * 5)
            lblMessage.Caption = "You win $" & (intBet * 5)
        End If
    ElseIf intNum1 = 1 And intNum2 = 2 Then
        intMoney = intMoney + (intBet * 2)
        lblMessage.Caption = "You win $" & (intBet * 2)
    Else
        intMoney = intMoney - intBet
        lblMessage.Caption = "You lose $" & intBet
    End If
End Sub

Private Sub Command4_Click()
    MsgBox "11 = Bet (x) 2" & vbLf & _
            "22 = Bet (x) 3" & vbLf & _
            "33 = Bet (x) 5" & vbLf & _
            "44 = Bet (x) 10" & vbLf & _
            "55 = Bet (x) 25" & vbLf & _
            "66 = Lose your bet" & vbLf & _
            "77 = Bet (x) 100" & vbLf & _
            "88 = Bet (x) 10" & vbLf & _
            "99 = Bet (x) 5" & vbLf & _
            "12 = Bet (x) 2", vbOKOnly, "Win List"
End Sub

Private Sub Form_Load()
    Randomize
End Sub
