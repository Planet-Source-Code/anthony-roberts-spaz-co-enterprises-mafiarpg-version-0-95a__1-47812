VERSION 5.00
Begin VB.Form frmTattoo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Open a New Account - $100"
      Height          =   735
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   4215
   End
   Begin VB.TextBox txtMoney 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Withdraw"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Deposit"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Leave"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblBankCash 
      Alignment       =   2  'Center
      Caption         =   "Money Deposited: $0"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Current Bank's Monthly Interest Rate: 5%"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      Caption         =   "Cash: $"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ontrai International Bank"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   10
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmTattoo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    Unload frmTattoo
End Sub

Private Sub Command2_Click()
    If intMoney >= txtMoney.Text Then
        intBank = intBank + txtMoney.Text
        intMoney = intMoney - txtMoney.Text
    Else
        MsgBox "You don't have enough money for this deposit!"
    End If
    lblBankCash.Caption = "Money Deposited: $" & intBank
    lblCash.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command3_Click()
    If intBank >= txtMoney.Text Then
        intBank = intBank - txtMoney.Text
        intMoney = intMoney + txtMoney.Text
    Else
        MsgBox "You don't have this much money in the bank!"
    End If
    lblBankCash.Caption = "Money Deposited: $" & intBank
    lblCash.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command4_Click()
    If bolBank = False Then
        If intMoney >= 100 Then
            bolBank = True
            MsgBox "You're account is open!"
            Command2.Enabled = True
            Command3.Enabled = True
        Else
            MsgBox "You don't have enough money!"
        End If
    End If
    lblCash.Caption = "Cash: $" & intMoney
End Sub

Private Sub Form_Load()
    lblCash.Caption = "Cash: $" & intMoney
    lblBankCash.Caption = "Money Deposited: $" & intBank
    If bolBank = True Then
        Command4.Enabled = False
        Command2.Enabled = True
        Command3.Enabled = True
    Else
        Command4.Enabled = True
        Command2.Enabled = False
        Command3.Enabled = False
    End If
End Sub
