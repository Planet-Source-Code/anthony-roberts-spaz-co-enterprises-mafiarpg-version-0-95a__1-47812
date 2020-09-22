VERSION 5.00
Begin VB.Form frmLoan 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Leave"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Take this Loan"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pay Loan"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblLoan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "My current interest rate: 50%"
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "How much ya wanna loan?"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   5
      X1              =   0
      X2              =   4560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jack the Loan Sharks Super Loan Discount Store"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   10
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    StartLoan = Text1.Text
    If StartLoan < 1 Or StartLoan > 10000 Then
        MsgBox "You can not loan this value! (Must be between $1 and $10000)"
        StartLoan = 0
        Exit Sub
    Else
        MsgBox "You take out the $" & StartLoan & " loan"
        CurrentLoan = StartLoan
        intMoney = intMoney + StartLoan
        intLoanTotalTime = 10
        Command1.Enabled = False
        Command2.Enabled = True
    End If
    If strCompany = "HC" Then
        lblLoan.Caption = "My current interest rate: 20%" & vbLf & "You owe me: $" & CurrentLoan
    Else
        lblLoan.Caption = "My current interest rate: 35%" & vbLf & "You owe me: $" & CurrentLoan
    End If
End Sub

Private Sub Command2_Click()
    If intMoney >= CurrentLoan Then
        intMoney = intMoney - CurrentLoan
        If strCompany = "HC" Then
            lblLoan.Caption = "My current interest rate: 20%" & vbLf & "You owe me: $" & CurrentLoan
        Else
            lblLoan.Caption = "My current interest rate: 35%" & vbLf & "You owe me: $" & CurrentLoan
        End If
    Else
        MsgBox "You've not enough money to pay this loan!"
    End If
End Sub

Private Sub Command3_Click()
    frmRPG.Enabled = True
    frmRPG.Visible = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    Unload frmLoan
End Sub

Private Sub Form_Load()
    If strCompany = "HC" Then
        lblLoan.Caption = "My current interest rate: 20%" & vbLf & "You owe me: $" & CurrentLoan
    Else
        lblLoan.Caption = "My current interest rate: 35%" & vbLf & "You owe me: $" & CurrentLoan
    End If
    If CurrentLoan > 0 Then
        Command1.Enabled = False
    Else
        Command2.Enabled = False
    End If
End Sub

