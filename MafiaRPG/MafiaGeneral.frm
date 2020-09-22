VERSION 5.00
Begin VB.Form frmGeneral 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Leave"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Purchase a City Map - $10"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Purchase a Hot Dog - $4"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Purchase a Bottle of Pop - $2"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Purchase a Condom - $1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    frmRPG.Enabled = True
    frmRPG.Visible = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    Unload frmGeneral
End Sub

Private Sub Command2_Click()
    If intMoney >= 1 Then
        intMoney = intMoney - 1
        bolCondom = True
    Else
        MsgBox "You don't have enough money!"
    End If
End Sub

Private Sub Command3_Click()
    If intMoney >= 3 Then
        If (intHP + 4) <= intMaxHP Then
            intHP = intHP + 4
            intMoney = intMoney - 3
        Else
            MsgBox "You have too much health for one of these!"
        End If
    Else
        MsgBox "You don't have enough money!"
    End If
End Sub

Private Sub Command4_Click()
    If intMoney >= 2 Then
        If (intHP + 2) <= intMaxHP Then
            intHP = intHP + 2
            intMoney = intMoney - 2
        Else
            MsgBox "You have too much health for one of these!"
        End If
    Else
        MsgBox "You don't have enough money!"
    End If
End Sub

Private Sub Command5_Click()
    If intMoney >= 10 Then
        If CityMap = True Then
            MsgBox "You already own a City Map!"
        Else
            CityMap = True
            Command5.Enabled = False
        End If
    Else
        MsgBox "You don't have enough money!"
    End If
End Sub

Private Sub Form_Load()
    If CityMap = True Then
        Command5.Enabled = False
    End If
End Sub
