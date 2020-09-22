VERSION 5.00
Begin VB.Form frmLacky 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Leave"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Bouncer $350 Each"
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Chem'y $300 Each"
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Prostitutes $150 Each"
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Hire / Lay Off"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton fire10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Fire      10"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton fire1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Fire        1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton buy10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Buy     10"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton buy1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808080&
         Caption         =   "Buy       1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Bouncers go kill enemies and bring back their cash."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   11
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Chem'ys make drugs and sell them to druggies."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   10
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Prostitutes go out on the street and each some cash."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Lackys"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmLacky"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If buy1.Value = True Then
        If intMoney >= 150 Then
            intPros = intPros + 1
            intMoney = intMoney - 150
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf buy10.Value = True Then
        If intMoney >= (150 * 10) Then
            intPros = intPros + 10
            intMoney = intMoney - (150 * 10)
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf fire1.Value = True Then
        If intPros >= 1 Then
            intPros = intPros - 1
        Else
            MsgBox "You don't have one!"
        End If
    ElseIf fire10.Value = True Then
        If intPros >= 10 Then
            intPros = intPros - 10
        Else
            MsgBox "You don't have this many!"
        End If
    End If
    Command1.Caption = "Prostitutes (" & intPros & ")" & vbLf & "$150 Each"
    Command2.Caption = "Chem'y (" & intChem & ")" & vbLf & "$300 Each"
    Command3.Caption = "Bouncer (" & intBounce & ")" & vbLf & "$350 Each"
End Sub

Private Sub Command2_Click()
    If buy1.Value = True Then
        If intMoney >= 300 Then
            intChem = intChem + 1
            intMoney = intMoney - 300
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf buy10.Value = True Then
        If intMoney >= (300 * 10) Then
            intChem = intChem + 10
            intMoney = intMoney - (300 * 10)
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf fire1.Value = True Then
        If intChem >= 1 Then
            intChem = intChem - 1
        Else
            MsgBox "You don't have one!"
        End If
    ElseIf fire10.Value = True Then
        If intChem >= 10 Then
            intChem = intChem - 10
        Else
            MsgBox "You don't have this many!"
        End If
    End If
    Command1.Caption = "Prostitutes (" & intPros & ")" & vbLf & "$150 Each"
    Command2.Caption = "Chem'y (" & intChem & ")" & vbLf & "$300 Each"
    Command3.Caption = "Bouncer (" & intBounce & ")" & vbLf & "$350 Each"
End Sub

Private Sub Command3_Click()
    If buy1.Value = True Then
        If intMoney >= 350 Then
            intBounce = intBounce + 1
            intMoney = intMoney - 350
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf buy10.Value = True Then
        If intMoney >= (350 * 10) Then
            intBounce = intBounce + 10
            intMoney = intMoney - (350 * 10)
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf fire1.Value = True Then
        If intBounce >= 1 Then
            intBounce = intBounce - 1
        Else
            MsgBox "You don't have one!"
        End If
    ElseIf fire10.Value = True Then
        If intBounce >= 10 Then
            intBounce = intBounce - 10
        Else
            MsgBox "You don't have this many!"
        End If
    End If
    Command1.Caption = "Prostitutes (" & intPros & ")" & vbLf & "$150 Each"
    Command2.Caption = "Chem'y (" & intChem & ")" & vbLf & "$300 Each"
    Command3.Caption = "Bouncer (" & intBounce & ")" & vbLf & "$350 Each"
End Sub

Private Sub Command4_Click()
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    Unload frmLacky
End Sub

Private Sub Form_Load()
    Command1.Caption = "Prostitutes (" & intPros & ")" & vbLf & "$150 Each"
    Command2.Caption = "Chem'y (" & intChem & ")" & vbLf & "$300 Each"
    Command3.Caption = "Bouncer (" & intBounce & ")" & vbLf & "$350 Each"
End Sub
