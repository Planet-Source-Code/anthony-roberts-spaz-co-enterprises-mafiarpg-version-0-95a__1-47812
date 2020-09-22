VERSION 5.00
Begin VB.Form frmInventory 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Ride this Car Around"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Automobiles"
      Height          =   975
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   4335
      Begin VB.OptionButton Option9 
         Alignment       =   1  'Right Justify
         Caption         =   "Taxi"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option10 
         Alignment       =   1  'Right Justify
         Caption         =   "Firefly"
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option11 
         Alignment       =   1  'Right Justify
         Caption         =   "Mustang"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option12 
         Alignment       =   1  'Right Justify
         Caption         =   "Viper"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option13 
         Alignment       =   1  'Right Justify
         Caption         =   "GTO-V7"
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Leave"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Equip / Unequip"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Weapons"
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Brass Knuckles"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gold Knuckles"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Steel Chains"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Switchblade"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pistol"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton Option6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Uzi"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2160
         Width           =   1935
      End
      Begin VB.OptionButton Option7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Assault Rifle"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1935
      End
      Begin VB.OptionButton Option8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Punisher .77mm"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   1935
      End
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      Caption         =   "Cash: $"
      Height          =   255
      Left            =   2400
      TabIndex        =   19
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      Height          =   975
      Left            =   2400
      TabIndex        =   18
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   2400
      TabIndex        =   17
      Top             =   960
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00008080&
      BorderWidth     =   3
      Height          =   5175
      Left            =   0
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option1.Value = True Then
        If bolBrass = True Then
            If strWeap = "Brass Knuckles" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option1.ForeColor = &H0&
            Else
                strWeap = "Brass Knuckles"
                intWeap = 2
                intTrain = 0
                Option1.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    ElseIf Option2.Value = True Then
        If bolGold = True Then
            If strWeap = "Gold Knuckles" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option2.ForeColor = &H0&
            Else
                strWeap = "Gold Knuckles"
                intWeap = 3
                intTrain = 0
                Option2.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    ElseIf Option3.Value = True Then
        If bolChain = True Then
            If strWeap = "Chains" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option3.ForeColor = &H0&
            Else
                strWeap = "Chains"
                intWeap = 4
                intTrain = 0
                Option3.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    ElseIf Option4.Value = True Then
        If bolSwitch = True Then
            If strWeap = "Switchblade" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option4.ForeColor = &H0&
            Else
                strWeap = "Switchblade"
                intWeap = 5
                intTrain = 0
                Option4.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    ElseIf Option5.Value = True Then
        If bolPistol = True Then
            If strWeap = "Pistol" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option5.ForeColor = &H0&
            Else
                strWeap = "Pistol"
                intWeap = 6
                intTrain = 0
                Option5.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    ElseIf Option6.Value = True Then
        If bolUzi = True Then
            If strWeap = "Uzi" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option6.ForeColor = &H0&
            Else
                strWeap = "Uzi"
                intWeap = 7
                intTrain = 0
                Option6.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    ElseIf Option7.Value = True Then
        If bolAssault = True Then
            If strWeap = "Assault Rifle" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option7.ForeColor = &H0&
            Else
                strWeap = "Assault Rifle"
                intWeap = 8
                intTrain = 0
                Option7.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    ElseIf Option8.Value = True Then
        If bolPunisher = True Then
            If strWeap = "Punisher" Then
                strWeap = "Fists"
                intWeap = 1
                intTrain = 0
                Option8.ForeColor = &H0&
            Else
                strWeap = "Punisher"
                intWeap = 10
                intTrain = 0
                Option8.ForeColor = &HC00000
            End If
        Else
            MsgBox "You don't own this weapon!"
        End If
    Else
        MsgBox "Select a weapon!"
    End If
    
    If strWeap <> "Brass Knuckles" Then
        Option1.ForeColor = &H0&
    End If
    If strWeap <> "Gold Knuckles" Then
        Option2.ForeColor = &H0&
    End If
    If strWeap <> "Chains" Then
        Option3.ForeColor = &H0&
    End If
    If strWeap <> "Switchblade" Then
        Option4.ForeColor = &H0&
    End If
    If strWeap <> "Pistol" Then
        Option5.ForeColor = &H0&
    End If
    If strWeap <> "Uzi" Then
        Option6.ForeColor = &H0&
    End If
    If strWeap <> "Assault Rifle" Then
        Option7.ForeColor = &H0&
    End If
    If strWeap <> "Punisher" Then
        Option8.ForeColor = &H0&
    End If
End Sub

Private Sub Command2_Click()
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    Unload frmInventory
End Sub

Private Sub Command3_Click()
    If Option9.Value = True Then
        If bolTaxi = True Then
            intSpeed = 2
            MsgBox "You hop into your Taxi!"
            Option9.ForeColor = &HC00000
        Else
            MsgBox "You don't own this car!"
        End If
    ElseIf Option10.Value = True Then
        If bolFirefly = True Then
            intSpeed = 4
            MsgBox "You hop into your Firefly!"
            Option10.ForeColor = &HC00000
        Else
            MsgBox "You don't own this car!"
        End If
    ElseIf Option11.Value = True Then
        If bolMustang = True Then
            intSpeed = 6
            MsgBox "You hop into your Mustang!"
            Option11.ForeColor = &HC00000
        Else
            MsgBox "You don't own this car!"
        End If
    ElseIf Option12.Value = True Then
        If bolViper = True Then
            intSpeed = 8
            MsgBox "You hop into your Viper!"
            Option12.ForeColor = &HC00000
        Else
            MsgBox "You don't own this car!"
        End If
    ElseIf Option13.Value = True Then
        If bolGTO = True Then
            intSpeed = 10
            MsgBox "You hop into your GTO-V7!"
            Option13.ForeColor = &HC00000
        Else
            MsgBox "You don't own this car!"
        End If
    Else
        MsgBox "Make a selection!"
    End If
    
    If bolTaxi = False Then
        Option9.ForeColor = &H80000012
    End If
    If bolFirefly = False Then
        Option10.ForeColor = &H80000012
    End If
    If bolMustang = False Then
        Option11.ForeColor = &H80000012
    End If
    If bolViper = False Then
        Option12.ForeColor = &H80000012
    End If
    If bolGTO = False Then
        Option13.ForeColor = &H80000012
    End If
End Sub

Private Sub Form_Load()
    If strWeap = "Brass Knuckles" Then
        Option1.ForeColor = &HC00000
    End If
    If strWeap = "Gold Knuckles" Then
        Option2.ForeColor = &HC00000
    End If
    If strWeap = "Chains" Then
        Option3.ForeColor = &HC00000
    End If
    If strWeap = "Switchblade" Then
        Option4.ForeColor = &HC00000
    End If
    If strWeap = "Pistol" Then
        Option5.ForeColor = &HC00000
    End If
    If strWeap = "Uzi" Then
        Option6.ForeColor = &HC00000
    End If
    If strWeap = "Assault Rifle" Then
        Option7.ForeColor = &HC00000
    End If
    If strWeap = "Punisher" Then
        Option8.ForeColor = &HC00000
    End If
End Sub

Private Sub Option1_Click()
    lblName.Caption = "Brass Knuckles"
    lblStats.Caption = "Attack: 2" & vbLf & "Value: $20"
End Sub

Private Sub Option2_Click()
    lblName.Caption = "Gold Knuckles"
    lblStats.Caption = "Attack: 3" & vbLf & "Value: $100"
End Sub

Private Sub Option3_Click()
    lblName.Caption = "Steel Chains"
    lblStats.Caption = "Attack: 4" & vbLf & "Value: $250"
End Sub

Private Sub Option4_Click()
    lblName.Caption = "Switchblade"
    lblStats.Caption = "Attack: 5" & vbLf & "Value: $500"
End Sub

Private Sub Option5_Click()
    lblName.Caption = "Pistol"
    lblStats.Caption = "Attack: 6" & vbLf & "Value: $1000"
End Sub

Private Sub Option6_Click()
    lblName.Caption = "Uzi"
    lblStats.Caption = "Attack: 7" & vbLf & "Value: $2500"
End Sub

Private Sub Option7_Click()
    lblName.Caption = "Assault Rifle"
    lblStats.Caption = "Attack: 8" & vbLf & "Value: $5000"
End Sub

Private Sub Option8_Click()
    lblName.Caption = "Punisher .77mm"
    lblStats.Caption = "Attack: 10" & vbLf & "Value: $10000"
End Sub
