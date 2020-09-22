VERSION 5.00
Begin VB.Form frmSnack 
   BorderStyle     =   0  'None
   Caption         =   "The ""Snack"" Shack"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Use Drugs"
      Height          =   2055
      Left            =   5400
      TabIndex        =   25
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command21 
         Caption         =   "Cure AIDS"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command20 
         Caption         =   "+100 Max HP"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command19 
         Caption         =   "+50 Max HP"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command18 
         Caption         =   "+25 Max HP"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command17 
         Caption         =   "+ 10 Max HP"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command16 
         Caption         =   "+ 5 Max HP"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command15 
         Caption         =   "+5 HP"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sell Drugs"
      Height          =   2055
      Left            =   3720
      TabIndex        =   2
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command14 
         Caption         =   "-- NIL --"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command13 
         Caption         =   "$2000"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command12 
         Caption         =   "$500"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command11 
         Caption         =   "$250"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "$150"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "$70"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "$30"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buy Drugs"
      Height          =   2055
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      Begin VB.CommandButton Command7 
         Caption         =   "$10000"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "$2500"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "$1000"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "$750"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "$500"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "$250"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "$150"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2280
      Width           =   4935
   End
   Begin VB.Timer Border 
      Interval        =   50
      Left            =   120
      Top             =   2040
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Cash: $0"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Calf-Een"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Coka-Cokea"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Hair-O-Win"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Lucky Charm"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Koocie"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Lazy Dazie"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Spoon Dope"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   1815
   End
   Begin VB.Shape shpBack 
      BorderColor     =   &H00180C01&
      BorderWidth     =   10
      Height          =   2895
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmSnack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Border_Timer()
    If intBorder = 1 Then
        shpBack.BorderColor = &HFCE1CB
        intBorder = 2
    ElseIf intBorder = 2 Then
        shpBack.BorderColor = &HFBD0AE
        intBorder = 3
    ElseIf intBorder = 3 Then
        shpBack.BorderColor = &HFBD0AE
        intBorder = 4
    ElseIf intBorder = 4 Then
        shpBack.BorderColor = &HFBD0AE
        intBorder = 5
    ElseIf intBorder = 5 Then
        shpBack.BorderColor = &HF57B16
        intBorder = 6
    ElseIf intBorder = 6 Then
        shpBack.BorderColor = &HF57B16
        intBorder = 7
    ElseIf intBorder = 7 Then
        shpBack.BorderColor = &HF57B16
        intBorder = 8
    ElseIf intBorder = 8 Then
        shpBack.BorderColor = &HF57B16
        intBorder = 9
    ElseIf intBorder = 9 Then
        shpBack.BorderColor = &H180C01
        intBorder = 10
    ElseIf intBorder = 10 Then
        shpBack.BorderColor = &HFCE1CB
        intBorder = 1
    End If
End Sub

Private Sub cmdLeave_Click()
    frmRPG.Enabled = True
    frmRPG.Visible = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    frmRPG.lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    Unload frmSnack
End Sub

Private Sub Command1_Click()
    If strCompany = "BI" Then
        If intMoney >= 125 Then
            intMoney = intMoney - 125
            intSpoon = intSpoon + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    Else
        If intMoney >= 150 Then
            intMoney = intMoney - 150
            intSpoon = intSpoon + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command10_Click()
    If intKoocie >= 1 Then
        intKoocie = intKoocie - 1
        If strCompany = "SHE" Then
            intMoney = intMoney + 120
        Else
            intMoney = intMoney + 150
        End If
    Else
        MsgBox "You have none of this drug to sell"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command11_Click()
    If intLucky >= 1 Then
        intLucky = intLucky - 1
        If strCompany = "SHE" Then
            intMoney = intMoney + 200
        Else
            intMoney = intMoney + 250
        End If
    Else
        MsgBox "You have none of this drug to sell"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command12_Click()
    If intHair >= 1 Then
        intHair = intHair - 1
        If strCompany = "SHE" Then
            intMoney = intMoney + 375
        Else
            intMoney = intMoney + 500
        End If
    Else
        MsgBox "You have none of this drug to sell"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command13_Click()
    If intCoka >= 1 Then
        intCoka = intCoka - 1
        If strCompany = "SHE" Then
            intMoney = intMoney + 1000
        Else
            intMoney = intMoney + 2000
        End If
    Else
        MsgBox "You have none of this drug to sell"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command15_Click()
    If intSpoon >= 1 Then
        intSpoon = intSpoon - 1
        intHP = intHP + 5
        MsgBox "You heal for 5."
    Else
        MsgBox "You have none of this drug to use"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command16_Click()
    If intLazy >= 1 Then
        intLazy = intLazy - 1
        intMaxHP = intMaxHP + 5
        intHP = intHP + 5
        MsgBox "You raise your Hit Points by 5!"
    Else
        MsgBox "You have none of this drug to use"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command17_Click()
    If intKoocie >= 1 Then
        intKoocie = intKoocie - 1
        intMaxHP = intMaxHP + 10
        intHP = intHP + 10
        MsgBox "You raise your Hit Points by 10!"
    Else
        MsgBox "You have none of this drug to use"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command18_Click()
    If intLucky >= 1 Then
        intLucky = intLucky - 1
        intMaxHP = intMaxHP + 25
        intHP = intHP + 25
        MsgBox "You raise your Hit Points by 25!"
    Else
        MsgBox "You have none of this drug to use"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command19_Click()
    If intHair >= 1 Then
        intHair = intHair - 1
        intMaxHP = intMaxHP + 50
        intHP = intHP + 50
        MsgBox "You raise your Hit Points by 50!"
    Else
        MsgBox "You have none of this drug to use"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command2_Click()
    If strCompany = "BI" Then
        If intMoney >= 200 Then
            intMoney = intMoney - 200
            intLazy = intLazy + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    Else
        If intMoney >= 250 Then
            intMoney = intMoney - 250
            intLazy = intLazy + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command20_Click()
    If intCoka >= 1 Then
        intCoka = intCoka - 1
        intMaxHP = intMaxHP + 100
        intHP = intHP + 100
        MsgBox "You raise your Hit Points by 100!"
    Else
        MsgBox "You have none of this drug to use"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command21_Click()
    If intCalf >= 1 Then
        intCalf = intCalf - 1
        bolAIDS = False
        MsgBox "You're cured of the AIDS virus!"
    Else
        MsgBox "You have none of this drug to use"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command3_Click()
    If strCompany = "BI" Then
        If intMoney >= 400 Then
            intMoney = intMoney - 400
            intKoocie = intKoocie + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    Else
        If intMoney >= 500 Then
            intMoney = intMoney - 500
            intKoocie = intKoocie + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command4_Click()
    If strCompany = "BI" Then
        If intMoney >= 650 Then
            intMoney = intMoney - 650
            intLucky = intLucky + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    Else
        If intMoney >= 750 Then
            intMoney = intMoney - 750
            intLucky = intLucky + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command5_Click()
    If strCompany = "BI" Then
        If intMoney >= 750 Then
            intMoney = intMoney - 750
            intHair = intHair + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    Else
        If intMoney >= 1000 Then
            intMoney = intMoney - 1000
            intHair = intHair + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command6_Click()
    If strCompany = "BI" Then
        If intMoney >= 1750 Then
            intMoney = intMoney - 1750
            intCoka = intCoka + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    Else
        If intMoney >= 2500 Then
            intMoney = intMoney - 2500
            intCoka = intCoka + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command7_Click()
    If strCompany = "BI" Then
        If intMoney >= 7500 Then
            intMoney = intMoney - 7500
            intCalf = intCalf + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    Else
        If intMoney >= 10000 Then
            intMoney = intMoney - 10000
            intCalf = intCalf + 1
        Else
            MsgBox "You do not have enough money to buy this."
        End If
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command8_Click()
    If intSpoon >= 1 Then
        intSpoon = intSpoon - 1
        If strCompany = "SHE" Then
            intMoney = intMoney + 20
        Else
            intMoney = intMoney + 30
        End If
    Else
        MsgBox "You have none of this drug to sell"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Command9_Click()
    If intLazy >= 1 Then
        intLazy = intLazy - 1
        If strCompany = "SHE" Then
            intMoney = intMoney + 50
        Else
            intMoney = intMoney + 70
        End If
    Else
        MsgBox "You have none of this drug to sell"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
End Sub

Private Sub Form_Load()
    intBorder = 1
    If strCompany = "CNS" Then
        Command7.Caption = "$20000"
    ElseIf strCompany = "BI" Then
        Command1.Caption = "$125"
        Command2.Caption = "$200"
        Command3.Caption = "$400"
        Command4.Caption = "$650"
        Command5.Caption = "$750"
        Command6.Caption = "$1750"
        Command7.Caption = "$7500"
    ElseIf strCompany = "SHE" Then
        Command8.Caption = "$20"
        Command9.Caption = "$50"
        Command10.Caption = "$120"
        Command11.Caption = "$200"
        Command12.Caption = "$375"
        Command13.Caption = "$1000"
    End If
    Label1.Caption = "Spoon Dope (" & intSpoon & ")"
    Label2.Caption = "Lazy Dazie (" & intLazy & ")"
    Label3.Caption = "Koocie (" & intKoocie & ")"
    Label4.Caption = "Lucky Charm (" & intLucky & ")"
    Label5.Caption = "Hair-O-Win (" & intHair & ")"
    Label6.Caption = "Coka-Cokea (" & intCoka & ")"
    Label7.Caption = "Calf-Een (" & intCalf & ")"
    Label8.Caption = "Cash: $" & intMoney
    If bolCalfAvailable = True Then
        Command7.Visible = True
        Label7.Visible = True
        Command14.Visible = True
        Command21.Visible = True
    Else
        Command7.Visible = False
        Label7.Visible = False
        Command14.Visible = False
        Command21.Visible = False
    End If
End Sub
