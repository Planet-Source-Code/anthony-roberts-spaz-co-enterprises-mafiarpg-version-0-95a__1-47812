VERSION 5.00
Begin VB.Form frmGarage 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Leave"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sell"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Purchase"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Automobiles"
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton Option2 
         Alignment       =   1  'Right Justify
         Caption         =   "Firefly"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton Option5 
         Alignment       =   1  'Right Justify
         Caption         =   "GTO-V7"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Alignment       =   1  'Right Justify
         Caption         =   "Viper"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton Option3 
         Alignment       =   1  'Right Justify
         Caption         =   "Mustang"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Taxi"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      Caption         =   "Welcome to the Garage"
      Height          =   855
      Left            =   1560
      TabIndex        =   7
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      Height          =   2775
      Left            =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmGarage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If strCompany <> "BA" Then
        If Option1.Value = True Then
            If bolTaxi = False Then
                If intMoney >= 1000 Then
                    intMoney = intMoney - 1000
                    lblStats.Caption = "You purchase a Taxi for $1,000"
                    bolTaxi = True
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own a this!"
            End If
        ElseIf Option2.Value = True Then
            If bolFirefly = False Then
                If intMoney >= 2000 Then
                    intMoney = intMoney - 2000
                    lblStats.Caption = "You purchase a Firefly for $2,000"
                    bolFirefly = True
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own this!"
            End If
        ElseIf Option3.Value = True Then
            If bolMustang = False Then
                If intMoney >= 5000 Then
                    intMoney = intMoney - 5000
                    lblStats.Caption = "You purchase a Mustang for $5,000"
                    bolMustang = True
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own this!"
            End If
        ElseIf Option4.Value = True Then
            If bolViper = False Then
                If intMoney >= 10000 Then
                    intMoney = intMoney - 10000
                    lblStats.Caption = "You purchase a Viper for $10,000"
                    bolViper = True
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own this!"
            End If
        ElseIf Option5.Value = True Then
            If bolGTO = False Then
                If intMoney >= 50000 Then
                    intMoney = intMoney - 50000
                    lblStats.Caption = "You purchase a GTO-V7 for $50,000"
                    bolGTO = True
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own this!"
            End If
        Else
            MsgBox "Make a selection!"
        End If
    Else
        If Option1.Value = True Then
            If bolTaxi = False Then
                If intMoney >= 750 Then
                    intMoney = intMoney - 750
                    lblStats.Caption = "You purchase a Taxi for $750"
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own a this!"
            End If
        ElseIf Option2.Value = True Then
            If bolFirefly = False Then
                If intMoney >= 1500 Then
                    intMoney = intMoney - 1500
                    lblStats.Caption = "You purchase a Firefly for $1,500"
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own this!"
            End If
        ElseIf Option3.Value = True Then
            If bolMustang = False Then
                If intMoney >= 3000 Then
                    intMoney = intMoney - 3000
                    lblStats.Caption = "You purchase a Mustang for $3,000"
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own this!"
            End If
        ElseIf Option4.Value = True Then
            If bolViper = False Then
                If intMoney >= 7500 Then
                    intMoney = intMoney - 7500
                    lblStats.Caption = "You purchase a Viper for $7,500"
                Else
                    MsgBox "You don't have enough money!"
                End If
            Else
                MsgBox "You already own this!"
            End If
        Else
            MsgBox "Make a selection!"
        End If
    End If
    If bolTaxi = True Then
        Option1.ForeColor = &HC0&
    End If
    If bolFirefly = True Then
        Option2.ForeColor = &HC0&
    End If
    If bolMustang = True Then
        Option3.ForeColor = &HC0&
    End If
    If bolViper = True Then
        Option4.ForeColor = &HC0&
    End If
    If bolGTO = True Then
        Option5.ForeColor = &HC0&
    End If
End Sub

Private Sub Command2_Click()
    If Option1.Value = True Then
        If bolTaxi = True Then
            bolTaxi = False
            intMoney = intMoney + 500
            lblStats.Caption = "You sold a Taxi for $500"
        Else
            MsgBox "You don't own this!"
        End If
    ElseIf Option2.Value = True Then
        If bolFirefly = True Then
            bolFirefly = False
            intMoney = intMoney + 1000
            lblStats.Caption = "You sold a Firefly for $1,000"
        Else
            MsgBox "You don't own this!"
        End If
    ElseIf Option3.Value = True Then
        If bolMustang = True Then
            bolMustang = False
            intMoney = intMoney + 2500
            lblStats.Caption = "You sold a Mustang for $2,500"
        Else
            MsgBox "You don't own this!"
        End If
    ElseIf Option4.Value = True Then
        If bolViper = True Then
            bolViper = False
            intMoney = intMoney + 7000
            lblStats.Caption = "You sold a Viper for $7,000"
        Else
            MsgBox "You don't own this!"
        End If
    ElseIf Option5.Value = True Then
        If bolGTO = True Then
            bolGTO = False
            intMoney = intMoney + 30000
            lblStats.Caption = "You sold a GTO for $30,000"
        Else
            MsgBox "You don't own this!"
        End If
    Else
        MsgBox "Make a selection!"
    End If
    If bolTaxi <> True Then
        Option1.ForeColor = &H0&
    End If
    If bolFirefly <> True Then
        Option2.ForeColor = &H0&
    End If
    If bolMustang <> True Then
        Option3.ForeColor = &H0&
    End If
    If bolViper <> True Then
        Option4.ForeColor = &H0&
    End If
    If bolGTO <> True Then
        Option5.ForeColor = &H0&
    End If
End Sub

Private Sub Command3_Click()
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    Unload frmGarage
End Sub

Private Sub Form_Load()
    If strCompany = "BA" Then
        Option5.Enabled = False
    End If
    
    If bolTaxi = True Then
        Option1.ForeColor = &HC0&
    End If
    If bolFirefly = True Then
        Option2.ForeColor = &HC0&
    End If
    If bolMustang = True Then
        Option3.ForeColor = &HC0&
    End If
    If bolViper = True Then
        Option4.ForeColor = &HC0&
    End If
    If bolGTO = True Then
        Option5.ForeColor = &HC0&
    End If
    
End Sub

Private Sub Option1_Click()
    lblStats.Caption = "Taxi" & vbLf & vbLf & "Gataway Speed: 2"
    Command1.Caption = "Purchase" & vbLf & "$1,000"
    Command2.Caption = "Sell" & vbLf & "$500"
    If strCompany = "BA" Then
        Command1.Caption = "Purchase" & vbLf & "$750"
        Command2.Caption = "Sell" & vbLf & "$500"
    End If
End Sub

Private Sub Option2_Click()
    lblStats.Caption = "Firefly" & vbLf & vbLf & "Gataway Speed: 4"
    Command1.Caption = "Purchase" & vbLf & "$2,000"
    Command2.Caption = "Sell" & vbLf & "$1,000"
    If strCompany = "BA" Then
        Command1.Caption = "Purchase" & vbLf & "$1,500"
        Command2.Caption = "Sell" & vbLf & "$1,000"
    End If
End Sub

Private Sub Option3_Click()
    lblStats.Caption = "Mustang" & vbLf & vbLf & "Gataway Speed: 6"
    Command1.Caption = "Purchase" & vbLf & "$5,000"
    Command2.Caption = "Sell" & vbLf & "$2,500"
    If strCompany = "BA" Then
        Command1.Caption = "Purchase" & vbLf & "$3,000"
        Command2.Caption = "Sell" & vbLf & "$2,500"
    End If
End Sub

Private Sub Option4_Click()
    lblStats.Caption = "Viper" & vbLf & vbLf & "Gataway Speed: 8"
    Command1.Caption = "Purchase" & vbLf & "$10,000"
    Command2.Caption = "Sell" & vbLf & "$7,000"
    If strCompany = "BA" Then
        Command1.Caption = "Purchase" & vbLf & "$7,500"
        Command2.Caption = "Sell" & vbLf & "$7,000"
    End If
End Sub

Private Sub Option5_Click()
    lblStats.Caption = "GTO-V7" & vbLf & vbLf & "Gataway Speed: 10"
    Command1.Caption = "Purchase" & vbLf & "$50,000"
    Command2.Caption = "Sell" & vbLf & "$30,000"
End Sub
