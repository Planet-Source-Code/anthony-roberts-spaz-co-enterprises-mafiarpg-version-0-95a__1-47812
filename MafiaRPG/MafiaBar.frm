VERSION 5.00
Begin VB.Form frmBar 
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
   Begin VB.Timer Bar 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   4080
      Top             =   2520
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Show chick your weapon"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      ToolTipText     =   "Pull out your gun and wave it around (No man, your REAL gun.  Then one with the trigger and bullets... You sick minded freak!)"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Grab chicks pussy"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "You'd like to stick your willie here... why not your hand for now"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Grab chicks breasts"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Her breasts need a LOT of attention..."
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Slap the chick"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   5
      ToolTipText     =   "Skap that submissive bitch"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Bang this chick"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      ToolTipText     =   "Make your move!"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buy chick a drink"
      Enabled         =   0   'False
      Height          =   495
      Left            =   600
      TabIndex        =   3
      ToolTipText     =   "For $5, you can try to get this woman drunk."
      Top             =   1080
      Width           =   1695
   End
   Begin VB.CommandButton cmdDrink 
      Caption         =   "Get a Drink"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      ToolTipText     =   "Get yourself a drink for $5 (Hit Points raise 1)"
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdChick 
      Caption         =   "Meet some Chicks"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      ToolTipText     =   "Find a chick to bang (Having sex will raise your max Hit Points, but they cost!"
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdReturn 
      Caption         =   "Leave"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Leave the Bar"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label lblMeet 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the Local Bar"
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   840
      Width           =   3735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000040C0&
      BorderStyle     =   2  'Dash
      BorderWidth     =   5
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Shape           =   2  'Oval
      Top             =   0
      Width           =   4695
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderWidth     =   5
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bar_Timer()
    bolBar = True
    Bar.Enabled = False
    frmRPG.Bar.Enabled = False
End Sub

Private Sub cmdChick_Click()
    If bolBar = True Then
        Bar.Enabled = True
        frmRPG.Bar.Enabled = True
        bolBar = False
        intChance = Rnd * 100
        If intChance >= 1 And intChance <= 40 Then
            lblMeet.Caption = "You meet up with a He/She ($50/5HP)"
            intBangCost = 50
            intHPIncreseChick = 5
        ElseIf intChance >= 41 And intChance <= 70 Then
            lblMeet.Caption = "You meet up with a Blonde ($70/10HP)"
            intBangCost = 70
            intHPIncreseChick = 10
        ElseIf intChance >= 71 And intChance <= 85 Then
            lblMeet.Caption = "You meet up with a Redhead ($100/15HP)"
            intBangCost = 100
            intHPIncreseChick = 15
        ElseIf intChance >= 86 And intChance <= 95 Then
            lblMeet.Caption = "You meet up with a Model ($150/25HP)"
            intBangCost = 150
            intHPIncreseChick = 25
        ElseIf intChance >= 96 And intChance <= 100 Then
            lblMeet.Caption = "You meet up with a Super Model! ($300/60HP)"
            intBangCost = 300
            intHPIncreseChick = 60
        End If
        Command1.Enabled = True
        Command2.Enabled = True
        Command3.Enabled = True
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        intChance = Rnd * 6
        If intChance = 1 Then
        'Submissive Chick
            bolBreast = True
            bolPussy = True
            bolSlap = True
            bolWeap = True
            bolDrink = True
            intGetMax = 5
        ElseIf intChance = 2 Then
        'Not Submissive
            bolBreast = False
            bolPussy = False
            bolSlap = False
            bolWeap = False
            bolDrink = True
            intGetMax = 1
        ElseIf intChance = 3 Then
        'Weapon Liker
            bolBreast = False
            bolPussy = False
            bolSlap = False
            bolWeap = True
            bolDrink = True
            intGetMax = 2
        ElseIf intChance = 4 Then
        'Enjoys beating
            bolBreast = False
            bolPussy = False
            bolSlap = True
            bolWeap = True
            bolDrink = True
            intGetMax = 3
        ElseIf intChance = 5 Then
        'Submissive non-drinker
            bolBreast = True
            bolPussy = True
            bolSlap = True
            bolWeap = True
            bolDrink = False
            intGetMax = 4
        ElseIf intChance = 6 Then
        'Whore
            bolBreast = False
            bolPussy = False
            bolSlap = False
            bolWeap = False
            bolDrink = False
            intGetMax = 0
        End If
    Else
        MsgBox "Slow down!  You need to wait awhile before you can nab another chick"
    End If
End Sub

Private Sub cmdDrink_Click()
    If intMoney >= 5 Then
        intHP = intHP + 5
        MsgBox "You heal for 5"
    Else
        MsgBox "You need $5 to get a drink, which you don't have."
    End If
End Sub

Private Sub cmdReturn_Click()
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    frmRPG.lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    frmRPG.Visible = True
    frmRPG.Enabled = True
    frmMove.Visible = True
    Unload frmBar
End Sub

Private Sub Command1_Click()
    If intMoney >= 5 Then
        If bolDrink = True Then
            MsgBox "She enjoys the drink... looks like your getting somewhere"
            intCurrentGet = intCurrentGet + 1
            Command1.Enabled = False
        Else
            MsgBox "She dosn't drink!  Great, now she's leaving..."
            intGetMax = 0
            intCurrentGet = 0
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command6.Enabled = False
            lblMeet.Caption = "Welcome to the Local Bar"
        End If
    Else
        MsgBox "You don't have enough money to buy this poor lady a drink... Try something else"
    End If
End Sub

Private Sub Command2_Click()
    If intMoney >= intBangCost Then
        If intCurrentGet = intGetMax Then
            MsgBox "When you wake up, you notice that $" & intBangCost & " are missing... guess all that wasn't free!  You also gain " & intHPIncreseChick & " Maximum Hit Points"
            intMaxHP = intMaxHP + intHPIncreseChick
            intHP = intHP + intHPIncreseChick
            intMoney = intMoney - intBangCost
            intChance = Rnd * 100 + 1
            If bolCondom = False Then
                If strCompany = "CNS" Then
                    If intChance >= 0 And intChance <= 5 Then
                        bolAIDS = True
                        MsgBox "You caught AIDS!"
                    End If
                Else
                    If intChance >= 0 And intChance <= 10 Then
                        bolAIDS = True
                        MsgBox "You caught AIDS!"
                    End If
                End If
            Else
                If strCompany = "CNS" Then
                    If intChance >= 0 And intChance <= 2 Then
                        bolAIDS = True
                        MsgBox "You caught AIDS!"
                    End If
                Else
                    If intChance >= 0 And intChance <= 7 Then
                        bolAIDS = True
                        MsgBox "You caught AIDS!"
                    End If
                End If
            End If
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command6.Enabled = False
            lblMeet.Caption = "Welcome to the Local Bar"
        Else
            MsgBox "She looks at you with large eyes, then slaps you - walking off... oh well - Better luck next time"
            intGetMax = 0
            intCurrentGet = 0
            Command1.Enabled = False
            Command2.Enabled = False
            Command3.Enabled = False
            Command4.Enabled = False
            Command5.Enabled = False
            Command6.Enabled = False
            lblMeet.Caption = "Welcome to the Local Bar"
        End If
    Else
        MsgBox "You don't have enough money to hire this chick... piss off"
    End If
End Sub

Private Sub Command3_Click()
    If bolSlap = True Then
        MsgBox "Great, worked. (Real Description here was cleaned up! Hehe)"
        intCurrentGet = intCurrentGet + 1
        Command3.Enabled = False
    Else
        MsgBox "...She slapped you back!  Watch who your slapping next time!"
        intGetMax = 0
        intCurrentGet = 0
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        lblMeet.Caption = "Welcome to the Local Bar"
    End If
End Sub

Private Sub Command4_Click()
    If bolBreast = True Then
        MsgBox "Great, worked. (Real Description here was cleaned up! Hehe)"
        intCurrentGet = intCurrentGet + 1
        Command4.Enabled = False
    Else
        MsgBox "OUCH!  She just slugged you one with her ringed fist!  That's gonna hurt in the morning..."
        intGetMax = 0
        intCurrentGet = 0
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        lblMeet.Caption = "Welcome to the Local Bar"
    End If
End Sub

Private Sub Command5_Click()
    If bolPussy = True Then
        MsgBox "Great, worked. (Real Description here was cleaned up! Hehe)"
        intCurrentGet = intCurrentGet + 1
        Command5.Enabled = False
    Else
        MsgBox "HOLY ----... Ouch... BIG ouch... You've been kicked in the balls... You haven't fell over yet... let me help... *shove*"
        intGetMax = 0
        intCurrentGet = 0
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        lblMeet.Caption = "Welcome to the Local Bar"
    End If
End Sub

Private Sub Command6_Click()
    If bolWeap = True Then
        MsgBox "She smiles in glee and admires your awesome weapon."
        intCurrentGet = intCurrentGet + 1
        Command6.Enabled = False
    Else
        MsgBox "She screams and runs out of the bar, 'He's got a gun!'"
        intGetMax = 0
        intCurrentGet = 0
        Command1.Enabled = False
        Command2.Enabled = False
        Command3.Enabled = False
        Command4.Enabled = False
        Command5.Enabled = False
        Command6.Enabled = False
        lblMeet.Caption = "Welcome to the Local Bar"
    End If
End Sub

Private Sub Form_Load()
    Randomize
End Sub
