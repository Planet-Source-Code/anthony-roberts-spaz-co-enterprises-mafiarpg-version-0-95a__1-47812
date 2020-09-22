VERSION 5.00
Begin VB.Form frmRPG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MafiaRPG - v0.95a"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   407
   StartUpPosition =   2  'CenterScreen
   Begin MafiaRPG.Button cmdFlee 
      Height          =   495
      Left            =   1440
      Top             =   2520
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      tx              =   "Flee"
      enab            =   -1  'True
      font            =   "MafiaRPG.frx":0000
   End
   Begin MafiaRPG.Button cmdAtt 
      Height          =   495
      Left            =   120
      Top             =   2520
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      tx              =   "Attack"
      enab            =   -1  'True
      font            =   "MafiaRPG.frx":002C
   End
   Begin VB.Timer MonthTime 
      Interval        =   1000
      Left            =   3240
      Top             =   120
   End
   Begin VB.Timer Dis 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   3360
   End
   Begin VB.CommandButton Turbo3 
      Caption         =   ">>>"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Triple Speed"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Turbo2 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Double Speed"
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton Turbo1 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Normal Speed"
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton Messages 
      Caption         =   "City Map"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Monthly News Reports"
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   6120
      Width           =   5895
      Begin VB.Label lblBounce 
         Alignment       =   2  'Center
         Caption         =   "You have no Bouncers!"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3120
         TabIndex        =   35
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblChem 
         Alignment       =   2  'Center
         Caption         =   "You have no Chem'ys!"
         ForeColor       =   &H00008080&
         Height          =   375
         Left            =   3120
         TabIndex        =   34
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblPros 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "You have no Prostitutes!"
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   2775
      End
      Begin VB.Line Line12 
         X1              =   3000
         X2              =   3000
         Y1              =   120
         Y2              =   1080
      End
      Begin VB.Label lblNews 
         Alignment       =   2  'Center
         Caption         =   "Welcome to MafiaRPG"
         ForeColor       =   &H00808000&
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Timer Bar 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3720
      Top             =   120
   End
   Begin VB.Timer Time 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4200
      Top             =   120
   End
   Begin VB.CommandButton cmdHire 
      Caption         =   "Lackies"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      ToolTipText     =   "Need a slave?  Excellent..."
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdInventory 
      Caption         =   "Equipment"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      ToolTipText     =   "Tired of these Brass Knuckles?  Try out your new gun!"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdHQ 
      Caption         =   "Company HQ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Enter your companies Head Quarters"
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdArmory 
      Caption         =   "Armory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   12
      ToolTipText     =   "Need a better toy?  Heh Heh"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdStore 
      Caption         =   "General Store"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   11
      ToolTipText     =   "Have we got some ""general"" items"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdGarage 
      Caption         =   "Garage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Need a quick fix on your ride?"
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdArena 
      Caption         =   "Fighting Arena"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   9
      ToolTipText     =   "Beat your enemy to win some money"
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdStock 
      Caption         =   "Stock Market"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      ToolTipText     =   "Need to earn some money?  Stocks are good."
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdTattoo 
      Caption         =   "International Bank"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Deposit money so that when you die, you keep some.  It collects interest as well."
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdLoan 
      Caption         =   "Loan Shark"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      ToolTipText     =   "Need some quick dough?  Well... make sure you pay your loan!"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdBar 
      Caption         =   "Local Bar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      ToolTipText     =   "Find some prostitutes to increse your maximum Hit Points... but don't be frightened by their cost!"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCasino 
      Caption         =   "Casino"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      ToolTipText     =   "Wager your cash to make some, or lose some."
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdHotel 
      Caption         =   "Hotel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "For $50, you can heal all your HP in the Hotel."
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdDrugs 
      Caption         =   """Snack"" Shack"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Drug dealing, anyone?"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdGym 
      Caption         =   "Local Gym"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Improve your training skill for $100"
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "@"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5400
      TabIndex        =   18
      ToolTipText     =   "Options"
      Top             =   2640
      Width           =   615
   End
   Begin VB.Frame fraStats 
      Caption         =   "Yourself"
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   840
      Width           =   2415
      Begin VB.Label lblStats 
         Height          =   1215
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame fraEnemy 
      Caption         =   "Enemy"
      Height          =   1575
      Left            =   2640
      TabIndex        =   22
      Top             =   840
      Width           =   1935
      Begin VB.Label lblEnemyStats 
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton Pause 
      Caption         =   "||"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Pause"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "QUIT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Quit the Game"
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Height          =   135
      Left            =   4560
      TabIndex        =   42
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Left            =   4680
      TabIndex        =   41
      ToolTipText     =   "Percentage of XP - TNL."
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   1455
      Left            =   4680
      TabIndex        =   40
      ToolTipText     =   "Percentage of Training for Current Weapon."
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   5880
      TabIndex        =   39
      ToolTipText     =   "Time Remining in this Month"
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape MonthBar 
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   5880
      Top             =   960
      Width           =   120
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1305
      Left            =   5880
      Top             =   960
      Width           =   120
   End
   Begin VB.Line Line3 
      X1              =   8
      X2              =   8
      Y1              =   264
      Y2              =   256
   End
   Begin VB.Label Display 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   4560
      TabIndex        =   36
      Top             =   3360
      Width           =   735
   End
   Begin VB.Line Line14 
      X1              =   16
      X2              =   0
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Line Line13 
      X1              =   352
      X2              =   352
      Y1              =   256
      Y2              =   0
   End
   Begin VB.Line Line10 
      X1              =   344
      X2              =   352
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line4 
      X1              =   304
      X2              =   312
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Line Line11 
      X1              =   8
      X2              =   8
      Y1              =   360
      Y2              =   368
   End
   Begin VB.Line Line9 
      X1              =   296
      X2              =   296
      Y1              =   360
      Y2              =   368
   End
   Begin VB.Line Line8 
      X1              =   296
      X2              =   408
      Y1              =   400
      Y2              =   400
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      Caption         =   "Cash: $"
      Height          =   495
      Left            =   4560
      TabIndex        =   30
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Line Line7 
      X1              =   304
      X2              =   304
      Y1              =   200
      Y2              =   256
   End
   Begin VB.Line Line6 
      X1              =   184
      X2              =   192
      Y1              =   200
      Y2              =   200
   End
   Begin VB.Line Line5 
      X1              =   8
      X2              =   8
      Y1              =   256
      Y2              =   200
   End
   Begin VB.Label lblMonth 
      Alignment       =   2  'Center
      Caption         =   "Month:"
      Height          =   495
      Left            =   4560
      TabIndex        =   29
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Line Line2 
      X1              =   296
      X2              =   296
      Y1              =   264
      Y2              =   256
   End
   Begin VB.Line Line1 
      X1              =   296
      X2              =   352
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Label lblXP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   28
      Top             =   2280
      Width           =   495
   End
   Begin VB.Shape shpXPPercent 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   4680
      Top             =   1680
      Width           =   495
   End
   Begin VB.Shape shpBackTwo 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1500
      Left            =   4680
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblPercent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape shpTrain 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   15
      Left            =   4680
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape shpBack 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   1500
      Left            =   4680
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      Caption         =   "Messages:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   120
      TabIndex        =   26
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label lblEnemy 
      Alignment       =   2  'Center
      Caption         =   "You are attacked by a(n): [Old Man]"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmRPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bar_Timer()
    bolBar = True
    Bar.Enabled = False
    frmBar.Bar.Enabled = False
End Sub

Private Sub cmdArena_Click()
    frmArena.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdArmory_Click()
    frmArmory.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdAtt_Click()
    lblMessage.ForeColor = &HC00000
    If cmdAtt.Caption = "Next Battle!" Then
        cmdAtt.Caption = "Attack"
    End If
    intHit = (Rnd * intWeap) + intSpec
    If intHit > 0 Then
        intEnHP = intEnHP - intHit
        strMessage = "You hit the " & strEnemy & " for " & intHit & " damage!"
        intTrain = intTrain + 1
        shpTrain.Height = intTrain
        lblPercent = intTrain & "%"
        If intTrain >= 100 Then
            intSpec = intSpec + 1
            MsgBox "During battle you learn a new move!  You Special attack increses by one."
            lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
            intTrain = 0
            shpTrain.Height = 0
            lblPercent = "0%"
        End If
    Else
        strMessage = "You miss the " & strEnemy
    End If
    If intEnHP < 1 Then
        strMessage = "You kill the " & strEnemy & "!"
        strMessage2 = ""
        lblMessage.ForeColor = &HC000&
        intExp = intExp + (intMaxEnHP + (intEnWeap * 2) + (intEnLevel * 2))
        lblXP.Caption = Int((intExp / intCurLevel) * 100) & "%"
        shpXPPercent.Height = (intExp / intCurLevel) * 100
        If ((intExp / intCurLevel) * 100) >= 100 Then
            lblXP.Caption = "100%"
            shpXPPercent.Height = 100
        End If
        intChance = Rnd * 20
        If intChance <= 5 Then
            intChance = Int(Rnd * ((intEnLevel * 2) + 25)) + 1
            intMoney = intMoney + intChance
            MsgBox "You find $" & intChance & " off this poor soul."
        End If
        intChance = Rnd * 50
        If intChance <= 5 Then
            intChance = Rnd * 100
            If intChance >= 0 And intChance <= 40 Then
                intSpoon = intSpoon + 1
                MsgBox "You find some Spoon Dope off this dead dude"
            ElseIf intChance >= 41 And intChance <= 70 Then
                intLazy = intLazy + 1
                MsgBox "You find some Lazy Dazie off this dead dude"
            ElseIf intChance >= 71 And intChance <= 90 Then
                intKoocie = intKoocie + 1
                MsgBox "You find some Koocie off this dead dude"
            ElseIf intChance >= 91 And intChance <= 100 Then
                intLucky = intLucky + 1
                MsgBox "You find some Lucky Charm off this dead dude"
            End If
        End If
        GoTo PP
    End If
    intHit = Rnd * intEnWeap
    If intHit > 0 Then
        intHP = intHP - intHit
        strMessage2 = "You are hit for " & intHit
    Else
        strMessage2 = "You are missed by the " & strEnemy
    End If
    If intHP < 1 Then
        strMessage = "You are killed!"
        lblMessage.ForeColor = &HC0&
        If strCompany = "LADC" Then
            intLevel = intLevel - 1
            If intLevel <= 0 Then
                MsgBox "You are too low of a level to be salvaged.  You lose the game"
                Unload Me
            End If
            bolBrass = False
            strWeap = "Fists"
            intWeap = 1
            bolGold = False
            bolChain = False
            bolSwitch = False
            bolPistol = False
            bolUzi = False
            bolAssault = False
            bolPunisher = False
            bolW1 = False
            bolW2 = False
            bolW3 = False
            bolW4 = False
            intSpoon = 0
            intLazy = 0
            intKoocie = 0
            intLucky = 0
            intHair = 0
            intCoka = 0
            intCalf = 0
            intMoney = 0
            intAmmo = 0
            intExp = intCurLevel / 2
            intCurLevel = intCurLevel / 2
            strMessage2 = "You restart this rank but lose your items"
            cmdAtt.Caption = "Next Battle!"
        Else
            intLevel = intLevel - 2
            If intLevel <= 0 Then
                MsgBox "You are too low of a level to be salvaged.  You lose the game"
                Unload Me
            End If
            bolBrass = False
            strWeap = "Fists"
            intWeap = 1
            bolGold = False
            bolChain = False
            bolSwitch = False
            bolPistol = False
            bolUzi = False
            bolAssault = False
            bolPunisher = False
            bolW1 = False
            bolW2 = False
            bolW3 = False
            bolW4 = False
            intSpoon = 0
            intSpoon = 0
            intLazy = 0
            intKoocie = 0
            intLucky = 0
            intHair = 0
            intCoka = 0
            intCalf = 0
            intMoney = 0
            intAmmo = 0
            intExp = intCurLevel / 4
            intCurLevel = intCurLevel / 4
            strMessage2 = "You lose one rank and your items"
            cmdAtt.Caption = "Next Battle!"
        End If
    End If
PP:
    lblMessage.Caption = "Messages:" & vbLf & strMessage & vbLf & strMessage2
    lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
    If intEnHP < 1 Then
        cmdAtt.Caption = "Next Battle!"
        If intExp < 500 Then
            intEnemy = Rnd * 4
            If intEnemy = 1 Then
                strEnemy = "Old Man"
                strEnWeap = "Cane"
                intEnWeap = 1
                intEnLevel = 1
                intEnHP = 10
                intMaxEnHP = 10
            ElseIf intEnemy = 2 Then
                strEnemy = "Old Lady"
                strEnWeap = "Walker"
                intEnWeap = 2
                intEnHP = 14
                intEnLevel = 1
                intMaxEnHP = 14
            ElseIf intEnemy = 3 Then
                strEnemy = "Low Life"
                strEnWeap = "Broken Beer Bottle"
                intEnWeap = 3
                intEnHP = 10
                intEnLevel = 1
                intMaxEnHP = 10
            Else
                strEnemy = "School Boy"
                strEnWeap = "Textbook"
                intEnWeap = 2
                intEnLevel = 1
                intEnHP = 7
                intMaxEnHP = 7
            End If
        ElseIf intExp < 1000 Then
            intEnemy = Rnd * 5
            If intEnemy = 1 Then
                strEnemy = "Wannabe"
                strEnWeap = "Chains"
                intEnWeap = 3
                intEnLevel = 2
                intEnHP = 15
                intMaxEnHP = 15
            ElseIf intEnemy = 2 Then
                strEnemy = "Thug"
                strEnWeap = "Silver Knuckles"
                intEnWeap = 3
                intEnHP = 5
                intEnLevel = 2
                intMaxEnHP = 5
            ElseIf intEnemy = 3 Then
                strEnemy = "Beggar"
                strEnWeap = "Rope"
                intEnWeap = 4
                intEnHP = 10
                intEnLevel = 2
                intMaxEnHP = 10
            Else
                strEnemy = "Wimp"
                strEnWeap = "Fists"
                intEnWeap = 1
                intEnHP = 50
                intEnLevel = 2
                intMaxEnHP = 50
            End If
        ElseIf intExp < 2000 Then
            intEnemy = Rnd * 4
            If intEnemy = 1 Then
                strEnemy = "Construction Worker"
                strEnWeap = "Rusted Shovel"
                intEnWeap = 5
                intEnHP = 25
                intEnLevel = 3
                intMaxEnHP = 25
            ElseIf intEnemy = 2 Then
                strEnemy = "Druggie"
                strEnWeap = "Spoon"
                intEnWeap = 4
                intEnLevel = 3
                intEnHP = 15
                intMaxEnHP = 15
            ElseIf intEnemy = 3 Then
                strEnemy = "Druggie"
                strEnWeap = "Spork"
                intEnWeap = 1
                intEnLevel = 3
                intEnHP = 25
                intMaxEnHP = 35
            Else
                strEnemy = "Wise Guy"
                strEnWeap = "Chains"
                intEnWeap = 4
                intEnLevel = 3
                intEnHP = 35
                intMaxEnHP = 35
            End If
        ElseIf intExp < 4000 Then
            intEnemy = Rnd * 4
            intEnLevel = 4
            If intEnemy = 1 Then
                strEnemy = "Smart Cookie"
                strEnWeap = "Pistol"
                intEnWeap = 6
                intEnHP = 30
                intMaxEnHP = 30
            ElseIf intEnemy = 2 Then
                strEnemy = "Dope Master"
                strEnWeap = "Gold Knuckles"
                intEnWeap = 7
                intEnHP = 25
                intMaxEnHP = 25
            ElseIf intEnemy = 3 Then
                strEnemy = "Ranger"
                strEnWeap = "Pistol"
                intEnWeap = 5
                intEnHP = 25
                intMaxEnHP = 25
            Else
                strEnemy = "Police"
                strEnWeap = "Pistol"
                intEnWeap = 6
                intEnHP = 35
                intMaxEnHP = 35
            End If
        ElseIf intExp < 8000 Then
            intEnemy = Rnd * 4
            intEnLevel = 5
            If intEnemy = 1 Then
                strEnemy = "Child"
                strEnWeap = "Fingernails"
                intEnWeap = 10
                intEnHP = 1
                intMaxEnHP = 1
            ElseIf intEnemy = 2 Then
                strEnemy = "Poo-Brain"
                strEnWeap = "Poop"
                intEnWeap = 3
                intEnHP = 50
                intMaxEnHP = 50
            ElseIf intEnemy = 3 Then
                strEnemy = "S.W.A.T. Member"
                strEnWeap = "Glock 9mm"
                intEnWeap = 10
                intEnHP = 25
                intMaxEnHP = 25
            Else
                strEnemy = "Sniffer Dog"
                strEnWeap = "Rabies Infested Teeth"
                intEnWeap = 8
                intEnHP = 25
                intMaxEnHP = 25
            End If
        ElseIf intExp < 16000 Then
            intEnemy = Rnd * 3
            intEnLevel = 6
            If intEnemy = 1 Then
                strEnemy = "Lazy Eyes"
                strEnWeap = "Drug Bomb"
                intEnWeap = 10
                intEnHP = 15
                intMaxEnHP = 15
            ElseIf intEnemy = 2 Then
                strEnemy = "Gangster"
                strEnWeap = "Switchblade"
                intEnWeap = 13
                intEnHP = 20
                intMaxEnHP = 20
            Else
                strEnemy = "Mobster"
                strEnWeap = "Glock 9mm"
                intEnWeap = 15
                intEnHP = 25
                intMaxEnHP = 25
            End If
        ElseIf intExp < 32000 Then
            intEnemy = Rnd * 3
            intEnLevel = 7
            If intEnemy = 1 Then
                strEnemy = "Mafia Boss"
                strEnWeap = "Uzi"
                intEnWeap = 12
                intEnHP = 20
                intMaxEnHP = 20
            ElseIf intEnemy = 2 Then
                strEnemy = "Psycho"
                strEnWeap = "Backpack Bomb"
                intEnWeap = 25
                intEnHP = 1
                intMaxEnHP = 1
            Else
                strEnemy = "Psycho"
                strEnWeap = "Glock 9mm"
                intEnWeap = 15
                intEnHP = 10
                intMaxEnHP = 10
            End If
        ElseIf intExp < 64000 Then
            intEnLevel = 8
            strEnemy = "Drugged Nut"
            strEnWeap = "Assault Rifle"
            intEnWeap = 20
            intEnHP = 50
            intMaxEnHP = 50
        ElseIf intExp < 128000 Then
            intEnLevel = 9
            strEnemy = "Godfather"
            strEnWeap = "Punisher .77mm"
            intEnWeap = 25
            intEnHP = 110
            intMaxEnHP = 110
        Else
            intEnLevel = 10
            strEnemy = "Mega Man"
            strEnWeap = "Blaster'thingy of Doom'ness"
            intEnWeap = 50
            intEnHP = 10000
            intMaxEnHP = 10000
        End If
    fraEnemy.Caption = strEnemy
    lblEnemy.Caption = "You are attacked by a level " & intEnLevel & " " & strEnemy
    lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
    lblMessage.ForeColor = &HC00000
    End If
    If intExp >= 500 And intLevel = 1 Then
        intLevel = intLevel + 1
        MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a School Boy"
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 10
        intCurLevel = intL3
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 1000 And intLevel = 2 Then
        intLevel = intLevel + 1
        MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Wannabe"
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 15
        intCurLevel = intL4
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 2000 And intLevel = 3 Then
        intLevel = intLevel + 1
        MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Thug"
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 20
        intCurLevel = intL5
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 4000 And intLevel = 4 Then
        intLevel = intLevel + 1
        If intDrugUse >= 5 Or bolDrugLine = True Then
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Druggie"
        Else
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Wise Guy"
        End If
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 25
        intCurLevel = intL6
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 8000 And intLevel = 5 Then
        intLevel = intLevel + 1
        If intDrugUse >= 10 Or bolDrugLine = True Then
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Dope Master"
        Else
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Smart Cookie"
        End If
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 30
        intCurLevel = intL7
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 16000 And intLevel = 6 Then
        intLevel = intLevel + 1
        If intDrugUse >= 15 Or bolDrugLine = True Then
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Lazy Eyes"
        Else
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Gangster"
        End If
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 35
        intCurLevel = intL8
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 32000 And intLevel = 7 Then
        intLevel = intLevel + 1
        If bolDrugLine = True Then
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Psycho"
        Else
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Mobster"
        End If
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 40
        intCurLevel = intL9
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 64000 And intLevel = 8 Then
        intLevel = intLevel + 1
        If bolDrugLine = True Then
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Druggie"
        Else
            MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "You are now a Wise Guy"
        End If
        intHeal = intHeal + 2
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 45
        intCurLevel = intL10
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    ElseIf intExp >= 128000 And intLevel = 9 Then
        intLevel = intLevel + 1
        MsgBox "You level up!  You are now at level " & intLevel & ", " & strName & vbLf & vbLf & "This it the highest level that you can get in this game.  Soon, you can goto level 50!  But untill then, keep battling MEGA MAN! (You also get 50 extra Medsticks for getting this far)"
        intHeal = intHeal + 50
        intHP = intMaxHP
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        intSet = intSet + 50
        intCurLevel = intL10
        intExp = 0
        lblXP.Caption = "0%"
        shpXPPercent.Height = 0
    End If
    lblName.Caption = strName & ", Level " & intLevel
    lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
    lblCash.Caption = "Cash: $" & vbLf & intMoney
    lblEnemy.Caption = "You are attacked by a level " & intEnLevel & " " & strEnemy
    lblMonth.Caption = "Month:" & vbLf & strMonth
    fraStats.Caption = strName
    fraEnemy.Caption = strEnemy
    lblXP.Caption = Int((intExp / intCurLevel) * 100) & "%"
    shpXPPercent.Height = (intExp / intCurLevel) * 100
    shpTrain.Height = intTrain
    lblPercent = intTrain & "%"
End Sub

Private Sub cmdBar_Click()
    frmBar.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdCasino_Click()
    frmCasino.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdDrugs_Click()
    frmSnack.Show
    frmRPG.Visible = False
    frmRPG.Enabled = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdFlee_Click()
    If intHeal > 0 Then
        intHP = intHP + 10
        If intHP > intMaxHP Then
            intHP = intMaxHP
        End If
        intHeal = intHeal - 1
        cmdFlee.Caption = "Medstick (" & intHeal & ")"
        strMessage = "--You heal yourself for 10 points--"
    Else
        strMessage = "--You have no more Heal Potions--"
        strMessage2 = "--You must attack--"
    End If
    lblMessage.Caption = "Messages:" & vbLf & strMessage & vbLf & strMessage2
    lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
End Sub

Private Sub cmdGarage_Click()
    frmGarage.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdGym_Click()
    If intMoney >= 100 Then
        intTrain = intTrain + 10
        intMoney = intMoney - 100
    Else
        MsgBox "You do not have $100 to train."
    End If
    lblCash.Caption = "Cash: $" & vbLf & intMoney
    shpTrain.Height = intTrain
    lblPercent = intTrain & "%"
    If intTrain >= 100 Then
        intSpec = intSpec + 1
        MsgBox "You learn a new move training!"
        lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
        intTrain = 0
        shpTrain.Height = 0
        lblPercent = "0%"
    End If
End Sub

Private Sub cmdHire_Click()
    frmLacky.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdHotel_Click()
    If intHP < intMaxHP And intMoney >= 50 Then
        intHP = intMaxHP
        intMoney = intMoney - 50
    ElseIf intMoney < 50 Then
        MsgBox "You do not have $50 to heal"
    Else
        MsgBox "You're fully healed, you don't need to rest"
    End If
    lblCash.Caption = "Cash: $" & vbLf & intMoney
    lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
End Sub

Private Sub cmdHQ_Click()
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
    frmHQ.Show
End Sub

Private Sub cmdInventory_Click()
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
    frmInventory.Show
End Sub

Private Sub cmdLoan_Click()
    If bolLoanShark = True Then
        frmLoan.Show
        frmRPG.Enabled = False
        frmRPG.Visible = False
        frmMove.Visible = False
        frmCityMap.Visible = False
    Else
        MsgBox "The Loan Shark isn't in yet.  He will be in but a few months"
    End If
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
End Sub

Private Sub cmdQuit_Click()
    If intHP > 0 Then
        intSelect = MsgBox("Are you sure you wish to quit?", vbYesNo, "Quit")
        If intSelect = vbYes Then
            Unload frmMove
            Unload Me
        End If
    Else
        Unload frmCityMap
        Unload frmMove
        Unload Me
    End If
End Sub

Private Sub cmdStock_Click()
    frmStock.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdStore_Click()
    frmGeneral.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub cmdTattoo_Click()
    frmTattoo.Show
    frmRPG.Enabled = False
    frmRPG.Visible = False
    frmMove.Visible = False
    frmCityMap.Visible = False
End Sub

Private Sub Command1_Click()
    frmMove.Show
End Sub

Private Sub Dis_Timer()
    If strSpeed = "P" Then
        If Display.Caption = "||" Then
            Display.Caption = " "
        Else
            Display.Caption = "||"
        End If
    ElseIf strSpeed = "1" Then
        Display.Visible = True
        Display.Caption = ">"
    ElseIf strSpeed = "2" Then
        Display.Visible = "True"
        If Display.Caption = ">" Then
            Display.Caption = ">>"
        ElseIf Display.Caption = ">>" Then
            Display.Caption = " "
        ElseIf Display.Caption = " " Then
            Display.Caption = ">"
        End If
    ElseIf strSpeed = "3" Then
        Display.Visible = "True"
        If Display.Caption = ">" Then
            Display.Caption = ">>"
        ElseIf Display.Caption = ">>" Then
            Display.Caption = ">>>"
        ElseIf Display.Caption = ">>>" Then
            Display.Caption = " "
        ElseIf Display.Caption = " " Then
            Display.Caption = ">"
        End If
    End If
End Sub

Private Sub Form_Load()
    Randomize
    If intLoadTemp <> 10 Then
        intMonth = 1
        lblMonth.Caption = "Month:" & vbLf & "Jan."
        If strName = "" Or strName = "Enter Name Here" Then
            strName = "Player"
        End If
        strMonth = "Jan."
        intLevel = 1
        intEnLevel = 1
        strWeap = "Brass Knuckles"
        intWeap = 2
        strEnWeap = "Cane"
        intEnWeap = 1
        intHP = 50
        X = 1
        Y = 1
        CityMap = False
        bolBar = True
        bolBrass = True
        intMaxHP = 50
        intEnHP = 10
        intMaxEnHP = 10
        intHeal = 15
        intCurLevel = 500
        strEnemy = "Old Man"
        lblName.Caption = strName & ", Level " & intLevel
        lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
        lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
        intHPIncrese = 5
        lblEnemy.Caption = "You are attacked by a level " & intEnLevel & " " & strEnemy
        If strName = "All your cheats are belong to us!" Then
            strName = InputBox("Please enter a name:")
            intLevel = 1
            strWeap = "Punisher .77mm"
            intWeap = 10
            intHP = 500
            intMoney = 1000000000
            intMaxHP = 500
            intSpec = 5
            intHeal = 500
            bolPunisher = True
            intSpoon = 100
            intSpoon = 100
            intLazy = 100
            intKoocie = 100
            intLucky = 100
            intHair = 100
            intCoka = 100
            intCalf = 10
            lblName.Caption = strName & ", Level " & intLevel
            lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
            lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
            intHPIncrese = 5
            MsgBox "Liar!  All YOUR cheats belong to ME!"
        End If
        'Declare Level Variables
        intL2 = 500
        intL3 = 1000
        intL4 = 2000
        intL5 = 4000
        intL6 = 8000
        intL7 = 16000
        intL8 = 32000
        intL9 = 64000
        intL10 = 128000
        lblCash.Caption = "Cash: $" & vbLf & intMoney
        totSCE = 10000
        totSHE = 7500
        totBI = 5000
        totLADC = 2500
        totTWM = 1000
        totCNS = 1000
        totBA = 1000
        totHC = 1000
        priSCE = 1000
        priSHE = 750
        priBI = 500
        priLADC = 500
        priTWM = 100
        priCNS = 100
        priBA = 100
        priHC = 100
        fraStats.Caption = strName
        fraEnemy.Caption = strEnemy
    End If
    lblName.Caption = strName & ", Level " & intLevel
    lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
    lblCash.Caption = "Cash: $" & vbLf & intMoney
    lblEnemy.Caption = "You are attacked by a level " & intEnLevel & " " & strEnemy
    lblMonth.Caption = "Month:" & vbLf & strMonth
    fraStats.Caption = strName
    fraEnemy.Caption = strEnemy
    lblXP.Caption = Int((intExp / intCurLevel) * 100) & "%"
    shpXPPercent.Height = (intExp / intCurLevel) * 100
    shpTrain.Height = intTrain
    lblPercent = intTrain & "%"
    frmMove.Show
    frmCityMap.Show
    If CityMap = True Then
        frmCityMap.Visible = True
    Else
        frmCityMap.Visible = False
    End If
End Sub

Private Sub Messages_Click()
    CityMap = True
    If CityMap = False Then
        MsgBox "You don't own a map!"
    Else
        If frmCityMap.Visible = True Then
            frmCityMap.Visible = False
        Else
            frmCityMap.Visible = True
            frmCityMap.Label4.Caption = "(" & X & "," & Y & ")"
        End If
    End If
End Sub

Private Sub MonthTime_Timer()
    If MonthBar.Height < 87 Then
        MonthBar.Height = MonthBar.Height + 3
    Else
        MonthBar.Height = 0
    End If
End Sub

Private Sub Pause_Click()
    Bar.Enabled = False
    frmBar.Bar.Enabled = False
    Time.Enabled = False
    cmdAtt.Enabled = False
    cmdFlee.Enabled = False
    cmdGym.Enabled = False
    cmdHotel.Enabled = False
    cmdBar.Enabled = False
    cmdDrugs.Enabled = False
    cmdCasino.Enabled = False
    cmdLoan.Enabled = False
    cmdTattoo.Enabled = False
    cmdStock.Enabled = False
    cmdArena.Enabled = False
    cmdGarage.Enabled = False
    cmdStore.Enabled = False
    cmdArmory.Enabled = False
    cmdHQ.Enabled = False
    cmdInventory.Enabled = False
    cmdHire.Enabled = False
    Messages.Enabled = False
    frmMove.Command1.Enabled = False
    frmMove.Command2.Enabled = False
    frmMove.Command3.Enabled = False
    frmMove.Command4.Enabled = False
    Pause.BackColor = &HE0E0E0
    Turbo1.BackColor = &H8000000F
    Turbo2.BackColor = &H8000000F
    Turbo3.BackColor = &H8000000F
    Display.ForeColor = &HC0C0C0
    Display.Caption = "||"
    Dis.Enabled = True
    MonthTime.Enabled = False
    strSpeed = "P"
End Sub

Private Sub Time_Timer()
    CurrentMonthTime = CurrentMonthTime + 1
    If CurrentMonthTime = 30 Then
        intTime = intTime + 1
        If intMonth < 12 Then
            intMonth = intMonth + 1
        Else
            intMonth = 1
        End If
        If intMonth = 1 Then
            strMonth = "Jan."
        ElseIf intMonth = 2 Then
            strMonth = "Feb."
        ElseIf intMonth = 3 Then
            strMonth = "Mar."
        ElseIf intMonth = 4 Then
            strMonth = "Apr."
        ElseIf intMonth = 5 Then
            strMonth = "May"
        ElseIf intMonth = 6 Then
            strMonth = "June"
        ElseIf intMonth = 7 Then
            strMonth = "July"
        ElseIf intMonth = 8 Then
            strMonth = "Aug."
        ElseIf intMonth = 9 Then
            strMonth = "Sep."
        ElseIf intMonth = 10 Then
            strMonth = "Oct."
        ElseIf intMonth = 11 Then
            strMonth = "Nov."
        ElseIf intMonth = 12 Then
            strMonth = "Dec."
        End If
        lblMonth.Caption = "Month:" & vbLf & strMonth
        If bolAIDS = True Then
            intMaxHP = intMaxHP - 10
            intHP = intHP - 10
        End If
        lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
        If intTime = 6 Then
            bolLoanShark = True
            lblNews.Caption = "Jack the Loan Shark has arrived!"
        ElseIf intTime = 20 Then
            bolCalfAvailable = True
            lblNews.Caption = "The cure to AIDS has been invented!"
        Else
            lblNews.Caption = "No new news this month"
        End If
        If strCompany = "HC" Then
            CurrentLoan = Int((CurrentLoan * 0.2) + CurrentLoan)
        Else
            CurrentLoan = Int((CurrentLoan * 0.35) + CurrentLoan)
        End If
        If CurrentLoan > 0 Then
            intLoanTotalTime = intLoanTotalTime - 1
            If intLoanTotalTime = 1 Then
                MsgBox "HEY!  You gotta pay your loan in a month!"
            ElseIf intLoanTotalTime = 0 Then
                intChance = Rnd * 10
                If strCompany = "HC" Then
                    If intChance <= 9 Then
                        MsgBox "You didn't pay your loan... You're dead."
                        Unload Me
                    Else
                        MsgBox "Hmm... This loan's erased!"
                        CurrentLoan = 0
                        StartLoan = 0
                    End If
                Else
                    If intChance <= 8 Then
                        MsgBox "You didn't pay your loan... You're dead."
                        Unload Me
                    Else
                        MsgBox "Hmm... This loan's erased!"
                        CurrentLoan = 0
                        StartLoan = 0
                    End If
                End If
            End If
        End If
        
        intProsEarn = Int(Rnd * (intPros * 30))
        intChemEarn = Int(Rnd * (intChem * 40))
        intBounceEarn = Int(Rnd * (intBounce * 60))
        If intPros = 0 Then
            lblPros.Caption = "You have no Prostitutes!"
        ElseIf intProsEarn = 0 Then
            lblPros.Caption = "Your Prostitues earned nothing!"
        Else
            lblPros.Caption = "Your Prostitues earned $" & intProsEarn
            intMoney = intMoney + intProsEarn
        End If
        
        If intChem = 0 Then
            lblChem.Caption = "You have no Chem'ys!"
        ElseIf intChemEarn = 0 Then
            lblChem.Caption = "Your Chem'ys sold nothing!"
        Else
            lblChem.Caption = "Your Chem'ys earned $" & intChemEarn
            intMoney = intMoney + intChemEarn
        End If
        
        If intBounce = 0 Then
            lblBounce.Caption = "You have no Bouncers!"
        ElseIf intBounceEarn = 0 Then
            lblBounce.Caption = "Your bouncers killed rats for fun..."
        Else
            lblBounce.Caption = "Your bouncers pilfered $" & intBounceEarn & " from enemies!"
            intMoney = intMoney + intBounceEarn
        End If
        
        'Stock Prices
        intChance = Rnd * 2
        If intChance = 1 Then
            bolSCE = True
            priSCE = priSCE + (Int(Rnd * 100))
        Else
            bolSCE = False
            priSCE = priSCE - (Int(Rnd * 100))
        End If
        
        intChance = Rnd * 2
        If intChance = 1 Then
            bolTWM = True
            priTWM = priTWM + (Int(Rnd * 10))
        Else
            bolTWM = False
            priTWM = priTWM - (Int(Rnd * 10))
        End If
        
        intChance = Rnd * 2
        If intChance = 1 Then
            bolSHE = True
            priSHE = priSHE + (Int(Rnd * 75))
        Else
            bolSHE = False
            priSHE = priSHE - (Int(Rnd * 75))
        End If
        
        intChance = Rnd * 2
        If intChance = 1 Then
            bolCNS = True
            priCNS = priCNS + (Int(Rnd * 10))
        Else
            bolCNS = False
            priCNS = priCNS - (Int(Rnd * 10))
        End If
        
        intChance = Rnd * 2
        If intChance = 1 Then
            bolBI = True
            priBI = priBI + (Int(Rnd * 50))
        Else
            bolBI = False
            priBI = priBI - (Int(Rnd * 50))
        End If
        
        intChance = Rnd * 2
        If intChance = 1 Then
            bolBA = True
            priBA = priBA + (Int(Rnd * 10))
        Else
            bolBA = False
            priBA = priBA - (Int(Rnd * 10))
        End If
        
        intChance = Rnd * 2
        If intChance = 1 Then
            bolLADC = True
            priLADC = priLADC + (Int(Rnd * 50))
        Else
            bolLADC = False
            priLADC = priLADC - (Int(Rnd * 50))
        End If
        
        intChance = Rnd * 2
        If intChance = 1 Then
            bolHC = True
            priHC = priHC + (Int(Rnd * 10))
        Else
            bolHC = False
            priHC = priHC - (Int(Rnd * 10))
        End If
        
        'Bank Interest
        intBank = intBank + (Int(intBank * 0.5))
        
        'Company Earnings
        dbl1 = dbl1 + Int(Rnd * 100000)
        dbl2 = dbl2 + Int(Rnd * 75000)
        dbl3 = dbl3 + Int(Rnd * 50000)
        dbl4 = dbl4 + Int(Rnd * 50000)
        dbl5 = dbl5 + Int(Rnd * 10000)
        dbl6 = dbl6 + Int(Rnd * 10000)
        dbl7 = dbl7 + Int(Rnd * 10000)
        dbl8 = dbl8 + Int(Rnd * 10000)
        
        CurrentMonthTime = 0
    End If
    'This must be at bottom
    lblName.Caption = strName & ", Level " & intLevel
    lblStats.Caption = "Weapon: " & strWeap & vbLf & "     (Att: " & intWeap & "+" & intSpec & ")" & vbLf & vbLf & "HP: " & intHP & " / " & intMaxHP & vbLf & vbLf & "EXP: " & intExp
    lblEnemyStats.Caption = "Weapon: " & strEnWeap & vbLf & "     (Att: " & intEnWeap & ")" & vbLf & vbLf & "HP: " & intEnHP & " / " & intMaxEnHP
    lblCash.Caption = "Cash: $" & vbLf & intMoney
    lblEnemy.Caption = "You are attacked by a level " & intEnLevel & " " & strEnemy
    lblMonth.Caption = "Month:" & vbLf & strMonth
    fraStats.Caption = strName
    fraEnemy.Caption = strEnemy
    shpTrain.Height = intTrain
    lblPercent = intTrain & "%"
End Sub

Private Sub Turbo1_Click()
    Bar.Enabled = True
    Bar.Interval = 10000
    frmBar.Bar.Enabled = True
    frmBar.Bar.Interval = 10000
    Time.Enabled = True
    Time.Interval = 1000
    cmdAtt.Enabled = True
    cmdFlee.Enabled = True
    cmdGym.Enabled = True
    cmdHotel.Enabled = True
    cmdBar.Enabled = True
    cmdDrugs.Enabled = True
    cmdCasino.Enabled = True
    cmdLoan.Enabled = True
    cmdTattoo.Enabled = True
    cmdStock.Enabled = True
    cmdArena.Enabled = True
    cmdGarage.Enabled = True
    cmdStore.Enabled = True
    cmdArmory.Enabled = True
    cmdHQ.Enabled = True
    cmdInventory.Enabled = True
    cmdHire.Enabled = True
    Messages.Enabled = True
    frmMove.Command1.Enabled = True
    frmMove.Command2.Enabled = True
    frmMove.Command3.Enabled = True
    frmMove.Command4.Enabled = True
    Pause.BackColor = &H8000000F
    Turbo1.BackColor = &HE0E0E0
    Turbo2.BackColor = &H8000000F
    Turbo3.BackColor = &H8000000F
    Display.ForeColor = &HFF00&
    Display.Caption = ">"
    Dis.Enabled = False
    Display.Visible = True
    MonthTime.Enabled = True
    MonthTime.Interval = 1000
    strSpeed = "1"
End Sub

Private Sub Turbo2_Click()
    Bar.Enabled = True
    Bar.Interval = 5000
    frmBar.Bar.Enabled = True
    frmBar.Bar.Interval = 5000
    Time.Enabled = True
    Time.Interval = 500
    cmdAtt.Enabled = True
    cmdFlee.Enabled = True
    cmdGym.Enabled = True
    cmdHotel.Enabled = True
    cmdBar.Enabled = True
    cmdDrugs.Enabled = True
    cmdCasino.Enabled = True
    cmdLoan.Enabled = True
    cmdTattoo.Enabled = True
    cmdStock.Enabled = True
    cmdArena.Enabled = True
    cmdGarage.Enabled = True
    cmdStore.Enabled = True
    cmdArmory.Enabled = True
    cmdHQ.Enabled = True
    cmdInventory.Enabled = True
    cmdHire.Enabled = True
    Messages.Enabled = True
    frmMove.Command1.Enabled = True
    frmMove.Command2.Enabled = True
    frmMove.Command3.Enabled = True
    frmMove.Command4.Enabled = True
    Pause.BackColor = &H8000000F
    Turbo1.BackColor = &H8000000F
    Turbo2.BackColor = &HE0E0E0
    Turbo3.BackColor = &H8000000F
    Display.ForeColor = &HFFFF&
    Display.Caption = ">"
    Dis.Enabled = True
    MonthTime.Enabled = True
    MonthTime.Interval = 500
    strSpeed = "2"
End Sub

Private Sub Turbo3_Click()
    Bar.Enabled = True
    Bar.Interval = 2500
    frmBar.Bar.Enabled = True
    frmBar.Bar.Interval = 2500
    Time.Enabled = True
    Time.Interval = 250
    cmdAtt.Enabled = True
    cmdFlee.Enabled = True
    cmdGym.Enabled = True
    cmdHotel.Enabled = True
    cmdBar.Enabled = True
    cmdDrugs.Enabled = True
    cmdCasino.Enabled = True
    cmdLoan.Enabled = True
    cmdTattoo.Enabled = True
    cmdStock.Enabled = True
    cmdArena.Enabled = True
    cmdGarage.Enabled = True
    cmdStore.Enabled = True
    cmdArmory.Enabled = True
    cmdHQ.Enabled = True
    cmdInventory.Enabled = True
    cmdHire.Enabled = True
    Messages.Enabled = True
    frmMove.Command1.Enabled = True
    frmMove.Command2.Enabled = True
    frmMove.Command3.Enabled = True
    frmMove.Command4.Enabled = True
    Pause.BackColor = &H8000000F
    Turbo1.BackColor = &H8000000F
    Turbo2.BackColor = &H8000000F
    Turbo3.BackColor = &HE0E0E0
    Display.ForeColor = &HFF&
    Display.Caption = ">"
    Dis.Enabled = True
    MonthTime.Enabled = True
    MonthTime.Interval = 250
    strSpeed = "3"
End Sub
