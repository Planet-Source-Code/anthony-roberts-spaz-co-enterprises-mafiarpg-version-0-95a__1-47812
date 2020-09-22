VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Save da freakin' game!"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "XP Bar"
      Height          =   1095
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   855
      Begin VB.OptionButton Option6 
         Caption         =   "Option6"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option5"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option4"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Train"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   855
      Begin VB.OptionButton Option3 
         Caption         =   "Option3"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   720
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   480
         TabIndex        =   8
         Top             =   480
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Return to Game"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set Options"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1200
      Top             =   1560
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      FillColor       =   &H0000C0C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1200
      Top             =   1320
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1200
      Top             =   1080
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   3240
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Change Bar Colours:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label1 
      Caption         =   "New Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Shape Shape4 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   10
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If txtName.Text <> "" Then
        strName = txtName.Text
    End If
    If Option1.Value = True Then
        frmRPG.shpTrain.FillColor = &HC0&
    ElseIf Option2.Value = True Then
        frmRPG.shpTrain.FillColor = &HC0C0&
    ElseIf Option3.Value = True Then
        frmRPG.shpTrain.FillColor = &H8000&
    End If
    If Option4.Value = True Then
        frmRPG.shpXPPercent.FillColor = &HC0&
    ElseIf Option5.Value = True Then
        frmRPG.shpXPPercent.FillColor = &HC0C0&
    ElseIf Option6.Value = True Then
        frmRPG.shpXPPercent.FillColor = &H8000&
    End If
End Sub

Private Sub Command2_Click()
    frmRPG.lblName.Caption = strName & ", Level " & intLevel
    frmRPG.Enabled = True
    frmRPG.Visible = True
    frmMove.Visible = True
    Unload frmOptions
End Sub

Private Sub Command3_Click()
    strPath = App.Path & "\save\"
    strFileName = InputBox("Please enter a file name for your save game: " & vbLf & "(Note:  Remember this name for when you load it)", "Save Game")
    If strFileName = Empty Then
        MsgBox "You have to enter a name!", vbCritical, "Error"
        Exit Sub
    End If
    strExtension = ".sms"
    strFileNameAndPath = strPath & strFileName & strExtension
    Open strFileNameAndPath For Output As #1
        Write #1, strName, intHP, intEnHP, intAtt, intEnAtt, intWeap, _
                    intEnWeap, intHit, intLevel, intSpec, strWeap, strEnWeap, _
                    intEnLevel, intExp, intMaxHP, intMaxEnHP, strMessage, strEnemy, _
                    strMessage2, intEnemy, intHeal, intSet, intHPIncrese, intTrain, _
                    intSelect, intL2, intL3, intL4, intL5, intL6, intL7, _
                    intL8, intL9, intL10, intCurLevel, intMoney, bolAIDS, intDrugUse, _
                    bolDrugLine, intTime, strMonth, intLoanTotalTime, intMonth, intAmmo, _
                    StartLoan, CurrentLoan, intStealing, intChance, intBorder, bolLoanShark, _
                    bolCalfAvailable, bolBreast, bolPussy, bolSlap, bolWeap, bolDrink, intBangCost, _
                    intGetMax, intCurrentGet, intHPIncreseChick, bolBar, _
                    strCompany, intSpoon, intLazy, intKoocie, _
                    intLucky, intHair, intCoka, intCalf, bolBrass, bolGold, bolChain, _
                    bolSwitch, bolPistol, bolUzi, bolAssault, bolPunisher, _
                    intBet, intNum1, intNum2, intNum3, intSCE, intSHE, intBI, intLADC, _
                    intTWM, intCNS, intBA, intHC, SCE, SHE, BI, LADC, _
                    TWM, CNS, BA, HC, totSCE, totSHE, totBI, totLADC, _
                    totTWM, totCNS, totBA, totHC, intProsEarn, intChemEarn, intBounceEarn, _
                    intPros, intChem, intBounce, intBank, bolBank, bolCondom, _
                    bolTaxi, bolFirefly, bolMustang, bolViper, bolGTO, dbl1, dbl2, dbl3, dbl4, dbl5, dbl6, dbl7, dbl8
                    
    Close #1
    MsgBox "Game Saved Successfully!"
End Sub
