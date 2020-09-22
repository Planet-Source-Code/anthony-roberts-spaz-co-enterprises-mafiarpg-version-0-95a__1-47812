VERSION 5.00
Begin VB.Form frmStart 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   ScaleHeight     =   222
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Quit Game"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   3135
   End
   Begin VB.FileListBox lstFiles 
      Height          =   1065
      Left            =   120
      Pattern         =   "*.sms"
      TabIndex        =   5
      Top             =   3360
      Width           =   4815
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Game, ya pipsqueek!"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lets do this thang..."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "MafiaStart.frx":0000
      Left            =   120
      List            =   "MafiaStart.frx":0010
      TabIndex        =   1
      Text            =   "-- Select Your Starting Company --"
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox txtName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Enter Name Here"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Company Description:"
      ForeColor       =   &H0000C0C0&
      Height          =   3135
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"MafiaStart.frx":0054
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00008000&
      BorderWidth     =   5
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLoad_Click()
    lstFiles.FileName = App.Path & "\save\*.sms"
    strPath = App.Path & "\save\"
    If cmdLoad.Caption = "Load Game, ya pipsqueek!" Then
        frmStart.Height = 4500
        cmdLoad.Caption = "Load dis Game, boss!"
    Else
        strFileName = lstFiles.FileName
        If strFileName = Empty Then
            MsgBox "You have to enter a name!", vbCritical
            Exit Sub
        End If
        strExtension = ".sms"
        strFileNameAndPath = strPath & strFileName & strExtension
        Open App.Path & "\save\" & lstFiles.FileName For Input As #1
            Input #1, strName, intHP, intEnHP, intAtt, intEnAtt, intWeap, _
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
        frmRPG.Show
        MsgBox "Game Loaded Successfully, " & strName & "!"
        intLoadTemp = 10
        Unload frmStart
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1.Text = "Hutch In Corp." Then
        lblCompany.Caption = "Company Description:" & vbLf & vbLf & "Hutch In Corp." & vbLf & vbLf & "Loans have 20% Interest instead of 35%" & vbLf & vbLf & "Chance of Loan Shark killing you is 10% Higher is debt not paid in time."
    ElseIf Combo1.Text = "Clit N' Sons" Then
        lblCompany.Caption = "Company Description:" & vbLf & vbLf & "Clit N' Sons" & vbLf & vbLf & "Less chance of AIDS" & vbLf & vbLf & "AIDS cure costs twice as much."
    ElseIf Combo1.Text = "The White Market" Then
        lblCompany.Caption = "Company Description:" & vbLf & vbLf & "The White Market" & vbLf & vbLf & "Weapons are half off." & vbLf & vbLf & "No permits for guns allowed."
    ElseIf Combo1.Text = "Blazing Auto's" Then
        lblCompany.Caption = "Company Description:" & vbLf & vbLf & "Blazing Auto's" & vbLf & vbLf & "Cars are cheaper to purchase" & vbLf & vbLf & "GTO's can't be purchased or sold."
    End If
End Sub

Private Sub Command1_Click()
    strName = txtName.Text
    If Combo1.Text = "Hutch In Corp." Then
        strCompany = "HC"
        intChance = 1
    ElseIf Combo1.Text = "Clit N' Sons" Then
        strCompany = "CNS"
        intChance = 1
    ElseIf Combo1.Text = "The White Market" Then
        strCompany = "TWM"
        intChance = 1
    ElseIf Combo1.Text = "Blazing Auto's" Then
        strCompany = "BA"
        intChance = 1
    Else
        MsgBox "You never entered a company.  Please select one."
        intChance = 0
    End If
    If intChance = 1 Then
        frmRPG.Show
        frmRPG.Bar.Enabled = True
        frmRPG.Time.Enabled = True
        Unload frmStart
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    MsgBox "WARNING:" & vbLf & vbLf & "This game is reccommended only for mature audiences of 18 years or older." & vbLf & vbLf & "Rated: M for Mature due to sexual content and light textural violence"
    lblCompany.Caption = "Company Description:" & vbLf & vbLf & "No Company Selected"
End Sub

Private Sub txtName_Click()
    If txtName.Text = "Enter Name Here" Then
        txtName.Text = ""
    End If
End Sub
