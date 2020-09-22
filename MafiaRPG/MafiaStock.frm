VERSION 5.00
Begin VB.Form frmStock 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLeave 
      Caption         =   "Leave"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdSell 
      Caption         =   "Sell 10"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuy 
      Caption         =   "Buy 10"
      Height          =   495
      Left            =   1560
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton BA 
      Caption         =   "BA"
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton HC 
      Caption         =   "HC"
      Height          =   495
      Left            =   2760
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton CNS 
      Caption         =   "CNS"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton TWM 
      Caption         =   "TWM"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton LADC 
      Caption         =   "LADC"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton BI 
      Caption         =   "BI"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton SHE 
      Caption         =   "SHE"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton SCE 
      Caption         =   "SCE"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblCostStock 
      Alignment       =   2  'Center
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Line Line16 
      X1              =   5640
      X2              =   5640
      Y1              =   600
      Y2              =   840
   End
   Begin VB.Line Line15 
      X1              =   1800
      X2              =   1800
      Y1              =   840
      Y2              =   600
   End
   Begin VB.Line Line14 
      X1              =   4680
      X2              =   4680
      Y1              =   2880
      Y2              =   2760
   End
   Begin VB.Line Line13 
      X1              =   5640
      X2              =   4680
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line12 
      X1              =   5640
      X2              =   5640
      Y1              =   2760
      Y2              =   3840
   End
   Begin VB.Line Line11 
      X1              =   3960
      X2              =   3960
      Y1              =   2280
      Y2              =   2040
   End
   Begin VB.Line Line10 
      X1              =   2640
      X2              =   2640
      Y1              =   2280
      Y2              =   2040
   End
   Begin VB.Line Line9 
      X1              =   3960
      X2              =   4080
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line8 
      X1              =   2640
      X2              =   2760
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line7 
      X1              =   1320
      X2              =   1440
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line6 
      X1              =   720
      X2              =   720
      Y1              =   3240
      Y2              =   3120
   End
   Begin VB.Line Line5 
      X1              =   720
      X2              =   720
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line4 
      X1              =   720
      X2              =   720
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   720
      X2              =   720
      Y1              =   1320
      Y2              =   1440
   End
   Begin VB.Label lblStats 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   1680
      TabIndex        =   9
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   720
      X2              =   1680
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   720
      Y1              =   840
      Y2              =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000080&
      BorderWidth     =   5
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BA_Click()
    lblName.Caption = "Blazing Auto's"
    lblCostStock.Caption = "Stock Price: $" & priBA
    lblStats.Caption = "You currently own " & intBA & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intBA / totBA) * 100)) & "% of the company"
    If bolBA = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub

Private Sub BI_Click()
    lblName.Caption = "Bones Inc."
    lblCostStock.Caption = "Stock Price: $" & priBI
    lblStats.Caption = "You currently own " & intBI & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intBI / totBI) * 100)) & "% of the company"
    If bolBI = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub

Private Sub cmdBuy_Click()
    If lblName.Caption = "Spaz Co. Enterprises" Then
        If intMoney >= (priSCE * 10) Then
            If intSCE < totSCE Then
                intMoney = intMoney - (priSCE * 10)
                priSCE = priSCE + 1
                intSCE = intSCE + 10
                lblCostStock.Caption = "Stock Price: $" & priSCE
                lblStats.Caption = "You currently own " & intSCE & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intSCE / totSCE) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf lblName.Caption = "The White Market" Then
        If intMoney >= (priTWM * 10) Then
            If intTWM < totTWM Then
                intMoney = intMoney - (priTWM * 10)
                priTWM = priTWM + 1
                intTWM = intTWM + 10
                lblCostStock.Caption = "Stock Price: $" & priTWM
                lblStats.Caption = "You currently own " & intTWM & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intTWM / totTWM) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf lblName.Caption = "Shadow Hentai Enterprises" Then
        If intMoney >= (priSHE * 10) Then
            If intSHE < totSHE Then
                intMoney = intMoney - (priSHE * 10)
                priSHE = priSHE + 1
                intSHE = intSHE + 10
                lblCostStock.Caption = "Stock Price: $" & priSHE
                lblStats.Caption = "You currently own " & intSHE & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intSHE / totSHE) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf lblName.Caption = "Clit N' Sons" Then
        If intMoney >= (priCNS * 10) Then
            If intCNS < totCNS Then
                intMoney = intMoney - (priCNS * 10)
                priCNS = priCNS + 1
                intCNS = intCNS + 10
                lblCostStock.Caption = "Stock Price: $" & priCNS
                lblStats.Caption = "You currently own " & intCNS & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intCNS / totCNS) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf lblName.Caption = "Bones Inc." Then
        If intMoney >= (priBI * 10) Then
            If intBI < totBI Then
                intMoney = intMoney - (priBI * 10)
                priBI = priBI + 1
                intBI = intBI + 10
                lblCostStock.Caption = "Stock Price: $" & priBI
                lblStats.Caption = "You currently own " & intBI & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intBI / totBI) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf lblName.Caption = "Blazing Auto's" Then
        If intMoney >= (priBA * 10) Then
            If intBA < totBA Then
                intMoney = intMoney - (priBA * 10)
                priBA = priBA + 1
                intBA = intBA + 10
                lblCostStock.Caption = "Stock Price: $" & priBA
                lblStats.Caption = "You currently own " & intBA & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intBA / totBA) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf lblName.Caption = "Life After Death Co." Then
        If intMoney >= (priLADC * 10) Then
            If intLADC < totLADC Then
                intMoney = intMoney - (priLADC * 10)
                priLADC = priLADC + 1
                intLADC = intLADC + 10
                lblCostStock.Caption = "Stock Price: $" & priLADC
                lblStats.Caption = "You currently own " & intLADC & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intLADC / totLADC) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    ElseIf lblName.Caption = "Hutch In Corp" Then
        If intMoney >= (priHC * 10) Then
            If intHC < totHC Then
                intMoney = intMoney - (priHC * 10)
                priHC = priHC + 1
                intHC = intHC + 10
                lblCostStock.Caption = "Stock Price: $" & priHC
                lblStats.Caption = "You currently own " & intHC & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intHC / totHC) * 100)) & "% of the company"
            Else
                MsgBox "You own the entire company already!"
            End If
        Else
            MsgBox "You don't have enough money!"
        End If
    End If
End Sub

Private Sub cmdLeave_Click()
    frmRPG.lblCash.Caption = "Cash: $" & vbLf & intMoney
    frmRPG.Enabled = True
    frmRPG.Visible = True
    frmMove.Visible = True
    Unload frmStock
End Sub

Private Sub cmdSell_Click()
    If lblName.Caption = "Spaz Co. Enterprises" Then
        If intSCE >= 10 Then
            intSCE = intSCE - 10
            priSCE = priSCE - 1
            lblCostStock.Caption = "Stock Price: $" & priSCE
            lblStats.Caption = "You currently own " & intSCE & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intSCE / totSCE) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    ElseIf lblName.Caption = "The White Market" Then
        If intTWM >= 10 Then
            intTWM = intTWM - 10
            priTWM = priTWM - 1
            lblCostStock.Caption = "Stock Price: $" & priTWM
            lblStats.Caption = "You currently own " & intTWM & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intTWM / totTWM) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    ElseIf lblName.Caption = "Shadow Hentai Enterprises" Then
        If intSHE >= 10 Then
            intSHE = intSHE - 10
            priSHE = priSHE - 1
            lblCostStock.Caption = "Stock Price: $" & priSHE
            lblStats.Caption = "You currently own " & intSHE & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intSHE / totSHE) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    ElseIf lblName.Caption = "Clit N' Sons" Then
        If intCNS >= 10 Then
            intCNS = intCNS - 10
            CNS = CNS - 1
            lblCostStock.Caption = "Stock Price: $" & CNS
            lblStats.Caption = "You currently own " & intCNS & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intCNS / totCNS) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    ElseIf lblName.Caption = "Bones Inc." Then
        If intBI >= 10 Then
            intBI = intBI - 10
            priBI = priBI - 1
            lblCostStock.Caption = "Stock Price: $" & priBI
            lblStats.Caption = "You currently own " & intBI & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intBI / totBI) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    ElseIf lblName.Caption = "Blazing Auto's" Then
        If intBA >= 10 Then
            intBA = intBA - 10
            priBA = priBA - 1
            lblCostStock.Caption = "Stock Price: $" & priBA
            lblStats.Caption = "You currently own " & intBA & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intBA / totBA) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    ElseIf lblName.Caption = "Life After Death Co." Then
        If intLADC >= 10 Then
            intLADC = intLADC - 10
            priLADC = priLADC - 1
            lblCostStock.Caption = "Stock Price: $" & priLADC
            lblStats.Caption = "You currently own " & intLADC & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intLADC / totLADC) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    ElseIf lblName.Caption = "Hutch In Corp." Then
        If intHC >= 10 Then
            intHC = intHC - 10
            priHC = priHC - 1
            lblCostStock.Caption = "Stock Price: $" & priHC
            lblStats.Caption = "You currently own " & intHC & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intHC / totHC) * 100)) & "% of the company"
        Else
            MsgBox "You have none of these shares to sell!"
        End If
    End If
End Sub

Private Sub CNS_Click()
    lblName.Caption = "Clit N' Sons"
    lblCostStock.Caption = "Stock Price: $" & priCNS
    lblStats.Caption = "You currently own " & intCNS & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intCNS / totCNS) * 100)) & "% of the company"
    If bolCNS = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub

Private Sub Form_Load()
    Randomize
End Sub

Private Sub HC_Click()
    lblName.Caption = "Hutch In Corp."
    lblCostStock.Caption = "Stock Price: $" & priHC
    lblStats.Caption = "You currently own " & intHC & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intHC / totHC) * 100)) & "% of the company"
    If bolHC = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub

Private Sub LADC_Click()
    lblName.Caption = "Life After Death Co."
    lblCostStock.Caption = "Stock Price: $" & priLADC
    lblStats.Caption = "You currently own " & intLADC & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intLADC / totLADC) * 100)) & "% of the company"
    If boLADC = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub

Private Sub SCE_Click()
    lblName.Caption = "Spaz Co. Enterprises"
    lblCostStock.Caption = "Stock Price: $" & priSCE
    lblStats.Caption = "You currently own " & intSCE & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intSCE / totSCE) * 100)) & "% of the company"
    If bolSCE = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub

Private Sub SHE_Click()
    lblName.Caption = "Shadow Hentai Enterprises"
    lblCostStock.Caption = "Stock Price: $" & priSHE
    lblStats.Caption = "You currently own " & intSHE & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intSHE / totSHE) * 100)) & "% of the company"
    If bolSHE = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub

Private Sub TWM_Click()
    lblName.Caption = "The White Market"
    lblCostStock.Caption = "Stock Price: $" & priTWM
    lblStats.Caption = "You currently own " & intTWM & " shares of this company." & vbLf & vbLf & "Totalling " & Int(((intTWM / totTWM) * 100)) & "% of the company"
    If bolTWM = True Then
        lblCostStock.ForeColor = &H8000&
    Else
        lblCostStock.ForeColor = &HC0&
    End If
End Sub
