VERSION 5.00
Begin VB.Form frmMove 
   BorderStyle     =   0  'None
   Caption         =   "Move"
   ClientHeight    =   1815
   ClientLeft      =   9165
   ClientTop       =   630
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "S"
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "W"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E"
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "N"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "1"
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Y:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   990
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   645
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0000C000&
      FillColor       =   &H0000C000&
      Height          =   1815
      Left            =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim X As Integer, Y As Integer

Private Sub Command1_Click()
    If Y > 1 Then
        Y = Y - 1
    End If
    Text2.Text = Y
    
    MapX(X) = True
    MapY(Y) = True
    frmCityMap.Label4.Caption = "(" & X & "," & Y & ")"
    frmCityMap.Refresh
End Sub

Private Sub Command2_Click()
    If X < 10 Then
        X = X + 1
    End If
    Text1.Text = X
    
    MapX(X) = True
    MapY(Y) = True
    frmCityMap.Label4.Caption = "(" & X & "," & Y & ")"
    frmCityMap.Refresh
End Sub

Private Sub Command3_Click()
    If X > 1 Then
        X = X - 1
    End If
    Text1.Text = X
    
    MapX(X) = True
    MapY(Y) = True
    frmCityMap.Label4.Caption = "(" & X & "," & Y & ")"
    frmCityMap.Refresh
End Sub

Private Sub Command4_Click()
    If Y < 10 Then
        Y = Y + 1
    End If
    Text2.Text = Y
    
    MapX(X) = True
    MapY(Y) = True
    frmCityMap.Label4.Caption = "(" & X & "," & Y & ")"
    frmCityMap.Refresh
End Sub

Private Sub Form_Load()
    X = 1
    Y = 1
    EnemyThere = True
    Text1.Text = X
    Text2.Text = Y
    
    MapX(X) = True
    MapY(Y) = True
    frmCityMap.Label4.Caption = "(" & X & "," & Y & ")"
    frmCityMap.Refresh
End Sub
