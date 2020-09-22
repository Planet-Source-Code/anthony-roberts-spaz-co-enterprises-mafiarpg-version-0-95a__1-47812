VERSION 5.00
Begin VB.Form frmCityMap 
   BorderStyle     =   0  'None
   Caption         =   "City Map"
   ClientHeight    =   3975
   ClientLeft      =   9165
   ClientTop       =   2430
   ClientWidth     =   1695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   1695
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Refresh 
      Interval        =   100
      Left            =   1200
      Top             =   3600
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   1680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "(1,1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "You're at:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Your Home"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "(1,1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image1 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   480
      Width           =   135
   End
   Begin VB.Image Image10 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   1200
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   1080
      Width           =   135
   End
   Begin VB.Image Image8 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   960
      Width           =   135
   End
   Begin VB.Image Image7 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   840
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image5 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   600
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   120
      Width           =   135
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   1200
      Width           =   135
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   1080
      Width           =   135
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   960
      Width           =   135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   840
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   600
      Width           =   135
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   480
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   120
      Width           =   135
   End
   Begin VB.Image Image6 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   9
      Left            =   1320
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   8
      Left            =   1200
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   7
      Left            =   1080
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   6
      Left            =   960
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   4
      Left            =   720
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   3
      Left            =   600
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   2
      Left            =   480
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   360
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   135
      Index           =   1
      Left            =   360
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   840
      Top             =   720
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   240
      Width           =   135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   240
      Top             =   360
      Width           =   135
   End
End
Attribute VB_Name = "frmCityMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, j As Integer

Private Sub Command1_Click()
    frmRPG.Enabled = True
    frmRPG.Visible = True
    frmMove.Visible = True
    Unload frmCityMap
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",1)"
    
    If Index = 0 Then
        Label2.Caption = "Your Home"
    Else
        Label2.Caption = ""
    End If
End Sub

Private Sub Image2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",2)"
    
    Label2.Caption = ""
End Sub

Private Sub Image3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",3)"
    
    Label2.Caption = ""
End Sub

Private Sub Image4_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",4)"
    
    Label2.Caption = ""
End Sub

Private Sub Image5_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",5)"
    
    Label2.Caption = ""
End Sub

Private Sub Image6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",6)"
    
    Label2.Caption = ""
End Sub

Private Sub Image7_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",7)"
    
    Label2.Caption = ""
End Sub

Private Sub Image8_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",8)"
    
    Label2.Caption = ""
End Sub

Private Sub Image9_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",9)"
    
    Label2.Caption = ""
End Sub

Private Sub Image10_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label1.Caption = "(" & (Index + 1) & ",10)"
    
    Label2.Caption = ""
End Sub

Private Sub Refresh_Timer()
    Label4.Caption = "(" & X & "," & Y & ")"
    If MapX(1) = True Then
        If MapY(1) = True Then
            Shape1(0).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(0).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(0).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(0).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(0).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(0).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(0).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(0).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(0).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(0).FillStyle = 1
        End If
    End If
    
    If MapX(2) = True Then
        If MapY(1) = True Then
            Shape1(1).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(1).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(1).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(1).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(1).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(1).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(1).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(1).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(1).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(1).FillStyle = 1
        End If
    End If
    
    If MapX(3) = True Then
        If MapY(1) = True Then
            Shape1(2).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(2).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(2).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(2).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(2).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(2).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(2).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(2).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(2).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(2).FillStyle = 1
        End If
    End If
    
    If MapX(4) = True Then
        If MapY(1) = True Then
            Shape1(3).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(3).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(3).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(3).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(3).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(3).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(3).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(3).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(3).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(3).FillStyle = 1
        End If
    End If
    
    If MapX(5) = True Then
        If MapY(1) = True Then
            Shape1(4).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(4).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(4).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(4).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(4).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(4).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(4).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(4).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(4).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(4).FillStyle = 1
        End If
    End If
    
    If MapX(6) = True Then
        If MapY(1) = True Then
            Shape1(5).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(5).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(5).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(5).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(5).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(5).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(5).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(5).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(5).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(5).FillStyle = 1
        End If
    End If
    
    If MapX(7) = True Then
        If MapY(1) = True Then
            Shape1(6).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(6).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(6).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(6).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(6).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(6).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(6).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(6).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(6).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(6).FillStyle = 1
        End If
    End If
    
    If MapX(8) = True Then
        If MapY(1) = True Then
            Shape1(7).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(7).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(7).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(7).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(7).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(7).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(7).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(7).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(7).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(7).FillStyle = 1
        End If
    End If
    
    If MapX(9) = True Then
        If MapY(1) = True Then
            Shape1(8).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(8).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(8).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(8).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(8).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(8).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(8).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(8).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(8).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(8).FillStyle = 1
        End If
    End If
    
    If MapX(10) = True Then
        If MapY(1) = True Then
            Shape1(9).FillStyle = 1
        End If
        
        If MapY(2) = True Then
            Shape2(9).FillStyle = 1
        End If
        
        If MapY(3) = True Then
            Shape3(9).FillStyle = 1
        End If
        
        If MapY(4) = True Then
            Shape4(9).FillStyle = 1
        End If
        
        If MapY(5) = True Then
            Shape5(9).FillStyle = 1
        End If
        
        If MapY(6) = True Then
            Shape6(9).FillStyle = 1
        End If
        
        If MapY(7) = True Then
            Shape7(9).FillStyle = 1
        End If
        
        If MapY(8) = True Then
            Shape8(9).FillStyle = 1
        End If
        
        If MapY(9) = True Then
            Shape9(9).FillStyle = 1
        End If
        
        If MapY(10) = True Then
            Shape10(9).FillStyle = 1
        End If
    End If
End Sub
