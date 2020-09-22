VERSION 5.00
Object = "*\AprjHolePunch.vbp"
Begin VB.Form frmTest 
   Caption         =   "Shape Test"
   ClientHeight    =   5190
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin prjHolePunch.TransShape TransShapeCustom 
      Height          =   690
      Left            =   2430
      Top             =   2220
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   1217
      Shape           =   9
   End
   Begin prjHolePunch.TransShape TransShape8 
      Height          =   5475
      Left            =   5400
      Top             =   435
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   9657
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   390
      Left            =   3540
      TabIndex        =   9
      Top             =   4290
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   390
      Left            =   3525
      TabIndex        =   8
      Top             =   3915
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   345
      TabIndex        =   7
      Top             =   3510
      Width           =   1950
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option1"
      Height          =   210
      Left            =   435
      TabIndex        =   6
      Top             =   3120
      Width           =   1770
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   210
      Left            =   435
      TabIndex        =   5
      Top             =   2790
      Width           =   1770
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   495
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   900
      Width           =   1890
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   510
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   510
      Width           =   1890
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   225
      Left            =   2940
      TabIndex        =   2
      Top             =   2010
      Width           =   1665
   End
   Begin VB.TextBox Text1 
      Height          =   1470
      Left            =   2925
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   465
      Width           =   2070
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Command1"
      Height          =   390
      Left            =   3510
      TabIndex        =   0
      Top             =   3540
      Width           =   1335
   End
   Begin prjHolePunch.TransShape TransShape12 
      Height          =   165
      Left            =   345
      Top             =   2460
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   291
   End
   Begin prjHolePunch.TransShape TransShape9 
      Height          =   1110
      Left            =   1635
      Top             =   -765
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1958
   End
   Begin prjHolePunch.TransShape TransShape10 
      Height          =   5370
      Left            =   2670
      Top             =   60
      Width           =   165
      _ExtentX        =   291
      _ExtentY        =   9472
   End
   Begin prjHolePunch.TransShape TransShape13 
      Height          =   960
      Left            =   4785
      Top             =   2055
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1693
      Shape           =   6
   End
   Begin prjHolePunch.TransShape TransShape11 
      Height          =   960
      Left            =   -525
      Top             =   2055
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1693
      Shape           =   4
   End
   Begin prjHolePunch.TransShape TransShape1 
      Height          =   330
      Left            =   -255
      Top             =   45
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   582
   End
   Begin prjHolePunch.TransShape TransShape3 
      Height          =   960
      Left            =   2145
      Top             =   4635
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1693
      Shape           =   3
   End
   Begin prjHolePunch.TransShape TransShape4 
      Height          =   765
      Left            =   -240
      Top             =   -15
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1349
      Shape           =   4
   End
   Begin prjHolePunch.TransShape TransShape5 
      Height          =   765
      Left            =   4980
      Top             =   0
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1349
      Shape           =   6
   End
   Begin prjHolePunch.TransShape TransShape6 
      Height          =   765
      Left            =   -435
      Top             =   4620
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1349
      Shape           =   3
   End
   Begin prjHolePunch.TransShape TransShape7 
      Height          =   765
      Left            =   5145
      Top             =   4605
      Width           =   750
      _ExtentX        =   1323
      _ExtentY        =   1349
      Shape           =   3
   End
   Begin prjHolePunch.TransShape TransShape14 
      Height          =   165
      Left            =   -1170
      Top             =   5145
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   291
   End
   Begin prjHolePunch.TransShape TransShape15 
      Height          =   5475
      Left            =   -180
      Top             =   60
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   9657
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
     Dim x(0 To 7)
     Dim y(0 To 7)
     ' Set shape to custom
     TransShapeCustom.Shape = CustomPolygon
     ' setup points
     x(0) = 13:  y(0) = 0
     x(1) = 0:   y(1) = 13
     x(2) = 0:   y(2) = 29
     x(3) = 13:  y(3) = 42
     x(4) = 30:  y(4) = 42
     x(5) = 42:  y(5) = 29
     x(6) = 42:  y(6) = 13
     x(7) = 29:   y(7) = 0
     ' call our custom shape method
     Call TransShapeCustom.DrawCustomShape(x, y, 8)
End Sub
