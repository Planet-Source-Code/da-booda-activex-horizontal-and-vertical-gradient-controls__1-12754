VERSION 5.00
Object = "*\AControlDemo.vbp"
Object = "*\AControlDemo2.vbp"
Begin VB.Form Controls 
   Caption         =   "Control Demo"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   369
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VGradientDemo.GradientVScroll Scroll2 
      Height          =   1905
      Left            =   3960
      TabIndex        =   10
      Top             =   360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3360
      StartRed        =   100
      StartGreen      =   100
      StartBlue       =   100
      EndRed          =   100
      EndGreen        =   100
      EndBlue         =   100
      CellOutLined    =   0   'False
      Max             =   24
      BackRed         =   100
      BackGreen       =   100
      BackBlue        =   100
      BarSolid        =   -1  'True
      Step            =   5
      Value           =   24
   End
   Begin HGradientDemo.GradientHScroll Scroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   450
      StartRed        =   100
      StartGreen      =   100
      StartBlue       =   100
      EndRed          =   100
      EndGreen        =   100
      EndBlue         =   100
      Max             =   24
      BackRed         =   100
      BackGreen       =   100
      BackBlue        =   100
      BarSolid        =   -1  'True
      Step            =   10
      Value           =   0
   End
   Begin VGradientDemo.GradientVScroll Percent4 
      Height          =   780
      Left            =   4320
      TabIndex        =   8
      Top             =   3120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   1376
      StartRed        =   255
      EndGreen        =   0
      EndBlue         =   0
      CellOutLined    =   0   'False
      Max             =   50
      Min             =   1
      BoxRed          =   200
      BoxGreen        =   200
      BoxBlue         =   200
      BackRed         =   255
      BackGreen       =   255
      BarVisible      =   0   'False
      Value           =   1
   End
   Begin HGradientDemo.GradientHScroll Percent3 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3000
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   450
      StartRed        =   255
      StartGreen      =   255
      EndGreen        =   0
      Max             =   250
      Min             =   1
      BoxRed          =   200
      BoxGreen        =   200
      BoxBlue         =   200
      BarVisible      =   0   'False
      ScrollEnabled   =   0   'False
      Value           =   10
   End
   Begin HGradientDemo.GradientHScroll Percent2 
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   450
      StartGreen      =   255
      EndRed          =   0
      EndBlue         =   0
      Max             =   25
      Min             =   1
      BoxRed          =   200
      BoxGreen        =   200
      BoxBlue         =   200
      BarVisible      =   0   'False
      Step            =   10
      ScrollEnabled   =   0   'False
      Value           =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   4200
      Top             =   3960
   End
   Begin HGradientDemo.GradientHScroll Percent 
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3720
      Width           =   3780
      _ExtentX        =   6668
      _ExtentY        =   450
      StartBlue       =   255
      EndRed          =   0
      EndGreen        =   0
      CellOutLined    =   0   'False
      Max             =   50
      Min             =   1
      BoxRed          =   200
      BoxGreen        =   200
      BoxBlue         =   200
      BarVisible      =   0   'False
      Step            =   5
      ScrollEnabled   =   0   'False
      Value           =   1
   End
   Begin HGradientDemo.GradientHScroll RedChoose 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   4440
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   450
      EndGreen        =   0
      EndBlue         =   0
      Max             =   25
      BoxRed          =   0
      BoxGreen        =   0
      BoxBlue         =   0
      Step            =   10
      Value           =   17
   End
   Begin HGradientDemo.GradientHScroll GreenChoose 
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   4800
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   450
      EndRed          =   0
      EndBlue         =   0
      Max             =   25
      BoxRed          =   0
      BoxGreen        =   0
      BoxBlue         =   0
      Step            =   10
      Value           =   15
   End
   Begin HGradientDemo.GradientHScroll BlueChoose 
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5160
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   450
      EndRed          =   0
      EndGreen        =   0
      Max             =   25
      BoxRed          =   0
      BoxGreen        =   0
      BoxBlue         =   0
      Step            =   10
      Value           =   14
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "----------Scroll Bars----------"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   75
      Left            =   135
      Shape           =   2  'Oval
      Top             =   375
      Width           =   150
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "----------Percent Bars----------"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "----------Color Chooser----------"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Width           =   3975
   End
End
Attribute VB_Name = "Controls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Form is just a demo
'nothing fancy
'Just shows you what the controls can be used for
'and how to optimize the look
Private Sub BlueChoose_Change()
Me.BackColor = RGB(RedChoose.Value * 10, GreenChoose.Value * 10, BlueChoose.Value * 10)
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(170, 150, 140)
End Sub

Private Sub GreenChoose_Change()
Me.BackColor = RGB(RedChoose.Value * 10, GreenChoose.Value * 10, BlueChoose.Value * 10)
End Sub

Private Sub RedChoose_Change()
Me.BackColor = RGB(RedChoose.Value * 10, GreenChoose.Value * 10, BlueChoose.Value * 10)
End Sub

Private Sub Scroll1_Change()
Dim a As Integer
a = Scroll1.Value
a = a * 10
Shape1.Left = 9 + a
End Sub

Private Sub Scroll2_Change()
Dim a As Integer
a = Scroll2.Value
a = 24 - a
a = a * 5
Shape1.Top = 27 + a
End Sub

Private Sub Timer1_Timer()
Dim a As Integer
a = Percent.Value
a = a + 1
If a > Percent.Max Then a = Percent.Min
Percent.Value = a
Percent2.Value = Int(a / 2)
Percent3.Value = Int(a / 2) * 10
Percent4.Value = a
End Sub
