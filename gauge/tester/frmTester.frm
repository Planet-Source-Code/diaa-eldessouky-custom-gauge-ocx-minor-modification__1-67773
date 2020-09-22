VERSION 5.00
Object = "*\A..\GaugeOCX\prjGauge.vbp"
Begin VB.Form frmTester 
   Caption         =   "OCX tester"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Begin prjGauge.gauge gauge1 
      Height          =   1605
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   2831
      max             =   10
   End
   Begin VB.HScrollBar HSreading 
      Height          =   285
      Left            =   60
      Max             =   100
      Min             =   20
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1710
      Value           =   20
      Width           =   1545
   End
End
Attribute VB_Name = "frmTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    HSreading.Max = gauge1.Max
    HSreading.Min = gauge1.Min
    HSreading.Value = gauge1.Min
End Sub

Private Sub HSreading_Change()
    gauge1.Value = HSreading.Value
End Sub

Private Sub HSreading_Scroll()
    HSreading_Change
End Sub
