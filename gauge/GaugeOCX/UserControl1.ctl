VERSION 5.00
Begin VB.UserControl gauge 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   ControlContainer=   -1  'True
   HitBehavior     =   2  'Use Paint
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   1590
   ScaleWidth      =   1590
   ToolboxBitmap   =   "UserControl1.ctx":84C2
   Begin VB.Shape Shape1 
      Height          =   1590
      Left            =   0
      Top             =   0
      Width           =   1590
   End
   Begin VB.Label lblReading 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   540
      TabIndex        =   0
      Top             =   1095
      Width           =   525
   End
   Begin VB.Line LinePointer 
      BorderWidth     =   2
      X1              =   810
      X2              =   315
      Y1              =   795
      Y2              =   795
   End
End
Attribute VB_Name = "gauge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' FREEWARE OCX guage
' by Diaa Eldessouky
' diaa1972@gmail.com
'*******************************************************************************
Option Explicit

Const pointerLength As Integer = 600 ' gauge pointer length
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Dim pointerUnitAngle As Single ' the unit angle by which the pointer will change
Dim PI As Double ' PI= 3.14159265358979 !!
Dim numMin As Single ' the minimum gauge reading
Dim numMax As Single ' the maximum gauge reading
Dim numValue As Single ' the current gauge reading

Public Property Get min() As Single
    min = numMin
End Property

Public Property Let min(mini As Single)
    numMin = mini
    PropertyChanged "min"
    initgauge min, max
End Property

Public Property Get max() As Single
    max = numMax
End Property

Public Property Let max(maxi As Single)
    numMax = maxi
    PropertyChanged "max"
    initgauge min, max
End Property

Public Property Get value() As Single
    value = numValue
End Property

Public Property Let value(reading As Single)
    numValue = reading
    PropertyChanged "value"
    setReading reading
End Property

Private Sub setReading(curValue As Single)
    curValue = curValue - min
    lblReading = curValue + min
    movePointer (180 + curValue * pointerUnitAngle)
        ' 180 to start the pointer motion at the left side of the gauge
End Sub

Private Sub initgauge(minReading As Single, maxReading As Single)
    Cls
    UserControl_Resize
    LinePointer.X1 = UserControl.Width / 2 - 22
    LinePointer.Y1 = UserControl.Height / 2 - 22
    DrawGaugeLimits
     
    If maxReading = minReading Then maxReading = minReading + 1 ' to avoid dev/zero
    pointerUnitAngle = 180 / (maxReading - minReading)
            ' 180 means that the pointer will move in half circle
    value = minReading ' initial gauge reading will be the minimum one
End Sub

Private Sub movePointer(curcurValue As Single)
   LinePointer.X2 = LinePointer.X1 + pointerLength * Cos(curcurValue * PI / 180)
   LinePointer.Y2 = LinePointer.Y1 + pointerLength * Sin(curcurValue * PI / 180)
End Sub

Private Sub UserControl_Initialize()
    PI = 4 * Atn(1) ' thanx to Ian Bunting
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 1600
    UserControl.Height = 1600
End Sub

Private Sub DrawGaugeLimits()
    ' thanx to Roger Gilchrist
   Dim tmpShift As Integer
   tmpShift = 130 ' this number is got by trials
   ForeColor = vbBlack
   TextOut hdc, tmpShift / 15, Height / 30, Str$(min), Len(Str$(min))
   ForeColor = vbRed
   TextOut hdc, (Width - tmpShift * 2 - (Len(Str$(max)) - 1) * 100) / 15, Height / 30, Str$(max), Len(Str$(max))
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "min", numMin, 0
    PropBag.WriteProperty "max", numMax, 1
    PropBag.WriteProperty "value", numValue, numMin
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    min = PropBag.ReadProperty("min", 0)
    max = PropBag.ReadProperty("max", 1)
    value = PropBag.ReadProperty("value", numMin)
End Sub

