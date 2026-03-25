VERSION 5.00
Begin VB.Form frmFace 
   BackColor       =   &H00000000&
   Caption         =   "Sierpinski Triangle"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   6285
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrTime1 
      Interval        =   500
      Left            =   1800
      Top             =   1800
   End
End
Attribute VB_Name = "frmFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Double           'iteration Number
Dim Radius As Double
Dim x1 As Double          'Top
Dim y1 As Double          'Top
Dim x2 As Double          'Left
Dim y2 As Double          'Left
Dim x3 As Double          'Right
Dim y3 As Double          'Right


Private Sub Form_Load()
 Cls
 x1 = 3200            'Me.Width / 2
 y1 = 1
 x2 = 1
 y2 = 4800            'Me.Height
 x3 = 6400            'Me.Width
 y3 = 4800            'Me.Height
End Sub

Private Sub tmrTime1_Timer()
For i = 1 To 100
  R = Rnd(1)
  If R < 0.34 Then
   x = (x + x1) / 2
   y = (y + y1) / 2
  ElseIf R < 0.67 Then
   x = (x + x2) / 2
   y = (y + y2) / 2
  Else
   x = (x + x3) / 2
   y = (y + y3) / 2
End If
PSet (x, y), RGB(255, 255, 255)
Next i
End Sub
