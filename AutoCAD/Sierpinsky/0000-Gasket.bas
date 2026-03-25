Sub Ch4_CreatePoint()
ThisDrawing.SetVariable "PDMODE", 0
ThisDrawing.SetVariable "PDSIZE", 1

 Dim pointObj As AcadPoint
 Dim location(0 To 2) As Double
 
 Dim i As Double           'Iteration Number
 Dim x1 As Double         'Vertices
 Dim y1 As Double
 Dim x2 As Double
 Dim y2 As Double
 Dim x3 As Double
 Dim y3 As Double
 Dim x4 As Double
 Dim y4 As Double

 x1 = 1
 y1 = 1
 x2 = -1
 y2 = 1
 x3 = -1
 y3 = -1
 x4 = 1
 y4 = -1

 For i = 1 To 10000
  R = Rnd(1)
  If R < 0.25 Then
   x = (x + x1) / 3
   y = (y + y1) / 3
  ElseIf R < 0.5 Then
   x = (x + x2) / 3
   y = (y + y2) / 3
  ElseIf R < 0.75 Then
   x = (x + x3) / 3
   y = (y + y3) / 3
  Else
   x = (x + x4) / 3
   y = (y + y4) / 3
 End If
  ' Define the location of the point
  location(0) = x
  location(1) = y
  location(2) = 0
  ' Create the point
  Set pointObj = ThisDrawing.ModelSpace.AddPoint(location)
 Next i
 ZoomExtents
End Sub

