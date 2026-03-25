Sub Ch4_CreatePoint()
ThisDrawing.SetVariable "PDMODE", 0
ThisDrawing.SetVariable "PDSIZE", 1

 Dim pointObj As AcadPoint
 Dim location(0 To 2) As Double
 
 Dim pi As Double
 Dim k As Double            'Inner Edges' Circle
 Dim i As Double            'Number of Iterations
 Dim x01 As Double          'Vertices
 Dim y01 As Double
 Dim x02 As Double
 Dim y02 As Double
 Dim x03 As Double
 Dim y03 As Double
 Dim x04 As Double
 Dim y04 As Double
 Dim x05 As Double
 Dim y05 As Double
 Dim x06 As Double
 Dim y06 As Double
 Dim x07 As Double
 Dim y07 As Double
 Dim x08 As Double
 Dim y08 As Double
 Dim x09 As Double
 Dim y09 As Double
 Dim x10 As Double
 Dim y10 As Double
 Dim x11 As Double
 Dim y11 As Double
 Dim x12 As Double
 Dim y12 As Double

 pi = 3.141592654
 k = Cos(pi * 6) / 1.5
 
 x01 = Cos(pi / 2)
 y01 = Sin(pi / 2)
 x02 = Cos(2 * pi / 3) * k
 y02 = Sin(2 * pi / 3) * k
 x03 = Cos(5 * pi / 6)
 y03 = Sin(5 * pi / 6)
 x04 = Cos(pi) * k
 y04 = Sin(pi) * k
 x05 = Cos(7 * pi / 6)
 y05 = Sin(7 * pi / 6)
 x06 = Cos(4 * pi / 3) * k
 y06 = Sin(4 * pi / 3) * k
 x07 = Cos(3 * pi / 2)
 y07 = Sin(3 * pi / 2)
 x08 = Cos(5 * pi / 3) * k
 y08 = Sin(5 * pi / 3) * k
 x09 = Cos(11 * pi / 6)
 y09 = Sin(11 * pi / 6)
 x10 = Cos(0) * k
 y10 = Sin(0) * k
 x11 = Cos(pi / 6)
 y11 = Sin(pi / 6)
 x12 = Cos(pi / 3) * k
 y12 = Sin(pi / 3) * k

 For i = 1 To 100000
  R = Rnd(1)
  If R < 1 / 12 Then
   x = (x + x01) / 12
   y = (y + y01) / 12
  ElseIf R < 2 / 12 Then
   x = (x + x02) / 12
   y = (y + y02) / 12
  ElseIf R < 3 / 12 Then
   x = (x + x03) / 12
   y = (y + y03) / 12
  ElseIf R < 4 / 12 Then
   x = (x + x04) / 12
   y = (y + y04) / 12
  ElseIf R < 5 / 12 Then
   x = (x + x05) / 12
   y = (y + y05) / 12
  ElseIf R < 6 / 12 Then
   x = (x + x06) / 12
   y = (y + y06) / 12
  ElseIf R < 7 / 12 Then
   x = (x + x07) / 12
   y = (y + y07) / 12
  ElseIf R < 8 / 12 Then
   x = (x + x08) / 12
   y = (y + y08) / 12
  ElseIf R < 9 / 12 Then
   x = (x + x09) / 12
   y = (y + y09) / 12
  ElseIf R < 10 / 12 Then
   x = (x + x10) / 12
   y = (y + y10) / 12
  ElseIf R < 11 / 12 Then
   x = (x + x11) / 12
   y = (y + y11) / 12
  Else
   x = (x + x12) / 12
   y = (y + y12) / 12
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