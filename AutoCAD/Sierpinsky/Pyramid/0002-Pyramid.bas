Sub Ch4_CreatePoint()
ThisDrawing.SetVariable "PDMODE", 0
ThisDrawing.SetVariable "PDSIZE", 1

Dim pointObj As AcadPoint
Dim location(0 To 2) As Double

Dim pi As Double
Dim i As Double           'iteration Number
Dim x(0 To 3) As Double
Dim y(0 To 3) As Double
Dim z(0 To 3) As Double
 
pi = 3.141592654

x(0) = 0
y(0) = 1
z(0) = 0
x(1) = Cos(7 * pi / 6)
y(1) = Sin(7 * pi / 6)
z(1) = 0
x(2) = Cos(11 * pi / 6)
y(2) = Sin(11 * pi / 6)
z(2) = 0
x(3) = 0
y(3) = 0
z(3) = 1.5

For i = 1 To 10000
 R = Rnd(1)
 Select Case R
 Case Is < 0.25
  location(0) = (location(0) + x(0)) / 2
  location(1) = (location(1) + y(0)) / 2
  location(2) = (location(2) + z(0)) / 2
 Case Is < 0.5
  location(0) = (location(0) + x(1)) / 2
  location(1) = (location(1) + y(1)) / 2
  location(2) = (location(2) + z(1)) / 2
 Case Is < 0.75
  location(0) = (location(0) + x(2)) / 2
  location(1) = (location(1) + y(2)) / 2
  location(2) = (location(2) + z(2)) / 2
 Case Else
  location(0) = (location(0) + x(3)) / 2
  location(1) = (location(1) + y(3)) / 2
  location(2) = (location(2) + z(3)) / 2
 End Select
 ' Create the point
 Set pointObj = ThisDrawing.ModelSpace.AddPoint(location)
 Next i
 ZoomExtents
End Sub