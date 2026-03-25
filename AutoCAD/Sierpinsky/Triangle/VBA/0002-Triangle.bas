Sub Ch4_CreatePoint()
ThisDrawing.SetVariable "PDMODE", 0
ThisDrawing.SetVariable "PDSIZE", 1

Dim pointObj As AcadPoint
Dim location(0 To 2) As Double

Dim pi As Double
Dim i As Double		'Iteration Number
Dim R As Double		'Random Number
Dim x(0 To 2) As Double 'X Coordinate
Dim y(0 To 2) As Double	'Y Coordinate

location(2) = 0		'Z Coordinate
pi = 3.1415926535

x(0) = Cos(pi / 2)	'Vertices
y(0) = Sin(pi / 2)
x(1) = Cos(7 * pi / 6)
y(1) = Sin(7 * pi / 6)
x(2) = Cos(11 * pi / 6)
y(2) = Sin(11 * pi / 6)

For i = 1 To 100000	
 R = Rnd(1)
 If R < 1 / 3 Then
  location(0) = (location(0) + x(0)) / 2
  location(1) = (location(1) + y(0)) / 2
 ElseIf R < 2 / 3 Then
  location(0) = (location(0) + x(1)) / 2
  location(1) = (location(1) + y(1)) / 2
 Else
  location(0) = (location(0) + x(2)) / 2
  location(1) = (location(1) + y(2)) / 2
End If
 ' Create the point
 Set pointObj = ThisDrawing.ModelSpace.AddPoint(location)
Next i
ZoomExtents
End Sub

