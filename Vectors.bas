Attribute VB_Name = "Vectors"
Public Function getDifference(ByRef p1() As Double, ByRef p2() As Double, ByRef diff() As Double)
For i = LBound(diff) To UBound(diff)
    diff(i) = p1(i) - p2(i)
Next
End Function
Public Function getAddition(ByRef p1() As Double, ByRef p2() As Double, ByRef diff() As Double)
For i = LBound(diff) To UBound(diff)
    diff(i) = p1(i) + p2(i)
Next
End Function
Public Function getDotProduct(ByRef p1() As Double, ByRef p2() As Double) As Double
Dim s As Double
s = 0
For i = LBound(p1) To UBound(p1) ' later check for same number of variables in array ' not time now
    s = s + p1(i) * p2(i)
Next
getDotProduct = s
End Function
Public Function getMagnitude(ByRef p1() As Double) As Double
Dim s As Double
s = 0
For i = LBound(p1) To UBound(p1)
    s = s + p1(i) * p1(i)
Next
getMagnitude = Math.Sqr(s)
End Function
Public Function divideRatio(r As Double, ByRef p1() As Double, ByRef p2() As Double, ByRef diff() As Double)
For i = LBound(diff) To UBound(diff)
    diff(i) = p1(i) + (p2(i) - p1(i)) * r
Next
End Function
Public Sub forTest()
Dim x1(0 To 2) As Double
Dim x2(0 To 2) As Double
Dim x3(0 To 2) As Double
x1(0) = 1
x2(0) = 2
y = getMagnitude(2, x2)
Debug.Print y
End Sub
