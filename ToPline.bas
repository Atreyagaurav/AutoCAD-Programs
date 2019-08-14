Attribute VB_Name = "ToPline"
Sub arc2lines()

Dim myArc As AcadArc
Dim objSel As AcadEntity
Dim myPL As AcadLWPolyline
Dim mypolarpoint
Dim bulge() As Double
Dim legs As Integer
Const PI = 3.14159265358979
Dim delta As Double

ThisDrawing.Utility.GetEntity objSel, returnObj, "Select Arc:"
If objSel.ObjectName = "AcDbArc" Then
Set myArc = objSel
delta = myArc.EndAngle - myArc.StartAngle
If delta < 0 Then delta = delta + (2 * PI)
Dim numOfSegments As Integer
Dim points() As Double
'adjust below for reality
numOfSegments = CInt(myArc.ArcLength) ' length of segment = 1, last segment = remainder
ReDim points(0 To 2 * numOfSegments + 1)
ang = 1 / myArc.Radius
points(0) = myArc.StartPoint(0)
points(1) = myArc.StartPoint(1)
adir = ang
For x = 2 To UBound(points) - 2 Step 2
mypolarpoint = ThisDrawing.Utility.PolarPoint(myArc.Center, myArc.StartAngle + adir, myArc.Radius)
adir = adir + ang
points(x) = mypolarpoint(0)
points(x + 1) = mypolarpoint(1)
Next x
points(UBound(points) - 1) = myArc.EndPoint(0)
points(UBound(points) - 0) = myArc.EndPoint(1)
Set myPL = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
myPL.Update
End If

End Sub

Public Sub Pclose()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
For Each obj In ss
    If obj.ObjectName Like "*Polyline" Then
        obj.Closed = True
    End If
Next
End Sub
Public Sub isclosed()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
Dim n As Integer
n = 0
For Each obj In ss
    If obj.ObjectName Like "*Polyline" Then
        ThisDrawing.Utility.Prompt Chr(10) & "Pline: " & CStr(n) & " : " & obj.Closed
        n = n + 1
    End If
Next
End Sub
