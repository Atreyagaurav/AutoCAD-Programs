Attribute VB_Name = "Draw"
Public Sub drawEquation()
Dim x() As Double
Dim y() As Double
Dim points() As Double
Dim obj As AcadObject
Dim pt As Variant
'==============================delta X and total distance=================================
delX = 0.5
maxX = 10
'=========================================================================================
n = Fix(maxX / delX)
ReDim x(0 To n)
ReDim y(0 To n)
ReDim points(0 To 2 * n + 1)
For i = 0 To n - 1:
    x(i) = delX * i
Next
x(n) = maxX
For i = 0 To n:
    '====================================Equation Here====================================
    'y(i) = Math.Sin(x(i))
    y(i) = -1 / 2 * (x(i)) ^ (1.85) / ((4.46) ^ (0.85)) 'ogeeee
    'y(i) = (1- * (x(i)) ^ (1.85) / ((4.46) ^ (0.85)) 'ellipse
    '=====================================================================================
Next
pt = ThisDrawing.Utility.GetPoint(, "Click at the curve origin")
For i = 0 To n:
    points(2 * i) = pt(0) + x(i)
    points(2 * i + 1) = pt(1) + y(i)
Next
Set obj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
End Sub
'=========================================================================================
Public Sub drawOgee()
Dim x() As Double
Dim y() As Double
Dim points() As Double
Dim obj As AcadObject
Dim pt As Variant
Dim H As Double
'==============================delta X and total distance=================================
delX = 0.5
maxX = 10
'=========================================================================================
n = Fix(maxX / delX)
ReDim x(0 To n)
ReDim y(0 To n)
ReDim points(0 To 2 * n + 1)
For i = 0 To n - 1:
    x(i) = delX * i
Next
x(n) = maxX
H = ThisDrawing.Utility.GetReal("Enter Head Over Crest")
For i = 0 To n:
    '====================================Equation Here====================================
    y(i) = -1 / 2 * (x(i)) ^ (1.85) / ((H) ^ (0.85)) 'ogeeee
    '=====================================================================================
Next
pt = ThisDrawing.Utility.GetPoint(, "Click at the Crest Point of the Dam")
For i = 0 To n:
    points(2 * i) = pt(0) + x(i)
    points(2 * i + 1) = pt(1) + y(i)
Next
Set obj = ThisDrawing.ModelSpace.AddLightWeightPolyline(points)
End Sub
'=========================================================================================
Public Sub SlimLayers()
For Each lyr In ThisDrawing.Layers
If lyr.Lineweight = -3 Then
lyr.Lineweight = acLnWt005
End If
Next
End Sub
Public Sub scaleLine()
Dim L As Double
Dim pt1 As Variant
Dim obj As AcadObject
ThisDrawing.Utility.GetEntity obj, pt1, "Click at the Line"
L = ThisDrawing.Utility.GetDistance(pt1, "the new distance")
obj.ScaleEntity obj.StartPoint, L / obj.Length
End Sub
Public Sub superIndexing()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
'ss.Clear
'ss.SelectOnScreen
ThisDrawing.Layers.Add ("0_Contour25")
ThisDrawing.Layers.Add ("0_Contour125")
For Each ob In ss:
    ob.Layer = "0_Contour25"
    If (ob.Elevation Mod 125 = 0) Then ob.Layer = "0_Contour125"
    'If (ob.Elevation Mod 10 = 0) Then ob.Layer = "Contour10"
    'If (ob.Elevation Mod 25 = 0) Then ob.Layer = "0_Contour25"
    'If (ob.Elevation Mod 50 = 0) Then ob.Layer = "Contour50"
Next
End Sub
Public Sub ThicknessByLayer()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
'ss.SelectOnScreen
For Each ob In ss
    ob.Lineweight = acLnWtByLayer
Next
End Sub
Public Sub cross()
Dim pl As AcadObject
Dim pt As AcadPoint
Dim ord As Variant
Dim pt1(0 To 2) As Double
Dim pt2(0 To 2) As Double
Dim pt3(0 To 2) As Double
On Error GoTo en
While True
    ThisDrawing.Utility.GetEntity pl, pt, "Click at a rectangle"
    ord = pl.Coordinates
    pt1(0) = ord(0)
    pt1(1) = ord(1)
    pt1(2) = pl.Elevation
    pt2(0) = ord(2)
    pt2(1) = ord(3)
    pt2(2) = pl.Elevation
    pt3(0) = ord(4)
    pt3(1) = ord(5)
    pt3(2) = pl.Elevation
    ThisDrawing.ModelSpace.AddLine pt1, pt3
    pt1(0) = ord(6)
    pt1(1) = ord(7)
    pt1(2) = pl.Elevation
    ThisDrawing.ModelSpace.AddLine pt1, pt2
Wend
en:
End Sub
Public Sub supertemp()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
ss.Clear
ss.SelectOnScreen
For Each ob In ss:
On Error Resume Next
If ob.ObjectName Like "*Polyline" Then
    If (ob.Layer = "CONTOUR" And Not (ob.Elevation Mod 5 = 0)) Then ob.Delete
End If
Next
End Sub
