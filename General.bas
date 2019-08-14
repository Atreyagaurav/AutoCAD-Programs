Attribute VB_Name = "General"
Public Sub area_multiple()
multiplearea.show
End Sub
Public Sub Numbering()
Dim txt As AcadObject
Dim pt As AcadPoint
Dim n As Integer
n = (ThisDrawing.Utility.GetInteger("Enter Initial Count"))
On Error GoTo er
While True
ThisDrawing.Utility.GetEntity txt, pt, "Pick text"
If n < 10 Then txt.TextString = "0" & CStr(n) Else: txt.TextString = CStr(n)
n = n + 1
Wend
er:
End Sub
Public Sub numberingABC()
Dim txt As AcadObject
Dim pt As AcadPoint
Dim n As Integer
n = 0
On Error GoTo er
While True
ThisDrawing.Utility.GetEntity txt, pt, "Pick text"
txt.TextString = Chr(Asc("A") + n)
n = n + 1
Wend
er:
End Sub

Public Sub AreaText()
Dim ar1 As Double
Dim ar2 As Double
Dim obj1 As AcadObject
Dim obj2 As AcadObject
Dim pt As AcadPoint
While True
On Error GoTo er:
ThisDrawing.Utility.GetEntity obj1, pt, "Click at Foundation Wall"
ThisDrawing.Utility.GetEntity obj2, pt, "Click at Excavation"
ar1 = obj1.area
ar2 = obj2.area
ThisDrawing.Utility.GetEntity obj1, pt, "Area of Foundation Wall"
obj1.TextString = str(Fix(ar1 * 1000000 + 0.5) / 1000000)
ThisDrawing.Utility.GetEntity obj2, pt, "Area of Excavation"
obj2.TextString = str(Fix(ar2 * 1000000 + 0.5) / 1000000)
Wend
er:
End Sub
Public Sub TextCopy()
Dim ob1 As AcadObject
Dim ob2 As AcadObject
Dim pt As AcadPoint

While True
On Error GoTo er:
ThisDrawing.Utility.GetEntity ob1, pt, "Click at Source text"
ThisDrawing.Utility.GetEntity ob2, pt, "Click at Destination text"
ob2.TextString = ob1.TextString
Wend
er:
End Sub
Public Sub BoudaryWalls()
Dim pt1 As Variant
Dim pt2 As Variant
Dim ob As AcadObject
Dim D As Double
Dim numJoint As Integer
Dim numPillar As Integer
Dim numWalls As Integer
Dim wallSeperation As Double
Dim templ As AcadLine
Dim vall As Variant
Dim pt As AcadPoint
Dim ptt(0 To 2) As Double
Dim n As Integer
n = 0
While True
On Error GoTo er:
pt1 = ThisDrawing.Utility.GetPoint(, "click at the corner of a pillar")
pt2 = ThisDrawing.Utility.GetPoint(pt1, "Click at the corner of another pillar")
Set templ = ThisDrawing.ModelSpace.AddLine(pt1, pt2)
D = templ.Length
templ.Delete
numJoint = Fix(D / 9.28) 'lenth between expansion joints c/c
numPillar = Fix((D - 9.28 * numJoint) / 3) + 4 * numJoint '  c/c length , no of wall segments and no of pillars  betn exp joints ' first part for residual part and second for proper segments
numWalls = numPillar - numJoint + 1
wallSeperation = (D - numPillar * 0.23 - numJoint * 0.05) / numWalls 'pillar width and expansion jointt width
ThisDrawing.Utility.GetEntity ob, pt, "Text: Number of joints"
ob.TextString = str(numJoint)
ThisDrawing.Utility.GetEntity ob, pt, "Text: Number of pillars"
ob.TextString = str(numPillar)
ThisDrawing.Utility.GetEntity ob, pt, "Text: Number of wall segments"
ob.TextString = str(numWalls)
ThisDrawing.Utility.GetEntity ob, pt, "Text: Distance of each segment"
ob.TextString = str(Fix(wallSeperation * 1000 + 0.5) / 1000)
'For i = 0 To numWalls - 1
'On Error GoTo er2
'ThisDrawing.Utility.GetEntity ob, pt, "Milayera click garne :P "
'vall = ob.GetDynamicBlockProperties
''ptt(0) = pt1(0) + (pt2(0) - pt1(0)) / d * (0.05 * Fix(n / 3) + 0.35 * (n + Fix(n / 3) - 1) + n * wallSeperation)
''ptt(1) = pt1(1) + (pt2(1) - pt1(1)) / d * (0.05 * Fix(n / 3) + 0.35 * (n + Fix(n / 3) - 1) + n * wallSeperation)
''ptt(2) = 0#
''ob.InsertionPoint = ptt
''n = n + 1
'For Each prop In vall
'If prop.PropertyName = "Distance1" Then
'prop.Value = wallSeperation
'End If
'Next
'Next
er2:
Wend
er:
End Sub

Public Sub SumText()
Dim ob As AcadObject
Dim su As Double
Dim ss As AcadSelectionSet
Dim pt As AcadPoint
su = 0#
Set ss = ThisDrawing.ActiveSelectionSet
For Each txt In ss
    On Error GoTo er:
    su = su + CDbl(txt.TextString)
er:
Next
ThisDrawing.Utility.GetEntity ob, pt, "Text to display sum"
ob.TextString = CStr(su)
End Sub

Public Sub texts()
Dim str As String
Dim ss As AcadSelectionSet
Dim ob As DataObject
Set ob = New DataObject
str = ""
Set ss = ThisDrawing.ActiveSelectionSet
For Each obb In ss
If obb.ObjectName = "AcDbText" Or obb.ObjectName = "AcDbMText" Then
    str = str & obb.TextString & ","
End If
Next
ob.SetText (str)
ob.PutInClipboard
End Sub
Public Sub co_ordinates()
Dim str As String
Dim ss As AcadSelectionSet
Dim ccl As AcadCircle
Dim ob As AcadObject
Dim obb As AcadObject
Dim i, j As Integer
Dim pt As Variant
Dim pt2(0 To 2) As Double
Dim x(0 To 40) As Double
Dim y(0 To 40) As Double
Dim z(0 To 40) As Double
Set ss = ThisDrawing.ActiveSelectionSet
ss.Clear
ss.SelectOnScreen
i = 0
For Each obj In ss
    If obj.ObjectName = "AcDbCircle" Then
        pt = obj.Center
        x(i) = pt(0)
        y(i) = pt(1)
        i = i + 1
    End If
Next
ss.Clear
ss.SelectOnScreen
i = 0
For Each obj In ss
    If obj.ObjectName = "AcDbText" Or obj.ObjectName = "AcDbMText" Then
        z(i) = CDbl(Right(obj.TextString, 9))
        i = i + 1
    End If
Next
ThisDrawing.Utility.GetEntity ob, pt, "Text to display x"
For j = 0 To i - 1
    Set obb = ob.Copy()
    pt = ob.InsertionPoint
    pt2(0) = pt(0)
    pt2(2) = pt(2)
    pt2(1) = pt(1) - 3 * (j + 1)
    obb.Move pt, pt2
    obb.TextString = CDbl(x(j))
Next
ThisDrawing.Utility.GetEntity ob, pt, "Text to display y"
For j = 0 To i - 1
    Set obb = ob.Copy()
    pt = ob.InsertionPoint
    pt2(0) = pt(0)
    pt2(2) = pt(2)
    pt2(1) = pt(1) - 3 * (j + 1)
    obb.Move pt, pt2
    obb.TextString = CDbl(y(j))
Next
ThisDrawing.Utility.GetEntity ob, pt, "Text to display z"
For j = 0 To i - 1
    Set obb = ob.Copy()
    pt = ob.InsertionPoint
    pt2(0) = pt(0)
    pt2(2) = pt(2)
    pt2(1) = pt(1) - 3 * (j + 1)
    obb.Move pt, pt2
    obb.TextString = CDbl(z(j))
Next

End Sub
Public Sub indexing()
Dim ss As AcadSelectionSet
Dim el As Double
Dim x As Double
Dim pt1(0 To 2) As Double
Dim pt2(0 To 2) As Double
pt1(0) = 0#
pt1(1) = 0#
pt1(2) = 0#
pt2(0) = 20000#
pt2(1) = 0#
pt2(2) = 0#
Set ss = ThisDrawing.ActiveSelectionSet
For Each obj In ss
    If obj.ObjectName Like "*Text" Then
        el = CDbl(obj.TextString)
        If (Fix(el / 25) * 25 = el) Then
            x = 1
        Else
            obj.Delete
        End If
    End If
    If obj.ObjectName = "AcDbPolyline" Then
        el = obj.Elevation
        If (Fix(el / 25) * 25 = el) Then
            obj.Move pt1, pt2
        End If
    End If
Next
End Sub
Public Sub points()
Dim ob As AcadObject
Dim el As Double
Dim ss As AcadSelectionSet
Dim pt1(0 To 2) As Double
Dim pt2(0 To 2) As Double
Dim pt As Variant
el = 0#
pt1(0) = 0#
pt1(1) = 0#
pt1(2) = 0#
Set ss = ThisDrawing.ActiveSelectionSet
For Each ob In ss
        pt = ob.Coordinates
        el = Fix(Rnd(1) * 10)
        If (el Mod 4 = 0) Then
            ob.Delete
        End If
Next
End Sub
Public Sub copies()
Dim ob As AcadObject
Dim el As Double
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
el = 0.001
For Each ob In ss
If ob.ObjectName = "AcDbPolyline" Then
    If el = ob.Elevation Then
        ob.Delete
    Else
        el = ob.Elevation
    End If
End If
Next
End Sub
Public Sub cadestral()
Dim ob As AcadObject
Dim ob2 As AcadObject
Dim pt(0 To 2) As Double
Dim pt2(0 To 2) As Double
pt(0) = 0#
pt(1) = 0#
pt(2) = 0#
Dim i As Integer
Dim ptt As AcadPoint
ThisDrawing.Utility.GetEntity ob, ptt, "Click at the first text"
n = CInt(ob.TextString)
For m = 1 To 40
    For n = 1 To 40
        pt2(0) = (n - 1) * 1250
        pt2(1) = -(m - 1) * 1250
        pt2(2) = 0#
        i = (m - 1) * 40 + n
        Set ob2 = ob.Copy
        ob2.TextString = worksheetfunction.Text(i, "0000")
        ob2.Move pt, pt2
    Next
Next
End Sub

