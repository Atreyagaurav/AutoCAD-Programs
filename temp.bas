Attribute VB_Name = "temp"
Public Sub cad()
Dim ob As AcadObject
Dim pt As AcadPoint
Dim tx As AcadText
Dim sr As String
Dim lr As AcadLayer
'On Error GoTo en
While True:
    ThisDrawing.Utility.GetEntity ob, pt, "polygon"
    ThisDrawing.Utility.GetEntity tx, pt, "text"
    sr = "015-" & tx.TextString
    Set lr = ThisDrawing.Layers.Add(sr)
    ob.Layer = lr.Name
Wend
en:
End Sub
Public Sub ObjectName()
Dim ob As AcadObject
Dim pt As AcadPoint
ThisDrawing.Utility.GetEntity ob, pt, "any"
MsgBox ob.ObjectName
End Sub
Public Sub prop()
Dim h1 As AcadObject
Dim h2 As AcadObject
Dim pt As AcadPoint
Dim p1 As Variant
Dim p2 As Variant
ThisDrawing.Utility.GetEntity h1, pt, "hatch1"
ThisDrawing.Utility.GetEntity h2, pt, "hatch2"
p1 = h1.Properties
p2 = h2.Properties
For i = LBound(p1) To UBound(p1)
    MsgBox (p1(i) & " " & p2(i))
Next

End Sub
Public Sub NameRivers()
Dim ob1 As AcadObject
Dim ob2 As AcadObject
Dim ob3 As AcadObject
Dim pt1 As AcadPoint
Dim pt2 As Variant
Dim pt3 As Variant
Dim prop As Variant
On Error GoTo en
ThisDrawing.Utility.GetEntity ob1, pt1, "Click at the label block"
pt2 = ob1.InsertionPoint
While True:
    ThisDrawing.Utility.GetEntity ob2, pt1, "Click at kholsi"
    Set ob3 = ob1.Copy
    pt3 = ThisDrawing.Utility.GetPoint(pt2, "click insertion location")
    ob3.Move pt2, pt3
    prop = ob3.GetAttributes
    For Each pr In prop
        If pr.TagString = "LABEL" Then
            pr.TextString = StrConv(ob2.Layer, vbUpperCase)
        End If
    Next
Wend
en:
End Sub
Public Sub colTexts()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
For Each obj In ss
    If obj.ObjectName = "AcDbText" Or obj.ObjectName = "AcDbMText" Then
        obj.TextString = obj.TextString & ":"
    End If
Next
End Sub
Public Sub DrawSlopePatternNoBlock()
Dim xp As Variant
Dim pt1(0 To 2) As Double
Dim pt2(0 To 2) As Double
Dim pt3(0 To 2) As Double
Dim pt4(0 To 2) As Double
Dim zero(0 To 2) As Double
Dim spacing As Double
Dim n As Integer
Dim r As Double
xp = ThisDrawing.Utility.GetPoint(, "Click at First point")
pt1(0) = xp(0)
pt1(1) = xp(1)
pt1(2) = xp(2)
xp = ThisDrawing.Utility.GetPoint(pt1, "Click at Second point along the slope")
pt2(0) = xp(0)
pt2(1) = xp(1)
pt2(2) = xp(2)
xp = ThisDrawing.Utility.GetPoint(pt1, "Click at Point upto where the slope is needed")
pt3(0) = xp(0)
pt3(1) = xp(1)
pt3(2) = xp(2)
xp = ThisDrawing.Utility.GetPoint(pt3, "Click at Last point")
pt4(0) = xp(0)
pt4(1) = xp(1)
pt4(2) = xp(2)
spacing = ThisDrawing.Utility.GetDistance(, "Spacing Distance")
getDifference pt2, pt1, pt2
getDifference pt3, pt1, pt3
getDifference pt4, pt1, pt4
Dim strL As Double
Dim endL As Double
Dim midL As Double
Dim strP(0 To 2) As Double
Dim endP(0 To 2) As Double
Dim tmp As AcadObject
Dim blk As AcadBlock
Set blk = ThisDrawing.Blocks.Add(zero, RandomString(6))
Set tmp = blk.AddLine(zero, pt2)
strL = tmp.Length
tmp.Delete
Set tmp = blk.AddLine(zero, pt3)
n = Fix(tmp.Length / spacing)
tmp.Delete
Set tmp = blk.AddLine(pt3, pt4)
endL = tmp.Length
tmp.Delete
Dim tempAr(0 To 2) As Double
getDifference pt4, pt3, tempAr
endL = getDotProduct(pt2, tempAr) / getMagnitude(pt2)
For i = 1 To n
    r = (strL - endL) / strL * i / n
    divideRatio i / n, zero, pt3, strP 'sets startpoint
    getAddition strP, pt2, endP ' sets endpoint
    If i Mod 2 = 1 Then divideRatio (1 - r) / 2, strP, endP, endP Else divideRatio 1 - r, strP, endP, endP
    blk.AddLine strP, endP
Next
Set tmp = ThisDrawing.ModelSpace.InsertBlock(pt1, blk.Name, 1#, 1#, 1#, 0)
tmp.Explode
blk.Delete
End Sub
