Attribute VB_Name = "Blocks"
'===========================================SOP BLOCKS====================================
Public Sub SOP()
Dim org_ob As AcadObject
Dim ob As AcadObject
Dim prop As Variant
Dim p1 As Variant
Dim p2 As Variant
Dim org_str As String
Dim n As Integer
ThisDrawing.Utility.GetEntity org_ob, p1, "Click at the Correct SOP Block"
prop = org_ob.GetAttributes
p1 = org_ob.InsertionPoint
org_str = ThisDrawing.Utility.GetString(0, "Enter Prefix:")
n = CInt(ThisDrawing.Utility.GetString(0, "first number"))
While True
On Error GoTo las:
p2 = ThisDrawing.Utility.GetPoint(p1, "Click at the Points")
Set ob = org_ob.Copy
ob.InsertionPoint = p2
prop = ob.GetAttributes
For Each pr In prop
If pr.TagString = "ELV" Then
pr.TextString = org_str & " " & CStr(n)
End If
Next
ob.Update
n = n + 1
Wend
las:
End Sub
Public Sub SOPExisting()
Dim org_ob As AcadObject
Dim ob As AcadObject
Dim prop As Variant
Dim p1 As Variant
Dim p2 As Variant
Dim org_str As String
Dim n As Integer
ThisDrawing.Utility.GetEntity org_ob, p1, "Click at the Correct SOP Block"
prop = org_ob.GetAttributes
p1 = org_ob.InsertionPoint
org_str = ThisDrawing.Utility.GetString(0, "Enter Prefix:")
n = CInt(ThisDrawing.Utility.GetString(0, "first number"))
While True
On Error GoTo las:
ThisDrawing.Utility.GetEntity ob, p1, "Click at the Correct SOP Block"
prop = ob.GetAttributes
For Each pr In prop
If pr.TagString = "ELV" Then
pr.TextString = org_str & " " & CStr(n)
End If
Next
ob.Update
n = n + 1
Wend
las:
End Sub
'=========================================ELEVATION BLOCK=======================================
Public Sub Elevation()
Dim org_ob As AcadObject
Dim ob As AcadObject
Dim prop As Variant
Dim p1 As Variant
Dim p2 As Variant
Dim ele As Double
Dim org_ele As Double
Dim org_y As Double
ThisDrawing.Utility.GetEntity org_ob, p1, "Click at the Correct Elevation Block"
prop = org_ob.GetAttributes
p1 = org_ob.InsertionPoint
org_y = p1(LBound(p1) + 1)
For Each pr In prop
If pr.TagString = "E" Then
org_ele = CDbl(pr.TextString)
End If
Next
While True
On Error GoTo las:
p2 = ThisDrawing.Utility.GetPoint(p1, "Click at the Points")
Set ob = org_ob.Copy
ob.InsertionPoint = p2
prop = ob.GetAttributes
For Each pr In prop
If pr.TagString = "E" Then
ThisDrawing.Utility.Prompt "The Elevation Difference: " & CStr(p2(LBound(p2) + 1) - org_y) & Chr(10)
ele = (Fix((p2(LBound(p2) + 1) - org_y + org_ele) * 1000 + 0.5))
pr.TextString = Left(CStr(ele), Len(CStr(ele)) - 3) & "." & Right(CStr(ele), 3)
End If
Next
ob.Update
Wend
las:
End Sub
Public Sub Elevation_ExistingBlock()
Dim org_ob As AcadObject
Dim ob As AcadObject
Dim prop As Variant
Dim p1 As Variant
Dim p2 As Variant
Dim ele As Double
Dim org_ele As Double
Dim org_y As Double
ThisDrawing.Utility.GetEntity org_ob, p1, "Click at the Correct Elevation Block"
prop = org_ob.GetAttributes
p1 = org_ob.InsertionPoint
org_y = p1(LBound(p1) + 1)
For Each pr In prop
If pr.TagString = "E" Then
org_ele = CDbl(pr.TextString)
End If
Next
While True
On Error GoTo las:
ThisDrawing.Utility.GetEntity ob, p2, "Click at the Elevation Block to Edit"
p2 = ob.InsertionPoint
prop = ob.GetAttributes
For Each pr In prop
If pr.TagString = "E" Then
ThisDrawing.Utility.Prompt "The Elevation Difference: " & CStr(p2(LBound(p2) + 1) - org_y) & Chr(10)
ele = (Fix((p2(LBound(p2) + 1) - org_y + org_ele) * 1000 + 0.5))
pr.TextString = Left(CStr(ele), Len(CStr(ele)) - 3) & "." & Right(CStr(ele), 3)
End If
Next
ob.Update
Wend
las:
End Sub
Public Sub EditBlock()
Dim org_ob As AcadObject
Dim prop As Variant
Dim sr As String
ThisDrawing.Utility.GetEntity org_ob, p1, "Click at the The Block"
prop = org_ob.GetAttributes
For Each pr In prop
sr = ThisDrawing.Utility.GetString(1, pr.TagString & "<default>:")
If Not (sr = "") Then pr.TextString = sr
If (sr = ".") Then pr.TextString = ""
Next
org_ob.Update
las:
End Sub
Public Sub copyBlockAttributeText()
Dim org_ob As AcadObject
Dim ob As AcadObject
Dim org_prop As Variant
Dim prop As Variant
Dim sr As String
Dim match As Boolean
ThisDrawing.Utility.GetEntity org_ob, p1, "Click at the The Original Block"
ThisDrawing.Utility.GetEntity ob, p1, "Click at the The New Block"
org_prop = org_ob.GetAttributes
prop = ob.GetAttributes
For Each pr In prop
    match = False
    For Each org_pr In org_prop
        If org_pr.TagString = pr.TagString Then
            sr = ThisDrawing.Utility.GetString(1, pr.TagString & "<copy from old>:")
            If Not (sr = "") Then pr.TextString = sr Else pr.TextString = org_pr.TextString
            match = True
        End If
    Next
    If match = False Then
        sr = ThisDrawing.Utility.GetString(1, pr.TagString & "<default>:")
        If Not (sr = "") Then pr.TextString = sr
        If (sr = ".") Then pr.TextString = ""
    End If
Next
org_ob.Update
las:
End Sub
Private Function getCenter(ob As Variant)
Dim mr As Variant
Dim pt(0 To 2) As Double
mr = ob.ControlPoints
Dim n As Integer
n = (UBound(mr) - LBound(mr) + 1) / 3
Dim x, y As Double
x = 0
y = 0
For i = 0 To n - 1
    x = x + mr(3 * i)
    y = y + mr(3 * i + 1)
Next
pt(0) = x / n
pt(1) = y / n
pt(2) = 0#
getCenter = pt
End Function
Public Sub copyToAll()
Dim ob As AcadObject
Dim ob2 As AcadObject
Dim ss As AcadSelectionSet
Dim pt As Variant
Dim pt0 As Variant
ThisDrawing.Utility.GetEntity ob, pt, "Click at the object to copy"
'ThisDrawing.Utility.GetEntity ob2, pt, "Click at the reference object"
pt0 = ob.InsertionPoint
Set ss = ThisDrawing.ActiveSelectionSet
For Each obj In ss:
    If obj.Layer = ob.Layer Then
        pt = obj.Center 'getCenter(obj)
        pt(0) = pt(0) '+ 3.5
        Set ob2 = ob.Copy
        ob2.Move pt0, pt
    End If
Next
End Sub
Public Sub HatchAll()
Dim ss As AcadSelectionSet
Dim ob As AcadHatch
Set ss = ThisDrawing.ActiveSelectionSet
Dim obb(0 To 0) As AcadObject
For Each obj In ss:
On Error Resume Next
If obj.ObjectName Like "*Polyline" Then
    Set ob = ThisDrawing.ModelSpace.AddHatch(0, "ANSI31", True)
    Set obb(0) = obj
    ob.AppendOuterLoop (obb)
    ob.Layer = obj.Layer
    ob.PatternScale = 0.1
    ob.Lineweight = acLnWt000
ElseIf obj.ObjectName Like "*Hatch" Then
    obj.Lineweight = acLnWt000
End If
Next
End Sub
Public Sub SendBlocks2Excel()
Dim ss As AcadSelectionSet
End Sub
