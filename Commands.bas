Attribute VB_Name = "Commands"
Public Sub Hmove()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
Dim spt As Variant
Dim ept As Variant
spt = ThisDrawing.Utility.GetPoint(, "Click at base point")
ept = ThisDrawing.Utility.GetPoint(, "Click at second point")
ept(1) = spt(1)
ept(2) = spt(2)
For Each obj In ss
    obj.Move spt, ept
Next
End Sub
Public Sub Vmove()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
Dim spt As Variant
Dim ept As Variant
spt = ThisDrawing.Utility.GetPoint(, "Click at base point")
ept = ThisDrawing.Utility.GetPoint(, "Click at second point")
ept(0) = spt(0)
ept(2) = spt(2)
For Each obj In ss
    obj.Move spt, ept
Next
End Sub
Public Sub Zmove()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
Dim spt As Variant
Dim ept As Variant
spt = ThisDrawing.Utility.GetPoint(, "Click at base point")
ept = ThisDrawing.Utility.GetPoint(, "Click at second point")
ept(0) = spt(0)
ept(1) = spt(1)
For Each obj In ss
    obj.Move spt, ept
Next
End Sub
Public Sub Red()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
For Each obj In ss
    obj.color = acRed
Next
End Sub
Public Sub OverRideDIM()
Dim ss As AcadSelectionSet
Dim str As String
Dim pre As Double
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
pre = ThisDrawing.Utility.GetReal("Enter precision:")
For Each obj In ss
    If obj.ObjectName Like "*Dimension" Then
    str = CStr(Fix(obj.Measurement / pre + 0.5) * pre * 100)
    If Fix(CInt(str) / 100) = 0 Then str = "0" + str
    obj.TextOverride = Left(str, Len(str) - 2) + "." + Right(str, 2)
    End If
Next
End Sub
