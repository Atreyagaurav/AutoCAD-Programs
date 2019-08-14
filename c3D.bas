Attribute VB_Name = "c3D"
Public Sub testc3d()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
For Each obb In ss
    If Not (obb.station Mod 250 = 0 Or obb.station = 50) Then obb.Delete
Next
End Sub
Public Sub tempr()
Dim ss As AcadSelectionSet
Dim pt As AcadPoint
Dim ob As AcadObject
Set ss = ThisDrawing.ActiveSelectionSet
For Each obb In ss
    ThisDrawing.Layers.Add (obb.TextString)
    ThisDrawing.Utility.GetEntity ob, pt, "click at object at layer: " + obb.TextString
    ob.Layer = obb.TextString
Next
End Sub

