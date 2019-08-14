Attribute VB_Name = "Formatting"
Public Sub UcaseTexts()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen 'this one doesn't work as the active selection set is not empty
End If
For Each obj In ss
    If obj.ObjectName = "AcDbText" Or obj.ObjectName = "AcDbMText" Then
        obj.TextString = StrConv(obj.TextString, vbUpperCase)
    End If
Next
End Sub
Public Sub LcaseTexts()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
For Each obj In ss
    If obj.ObjectName = "AcDbText" Or obj.ObjectName = "AcDbMText" Then
        obj.TextString = StrConv(obj.TextString, vbLowerCase)
    End If
Next
End Sub
Public Sub PcaseTexts()
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
If ss.Count = 0 Or ThisDrawing.GetVariable("PICKFIRST") = 0 Then
    ss.SelectOnScreen
End If
For Each obj In ss
    If obj.ObjectName = "AcDbText" Or obj.ObjectName = "AcDbMText" Then
        obj.TextString = StrConv(obj.TextString, vbProperCase)
    End If
Next
End Sub
