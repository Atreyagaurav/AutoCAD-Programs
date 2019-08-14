Attribute VB_Name = "InterLayout"
Public Sub NumberingLayouts()
Dim lots As AcadLayouts
Dim blo As AcadObject
Dim num As Integer
Dim str As String
Dim prop As Variant
Set lots = ThisDrawing.Layouts
ThisDrawing.Utility.GetEntity blo, "Click at Sheet block"
prop = blo.GetAttributes
For Each pro In prop
On Error Resume Next
    If pro.TagString = "DRAWING_NUMBER" Then
        num = CInt(Right(pro.TextString, 2))
        str = Left(pro.TextString, Len(pro.TextString) - 2)
    End If
Next
For Each lot In lots
    ThisDrawing.ActiveLayout = lot
    'ThisDrawing.Utility.GetEntity blo, "Click at Index block"
    'BringToFront (blo)
    ThisDrawing.Utility.GetEntity blo, "Click at Sheet block"
    num = CInt(Right(lot.Name, 2))
    prop = blo.GetAttributes
    For Each pro In prop
        If pro.TagString = "DRAWING_NUMBER" Then
        If num > 9 Then pro.TextString = str & CStr(num) Else pro.TextString = str & "0" & CStr(num)
        End If
    Next
Next
'DRAWING_NUMBER
End Sub
Public Sub CopyTextLayouts()
Dim lots As AcadLayouts
Dim tx1 As AcadObject
Dim tx2 As AcadObject
Dim ob As AcadObject
Set lots = ThisDrawing.Layouts
ThisDrawing.Utility.GetEntity tx1, "Click at Text1"
ThisDrawing.Utility.GetEntity tx2, "Click at Text1"
For Each lot In lots
    On Error Resume Next
    ThisDrawing.ActiveLayout = lot
    'ThisDrawing.Utility.GetEntity blo, "Click at Index block"
    'BringToFront (blo)
    ThisDrawing.Utility.GetEntity ob, "Click at text1"
    ob.TextString = tx1.TextString
    ThisDrawing.Utility.GetEntity ob, "Click at text2"
    ob.TextString = tx2.TextString
Next
End Sub
Public Sub makeLayouts()
'=====makes layout from the selected polylines===
Dim ss As AcadSelectionSet
Set ss = ThisDrawing.ActiveSelectionSet
Dim pr(0 To 3) As Double
Dim point1 As Variant
Dim point2 As Variant
Dim prop As Variant
Dim oLay1 As AcadLayout
Dim oLay2 As AcadLayout
Dim tempObj As AcadObject
Dim pt As Variant
Dim i As Long
Dim pre As String
Dim j As Integer
j = 0
Dim vEnts() As AcadObject
Set oLay1 = ThisDrawing.Layouts(0)
Debug.Print oLay1.Name
Debug.Print oLay1.ObjectName
pre = ThisDrawing.Utility.GetString(1, "Enter Prefix for Layouts")
For Each obj In ss
j = j + 1
    'creating the layout by copying
    Set oBlk1 = oLay1.Block
    Set oLay2 = ThisDrawing.Layouts.Add(pre + "_" + CStr(j))
Debug.Print oLay2.Name
Debug.Print oLay2.ObjectName
    oLay2.CopyFrom oLay1
    Set oBlk2 = oLay2.Block
    ReDim vEnts(0 To oBlk1.Count - 1)
    i = 0
    For Each oEnt In oBlk1
    Set vEnts(i) = oEnt
    i = i + 1
    Next
    ThisDrawing.CopyObjects vEnts, oBlk2
    'activating the modelspace
    ThisDrawing.ActiveLayout = oLay2
            'newVport.Display True
    'zooming the modelspace
    ThisDrawing.SendCommand ("_.Mspace" + Chr(10))
    obj.GetBoundingBox point1, point2
    ThisDrawing.Application.ZoomWindow point1, point2
    ThisDrawing.MSpace = False
'=====to number it accordingly
'ThisDrawing.Utility.GetEntity tempObj, pt, "Click at the Sheet Block"
'prop = tempObj.GetAttributes
'    For Each pro In prop
'        If pro.TagString = "DRAWING_NUMBER" Then
'        If j > 9 Then pro.TextString = "MCHP-TAC-TM-CA-" & CStr(j) Else pro.TextString = "MCHP-TAC-TM-CA-" & "0" & CStr(j)
'        End If
'    Next
Next
End Sub

Public Sub D_CardLayouts()
'=====makes layout from the selected cells of D_card excel sheet===
Dim vbAns As VbMsgBoxResult
 vbAns = MsgBox("Have you selected the Stations in the Excel sheet?", vbYesNo, "Check Excel")
If vbAns = vbYes Then
    Dim prop As Variant
    Dim oLay1 As AcadLayout
    Dim oLay2 As AcadLayout
    Dim tempObj As AcadObject
    Dim p(0 To 2) As Double
    Dim p_low(0 To 2) As Double
    Dim p_high(0 To 2) As Double
    Dim pt As Variant
    Dim i As Long
    Dim j As Integer
    Dim excelSht As Worksheet
    Dim excelApp As Excel.Application
    j = 0
    Dim index As Integer
    Dim vEnts() As AcadObject
    Set oLay1 = ThisDrawing.Layouts(0)
    Set excelApp = GetObject(, "excel.application")
    Set excelSht = excelApp.ActiveSheet
    For Each cel In excelApp.Selection
        'creating the layout by copying
        Set oBlk1 = oLay1.Block
        Set oLay2 = ThisDrawing.Layouts.Add(cel.Value)
        oLay2.CopyFrom oLay1
        Set oBlk2 = oLay2.Block
        ReDim vEnts(0 To oBlk1.Count - 1)
        i = 0
        For Each oEnt In oBlk1
        Set vEnts(i) = oEnt
        i = i + 1
        Next
        ThisDrawing.CopyObjects vEnts, oBlk2
        'activating the modelspace
        ThisDrawing.ActiveLayout = oLay2
    '=====to input all the details
    ThisDrawing.Utility.GetEntity tempObj, pt, "Click at the Sheet Block"
    prop = tempObj.GetAttributes
        For Each pro In prop
            For index = 0 To 12
            If excelSht.Range(Chr(Asc("A") + index) & "1").Value = pro.TagString Then
                pro.TextString = excelSht.Range(Chr(Asc("A") + index) & CStr(cel.Row)).Value
                If pro.TagString = "ORD_E" Then
                    p(0) = CDbl(pro.TextString)
                ElseIf pro.TagString = "ORD_N" Then
                    p(1) = CDbl(pro.TextString)
                'ElseIf pro.TagString = "ORD_Z" Then
                    'p(2) = CDbl(pro.TextString)
                End If
            End If
            Next
        Next
        p_low(0) = p(0) - 39.42878256
        p_low(1) = p(1) - 25.52760476
        p_high(0) = p(0) + 39.42878256
        p_high(1) = p(1) + 25.52760476
        p_low(2) = 0
        p_high(2) = 0
        'zooming the modelspace
        ThisDrawing.SendCommand ("_.Mspace" + Chr(10))
        ThisDrawing.Application.ZoomWindow p_low, p_high
        ThisDrawing.MSpace = False
    Next
End If
End Sub
Public Sub D_Card_UpadateLayout()
'=====makes layout from the selected cells of D_card excel sheet===
Dim vbAns As VbMsgBoxResult
 vbAns = MsgBox("Have you selected the Stations in the Excel sheet?", vbYesNo, "Check Excel")
If vbAns = vbYes Then
    Dim prop As Variant
    Dim p(0 To 2) As Double
    Dim tempObj As AcadObject
    Dim p_low(0 To 2) As Double
    Dim p_high(0 To 2) As Double
    Dim pt As Variant
    Dim i As Long
    Dim j As Integer
    Dim excelSht As Worksheet
    Dim excelApp As Excel.Application
    j = 0
    Dim index As Integer
    Dim vEnts() As AcadObject
    Set oLay1 = ThisDrawing.Layouts(0)
    Set excelApp = GetObject(, "excel.application")
    Set excelSht = excelApp.ActiveSheet
    For Each cel In excelApp.Selection
    'On Error Resume Next
    ThisDrawing.ActiveLayout = ThisDrawing.Layouts(cel.Value)
    '=====to input all the details
    ThisDrawing.Utility.GetEntity tempObj, pt, "Click at the D-Card Block"
    prop = tempObj.GetAttributes
        For Each pro In prop
            For index = 0 To 12
            If excelSht.Range(Chr(Asc("A") + index) & "1").Value = pro.TagString Then
                pro.TextString = excelSht.Range(Chr(Asc("A") + index) & CStr(cel.Row)).Value
                If pro.TagString = "ORD_E" Then
                    p(0) = CDbl(pro.TextString)
                ElseIf pro.TagString = "ORD_N" Then
                    p(1) = CDbl(pro.TextString)
                'ElseIf pro.TagString = "ORD_Z" Then
                    'p(2) = CDbl(pro.TextString)
                End If
            End If
            Next
        Next
        p_low(0) = p(0) - 39.42878256
        p_low(1) = p(1) - 25.52760476
        p_high(0) = p(0) + 39.42878256
        p_high(1) = p(1) + 25.52760476
        p_low(2) = 0
        p_high(2) = 0
        'zooming the modelspace
        ThisDrawing.MSpace = True
        ThisDrawing.SendCommand ("_.Mspace" + Chr(10))
        ThisDrawing.Application.ZoomWindow p_low, p_high
        ThisDrawing.MSpace = False
    Next
End If
End Sub
Public Sub CopyPlotConfig()
Dim oLay1 As AcadLayout
Dim oLay2 As AcadLayout
Set oLay1 = ThisDrawing.ActiveLayout
For Each oLay In ThisDrawing.Layouts
    If Not (oLay.Name Like "Model" Or oLay.Name = oLay1.Name) Then oLay.CopyFrom oLay1
Next
End Sub
