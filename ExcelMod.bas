Attribute VB_Name = "ExcelMod"
Public Sub TextToExcel()
Dim ob As AcadObject
Dim ob2 As AcadObject
Dim pt As AcadPoint
Dim ord As Variant
Dim app As Excel.Application
Dim shtobj As Excel.Worksheet
Dim sheetname As String
Dim i As Integer
Set app = GetObject(, "Excel.Application")
sheetname = app.ActiveSheet.Name
Set shtobj = app.Worksheets(sheetname)
On Error GoTo en
i = 0
While True:
ThisDrawing.Utility.GetEntity ob, pt, "Click at a block"
ThisDrawing.Utility.GetEntity ob2, pt, "Click at a text"
ord = ob.InsertionPoint
app.ActiveCell.Offset(i, 0).Range("A1").Value = ob.Rotation
app.ActiveCell.Offset(i, 1).Range("A1").Value = ob2.TextString
app.ActiveCell.Offset(i, 2).Range("A1").Value = ord(0)
app.ActiveCell.Offset(i, 3).Range("A1").Value = ord(1)
ob.Delete
ob2.Delete
i = i + 1
'shtobj.Cells(i, 1).Value = ob.TextString
Wend
en:
End Sub
