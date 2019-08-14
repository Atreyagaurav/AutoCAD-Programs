Attribute VB_Name = "SlopePattern"

Public Sub DrawSlopePattern()
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
ThisDrawing.ModelSpace.InsertBlock pt1, blk.Name, 1#, 1#, 1#, 0
End Sub
Private Function ratio(r As Double, x1 As Double, x2 As Double) As Double
ratio = x1 + r * (r2 - r1)
End Function

Function RandomString(Length As Integer)
'PURPOSE: Create a Randomized String of Characters
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
Dim CharacterBank As Variant
Dim x As Long
Dim str As String

'Test Length Input
  If Length < 1 Then
    MsgBox "Length variable must be greater than 0"
    Exit Function
  End If

CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
  "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "A", "B", "C", "D", "E", "F", "G", "H", _
  "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
  "W", "X", "Y", "Z")

'Randomly Select Characters One-by-One
  For x = 1 To Length
    Randomize
    str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
  Next x

'Output Randomly Generated String
  RandomString = str

End Function
