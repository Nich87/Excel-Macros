Sub FlipRows()
Dim Rng As Range
Dim WorkRng As Range
Dim Arr As Variant
Dim i As Integer, j As Integer, k As Integer
On Error Resume Next
xTitleId = "順序を入れ替える範囲を選択..."
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("h", xTitleId, WorkRng.Address, Type:=8)
Arr = WorkRng.Formula
For j = 1 To UBound(Arr, 2)
    k = UBound(Arr, 1)
    For i = 1 To UBound(Arr, 1) / 2
        xTemp = Arr(i, j)
        Arr(i, j) = Arr(k, j)
        Arr(k, j) = xTemp
        k = k - 1
    Next
Next
WorkRng.Formula = Arr
End Sub
