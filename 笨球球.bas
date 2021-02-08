Attribute VB_Name = "±ø«Ú«Ú"
Sub ±ø«Ú«Ú()

Dim rg As Range, i As Integer, nocl As Integer
Dim k As Integer, r As Integer
Dim arr

ncol = Selection.Column
nrow = Selection.Row

r = nrow
For Each rg In Selection
    k = 0
    arr = Split(CStr(rg.Value), "£ª")
    
    For i = ncol To UBound(arr)
        Cells(r, i) = arr(k)
        k = k + 1
    Next
    
    r = r + 1
Next
End Sub
