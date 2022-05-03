Remove Duplicates But Keep Rest Of Row Values With VBA


In Excel, there is a VBA code that also can remove duplicates but keep rest of row values.

1. Press Alt + F11 keys to display Microsoft Visual Basic for Applications window.

2. Click Insert > Module, and paste below code to the Module.

VBA: Remove duplicates but keep rest of row values


Sub RemoveDuplicates()
'UpdatebyExtendoffice20160918

    Dim xRow As Long
    Dim xCol As Long
    Dim xrg As Range
    Dim xl As Long
    On Error Resume Next
    Set xrg = Application.InputBox("Select a range:", "Kutools for Excel", _
                                    ActiveWindow.RangeSelection.AddressLocal, , , , , 8)

    xRow = xrg.Rows.Count + xrg.Row - 1
    xCol = xrg.Column
    'MsgBox xRow & ":" & xCol
    Application.ScreenUpdating = False
    For xl = xRow To 2 Step -1
        If Cells(xl, xCol) = Cells(xl - 1, xCol) Then
            Cells(xl, xCol) = ""
        End If
    Next xl
    Application.ScreenUpdating = True
    
End Sub

3. Press F5 key to run the code, a dialog pops out to remind you to select a range to remove duplicate values from. See screenshot:
4. Click OK, now the duplicate values have been removed from selection and leave blank cells.

 Remove Duplicates But Keep Rest Of Row Values With Kutools For Excel

From:
https://www.extendoffice.com/documents/excel/4043-excel-remove-duplicate-value-but-keep-rest-of-the-row-values.html
