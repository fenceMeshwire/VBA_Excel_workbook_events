Option Explicit

' ________________________________________________________________________________________
' Description:
' Place the following code snippet in "ThisWorkbook" in order to select "Sheet1"
' when opening the corresponding XLSM file, s. code below.
' ________________________________________________________________________________________

Private Sub Workbook_Open()
    Sheet1.Activate
End Sub
