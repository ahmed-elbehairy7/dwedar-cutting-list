Attribute VB_Name = "update"
Sub updateAll()
    Dim sheet As Worksheet
    Application.ScreenUpdating = False
    For Each sheet In Worksheets
        If sheet.index = 3 _
        Or sheet.index = 1 _
        Or sheet.index = 2 Then
        Else
            sheet.Select
            Call updateSheet
        End If
    Next sheet
    Application.ScreenUpdating = True
End Sub
Sub updateSheet()
    
    Dim x As Integer
    x = StartingPoint

    Dim cellValue As Variant
    Dim value As String
    
    Do Until Range("A" & x).value = ""
        cellValue = Range("A" & x).value
        If cellValue = oNormalRow _
        Or cellValue = oVisibleRow _
        Or cellValue = oBackRow Then
            If rowHasIf(x) Then
                MsgBox "found " & ActiveSheet.Name & ": " & x
                Exit Sub
            End If
        End If
        x = x + 1
    Loop
    
End Sub
Function rowHasIf(rowNum As Integer)

    Dim col As Integer
    Dim cellValue As String
    
    For col = 5 To 8
        cellValue = Range(numberToLetter(col) & rowNum).formula
        If InStr(1, cellValue, "IF(") Then
            rowHasIf = True
            Exit Function
        End If
    Next col
    
    rowHasIf = False
End Function
