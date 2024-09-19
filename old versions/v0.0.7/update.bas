Attribute VB_Name = "update"
Sub update()
     
    Dim sheet As Worksheet
    
    For Each sheet In Worksheets
        If sheet.index > 2 Then
            
            sheet.Select
            Call updateBlockTypesIds
            Call updateTablesDeletion
            Call updateRowTypes
            Call updateLatestIndex
            Call updateBlockIds
        End If
    Next sheet
End Sub
Public Sub updateBlockIds()
    Dim x As Integer
    x = StartingPoint
    Do Until Range("A" & x).Value = ""
        If Not IsNumeric(Range("A" & x).Value) Then
            Range("J" & x).Value = "=letterToNumber(A" & x & ")"
        End If
        x = x + 1
    Loop
End Sub
Public Sub updateBlockTypesIds()
    Dim x As Integer
    x = StartingPoint
    
    Do Until Range("A" & x).Value = ""
        If IsNumeric(Range("A" & x).Value) Then
            Range("A" & x).Value = Range("A" & x).Value
            Range("E" & x).Value = Range("E" & x).Value
            
        End If
        x = x + 1
    Loop
End Sub
Public Sub updateTablesDeletion()

    Dim x As Integer
    x = 18
    Do Until Cells(3, x).Value = ""
        x = x + 1
    Loop
    
    Range("R3:" & numberToLetter(x) & "4").ClearContents
End Sub
Public Sub updateLatestIndex()
    
    Dim x As Integer
    x = StartingPoint
    Dim lastIndex As Integer
    lastIndex = 0
    Do Until Range("A" & x).Value = ""
        If Not IsNumeric(Range("A" & x).Value) Then
            lastIndex = letterToNumber(Range("A" & x).Value)
        End If
        x = x + 1
    Loop
    Range(lastIndexAddress).Value = lastIndex
End Sub
Public Sub updateRowTypes()

    If Range("A13").Value = "A" Then
        Exit Sub
    End If
        
    Dim x As Integer
    x = StartingPoint
    Dim cell As Range
    
    Do Until Range("A" & x).Value = ""
        Set cell = Range("A" & x)
        If IsNumeric(cell.Value) Then
            cell.Value = numberToLetter(cell.Value)
        ElseIf cell.Value = "INPUT" Then
            cell.Value = inputRow
        ElseIf cell.Value = "NULL" Then
            cell.Value = nullRow
        ElseIf cell.Value = "OTITLES" Then
            cell.Value = oTitlesRow
        ElseIf cell.Value = "ONORMAL" Then
            cell.Value = oNormalRow
        ElseIf cell.Value = "OVISIBLE" Then
            cell.Value = oVisibleRow
        ElseIf cell.Value = "OBACK" Then
            cell.Value = oBackRow
        ElseIf cell.Value = "END" Then
            cell.Value = endRow
        End If
        x = x + 1
    Loop
End Sub
