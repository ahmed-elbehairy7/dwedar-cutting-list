Attribute VB_Name = "testing"
Sub test()
Attribute test.VB_ProcData.VB_Invoke_Func = " \n14"
    Call testRowsTypes
End Sub
Public Sub testRowsTypes()
    Dim x As Integer
    x = StartingPoint
    Dim cellValue As String
       
    Do Until Cells(x, 1).Value = ""
        cellValue = Cells(x, 1).Value
        
        If cellValue = inputRow Or IsNumeric(cellValue) Then
            Call testInputRow(x)
        ElseIf cellValue = nullRow Then
            Call testNullRow(x)
        ElseIf cellValue = oTitlesRow Then
            Call testOTitlesRow(x)
        ElseIf cellValue = oNormalRow Then
            Call testONormalRow(x)
        ElseIf cellValue = oVisibleRow Then
            Call testOVisibleRow(x)
        ElseIf cellValue = oBackRow Then
            Call testOBackRow(x)
        ElseIf cellValue = endRow Then
            Call testEndRow(x)
        End If
        If Cells(1, 1).Value = 1 Then
            Cells(x, 1).Select
            Dim continueOrNot As String
            
            continueOrNot = InputBox("Row " & x & " has a problem" & _
            vbCrLf & "Do you wish to continue?" & vbCrLf _
            & "(y) : continue" & vbCrLf & "(n) : exit")
            If continueOrNot = "y" Then
            Else
                Exit Sub
            End If
        End If
        x = x + 1
    Loop
End Sub
Sub testInputRow(x As Integer)
    Dim i As Integer
    Dim returnNum As Integer
    returnNum = 1
    For i = 3 To 8
        If Cells(x, i).Interior.Color = 14348258 Then
            returnNum = 0
        End If
    Next i
    returnValue (returnNum)
End Sub
Sub testNullRow(x As Integer)
    Dim i As Integer
    Dim returnNum As Integer
    returnNum = 0
    For i = 3 To 8
        If Cells(x, i).Interior.Color <> 16777215 Then
            returnNum = 1
        End If
    Next i
    returnValue (returnNum)
End Sub
Sub testOTitlesRow(x As Integer)
    Dim titles As New Collection
    titles.Add ("„")
    titles.Add ("Œ«„…")
    titles.Add ("Ê’›")
    titles.Add ("ÿÊ· ")
    titles.Add ("⁄—÷")
    titles.Add ("⁄œœ")
    titles.Add ("„Ê«œ")
    titles.Add ("« Ã«Â «·ﬁ‘—…")
    
    Dim returnNum As Integer
    returnNum = 0
    Dim y As Integer
    For y = 3 To 8
        If Cells(x, y).Value <> titles.Item(y - 2) Then
            returnNum = 1
        End If
    Next y
    
    returnValue (returnNum)
End Sub
Sub testONormalRow(x As Integer)
    Dim i As Integer
    Dim returnNum As Integer
    returnNum = 0
    If IsNumeric(Cells(x, 2).Value) Or Cells(x, 2).Value = "" Then
        returnNum = 1
    End If
    For i = 3 To 7
        If Cells(x, i).Interior.Color <> 16777215 _
        And Cells(x, i).Interior.Color <> 15132391 _
        And Cells(x, i).Interior.Color <> 15921906 Then
            returnNum = 1
        End If
    Next i
    returnValue (returnNum)
End Sub
Sub testOVisibleRow(x As Integer)
    Dim i As Integer
    Dim returnNum As Integer
    returnNum = 0
    If IsNumeric(Cells(x, 2).Value) Or Cells(x, 2).Value = "" Then
        returnNum = 1
    End If
    For i = 3 To 7
        If Cells(x, i).Interior.Color <> 15917529 Then
            returnNum = 1
        End If
    Next i
    returnValue (returnNum)
End Sub
Sub testOBackRow(x As Integer)
    Dim i As Integer
    Dim returnNum As Integer
    returnNum = 0
    If IsNumeric(Cells(x, 2).Value) Or Cells(x, 2).Value = "" Then
        returnNum = 1
    End If
    For i = 3 To 7
        If Cells(x, i).Interior.Color <> 15592941 Then
            returnNum = 1
        End If
    Next i
    returnValue (returnNum)
End Sub
Sub testEndRow(x As Integer)
    Dim i As Integer
    Dim returnNum As Integer
    returnNum = 0
    For i = 2 To 7
        If Cells(x, i).Interior.Color <> 16777215 Or _
        Cells(x, i).Value <> "" Then
            returnNum = 1
        End If
    Next i
    returnValue (returnNum)
End Sub
