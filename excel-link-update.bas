Sub UpdateLinks()
    Dim wb As Workbook
    Dim oldLink, newLink, result, findText, replaceText As String
    Dim results() As Variant
    Dim i As Long
    Dim links As Variant

    'Prompt the user for find/replace text
    findText = InputBox("Enter the text to find", "Find Text")
    replaceText = InputBox("Enter the text to replace", "Replace Text")
    
    'Get all external links
    links = ActiveWorkbook.LinkSources(xlExcelLinks)
    
    'Loop through all links
    For i = LBound(links) To UBound(links)
        oldLink = links(i)
        
        'Do find/replace on the string
        newLink = Replace(oldLink, findText, replaceText)
        
        'Try to open the new workbook
        On Error Resume Next
        Set wb = Workbooks.Open(newLink)
        If Err.Number <> 0 Then
            Err.Clear
            result = "Error Opening Workbook"
            Set wb = Nothing
        Else
            On Error GoTo 0
            'Change the link
            ActiveWorkbook.ChangeLink oldLink, newLink, xlLinkTypeExcelLinks
            wb.Close SaveChanges:=False
            result = "Updated Successfully"
        End If
        
        'Add the result to the results array
        ReDim Preserve results(1 To 3, 1 To i)
        results(1, i) = oldLink
        results(2, i) = newLink
        results(3, i) = result
    Next i
    
    'Remove the old sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("VbaLinkUpdate").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    'Create a new sheet for the results
    With ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        .Name = "VbaLinkUpdate"
        .Range("A1:C1").Value = Array("Original Link", "Updated Link", "Result")
        .Range("A2").Resize(UBound(results, 2), UBound(results, 1)).Value = Application.Transpose(results)
    End With
End Sub