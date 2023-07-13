'Globally-scoped variables to hold the find/replace text. These are used in the UpdateLinks Sub procedure,
'but are populated by the private AddFindReplaceText Sub procedure.
Dim findText(), replaceText() As Variant

'''
'===================================================================================================================
'============== AddFindReplaceText =================================================================================
'===================================================================================================================

' AddFindReplaceText is a Sub procedure that prompts the user to enter find/replace text pairs that will be used
' to update links in the active Excel workbook.

' The procedure initializes a userHasQuit flag (set as False initially) and a counter variable t (set as 1 initially).
' Then, it enters a loop which runs until the userHasQuit flag becomes True.

' Within this loop:

'     1. The findText and replaceText arrays are resized to accommodate the new find/replace text pair.

'     2. The user is prompted to enter the find text. If it's the first iteration (t=1), the prompt does not
'        include the option to quit. From the second iteration onwards, the user is told to enter 'quit' if they
'        wish to quit.

'     3. If the user enters 'quit' and it's not the first iteration, the procedure sets the corresponding replace
'        text as 'quit', sets the userHasQuit flag as True, and breaks out of the loop. Otherwise, it prompts the
'        user to enter the replace text.

'     4. The counter variable t is incremented by 1 to move to the next find/replace pair.

' After the user quits the loop, the findText and replaceText arrays are filled with find/replace text pairs.
' These arrays are used in the UpdateLinks Sub procedure to perform the find/replace operation on all external links
' in the active workbook.
'''
Sub AddFindReplaceText()
    Dim userHasQuit As Boolean
    Dim t As Integer

    'Initialize variables
    userHasQuit = False ' Flag to indicate if the user has quit the loop of adding find/replace text
    t = 1 ' Counter for the find/replace text

    'Prompt the user for find/replace text
    Do Until userHasQuit = True

      'Resize the arrays to hold the new find/replace text
      ReDim Preserve findText(1 To t)
      ReDim Preserve replaceText(1 To t)

      'Prompt the user for find/replace text
      If findText(t) = "" Then
        userHasQuit = True
      ElseIf t = 1 Then ' First time through, don't prompt to quit
        findText(t) = InputBox("Enter the text to find. Note: this is your only opportunity to include the word 'quit'.", "Find Text")
      Else ' After the first time, prompt to quit
        findText(t) = InputBox("Enter the text to find, or 'quit' to quit", "Find Text")
      End If

      ' check the lowercase version of the find text for "quit", and ensure t > 1
      If LCase(findText(t)) = "quit" And t > 1 Then
        replaceText(t) = findText(t) ' Set the replace text to "quit" so that the loop will quit
        userHasQuit = True
      Else
        replaceText(t) = InputBox("Enter the text to replace", "Replace Text")
        t = t + 1
      End If
    Loop
End Sub


'''
'===================================================================================================================
'============== UpdateLinks ========================================================================================
'===================================================================================================================

' UpdateLinks is a Sub procedure that is responsible for updating the links in an active Excel workbook.

' This procedure gets all external links in the active workbook and then performs a find/replace operation
' on each of these links based on user-provided text.

' The find/replace text pairs are collected from the user by calling the AddFindReplaceText Sub procedure,
' where the user is repeatedly prompted to enter find/replace text pairs until they input "quit".

' For each external link in the active workbook, this procedure attempts to:

'     1. Replace the find text in the link with the replace text.
'     2. Open the new link (i.e., the modified link) as an Excel workbook.
'     3. If the workbook opens successfully, it replaces the old link with the new link in the active workbook,
'        and then closes the newly opened workbook. The result of this operation is recorded as "Updated Successfully".
'     4. If there is an error opening the workbook, it records the result as "Error Opening Workbook".

' The original link, the updated link, and the result of each operation are stored in a 2D array.

' Finally, this procedure adds a new sheet to the active workbook named "VbaLinkUpdate". If such a sheet already exists,
' it is deleted before the new one is added. This new sheet contains a table with three columns: "Original Link",
' "Updated Link", and "Result", and each row in this table corresponds to an external link in the workbook and contains
' the data from the aforementioned 2D array.
'''
Sub UpdateLinks()
    Dim wb As Workbook
    Dim oldLink, newLink, result As String
    Dim results() As Variant
    Dim i As Long
    Dim links As Variant

    'Get the find/replace text -- see AddFindReplaceText above
    Call AddFindReplaceText
    
    'Get all external links
    links = ActiveWorkbook.LinkSources(xlExcelLinks)
    
    'Loop through all links
    For i = LBound(links) To UBound(links)
        oldLink = links(i)
        newLink = oldLink ' Reset the newLink variable
        
        'Do find/replace on the string
        If findText(1) = "" Then
            'do nothing
        Else
            For j = LBound(findText) To UBound(findText)
              newLink = Replace(newLink, findText(j), replaceText(j))
            Next j
            ' newLink = Replace(oldLink, findText, replaceText)
            
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
        End If
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
