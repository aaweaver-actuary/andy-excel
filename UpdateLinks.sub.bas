'''
'===================================================================================================================
'============== FileUpdate =========================================================================================
'===================================================================================================================
' ** Note that I switch back and forth between the terms "macro" and "subprocedure" in this file. A macro is a
' subprocedure that is called from the macro dialog box in Excel. Subprocedure is a more general term that is 
' used to refer to any BASIC procedure that is not a function. In this file, I use the term "macro" to refer to
' the same concept as "subprocedure" because the procedures in this file are all called from the macro dialog
' box in Excel. **

' This module contains the `UpdateLinks` subprocedure, which is the main procedure for updating external links in
' the active Excel workbook. It also contains helper functions and subprocedures that are used by `UpdateLinks`.

' The macro is called from the macro dialog box in Excel. When you run the macro:
'   1. The user is prompted to enter find/replace text pairs that will be used to update links in the active
'      Excel workbook. The user can enter as many find/replace text pairs as they wish. The user can also enter
'      the word 'quit' as the find text to quit the loop of adding find/replace text pairs.

'      The macro builds two arrays: one with the old link and one with the new link. The old link is the
'      current link, and the new link is the current link with the find text replaced by the replace text
'      in the same order as the find/replace text pairs were entered by the user.

'   2. The macro loops through all external links in the active Excel workbook. For each external link:
'       a. The macro opens the workbook at the link in read-only mode.
'       b. The macro changes the link from the old file to the new file.
'       c. The macro closes the workbook that was just used to update the link.
'       d. The macro then moves to the next link that was found in the original Excel workbook.

'   3. If there are no links or some other error occurs, the macro displays a message box indicating that
'      no links were found or that an error occurred and quits.
'''


'These are some globally-scoped variables to hold the find/replace text. These are used in the
'`UpdateLinks` subprocedure, but are populated by the private `AddFindReplaceText` subprocedure.
Dim findText, replaceText As Variant

'''
'===================================================================================================================
'============== IsInArray ==========================================================================================
'===================================================================================================================
' `IsInArray` is a helper function that checks if a value is in an array.
'
' Parameters
' ----------
' valToBeFound : Variant
'     The value to be found in the array.
' arr : Variant
'     The array to be searched.

' Returns
' -------
' Boolean
'     True if the value is in the array, False otherwise.
'''
Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    On Error GoTo IsInArrayError: ' if valToBeFound is not found in arr then an error occurs
    IsInArray = Application.WorksheetFunction.Match(valToBeFound, arr, 0)
    Exit Function
IsInArrayError:
    On Error GoTo 0
    IsInArray = False
End Function

'''
'===================================================================================================================
'============== AddFindReplaceText =================================================================================
'===================================================================================================================
' `AddFindReplaceText` is a private subprocedure that prompts the user to enter find/replace text pairs that will be
' used to update links in the active Excel workbook.

' The procedure initializes a `userHasQuit` flag (set as `False` initially) and a counter variable `t` (set as
' 0 initially). Then, it enters a loop which runs until the `userHasQuit` flag becomes `True`.

' Within this loop:

' 1. The `findText` and `replaceText` arrays are resized to accommodate the new find/replace text pair.

' 2. The user is prompted to enter the find text. If it's the first iteration (eg if `t=0`), the prompt
' does not include the option to quit. From the second iteration onwards, the user is told to enter 'quit'
' if they wish to quit, meaning if they are done entering find/replace text pairs.

' 3. If the user enters 'quit' and it's not the first iteration, the procedure sets the corresponding replace
' text as 'quit', sets the `userHasQuit` flag as `True`, and breaks out of the loop. Otherwise, it prompts the
' user to enter the replace text.

' 4. The counter variable `t` is incremented by 1 to move to the next find/replace pair.

' After the user quits the loop, the `findText` and `replaceText` arrays are filled with find/replace text pairs.
' These arrays are used in the main `UpdateLinks` Sub procedure to perform the find/replace operation on all
' external links in the active workbook.

' Note that this procedure uses the `ReDim` statement to resize the arrays. This statement is used to resize
' an array that has already been declared. The `ReDim Preserve` statement is used to resize an array while
' preserving the existing values in the array. This is necessary because the `findText` and `replaceText`
' arrays are resized each time the user enters a new find/replace text pair, and we want to preserve the
' existing values in the array.
'''
Private Sub AddFindReplaceText()
    Dim userHasQuit, hasOneTextBox As Boolean
    Dim t As Integer

    'Initialize variables
    userHasQuit = False ' Flag to indicate if the user has quit the loop of adding find/replace text
    hasOneTextBox = False ' Flag to indicate if the user has seen at least one text box
    t = 0 ' Counter for the find/replace text

    'Prompt the user for find/replace text
    Do Until userHasQuit = True

      'Resize the arrays to hold the new find/replace text
      If IsEmpty(findText) Then
        ReDim findText(0 To t)
      Else
        ReDim Preserve findText(0 To t)
      End If
      If IsEmpty(replaceText) Then
        ReDim replaceText(0 To t)
      Else
        ReDim Preserve replaceText(0 To t)
      End If

      'Prompt the user for find/replace text
      If findText(0) = "" And hasOneTextBox Then
        userHasQuit = True
      ElseIf t = 0 Then ' First time through, don't prompt to quit
        findText(t) = InputBox("Enter the text to find. Note: this is your only opportunity to include the word 'quit'.", "Find Text")
      Else ' After the first time, prompt to quit
        findText(t) = InputBox("Enter the text to find, or 'quit' to quit", "Find Text")
      End If

      ' check the lowercase version of the `findText` for "quit", and ensure `t` > 1
      If LCase(findText(t)) = "quit" And t > 0 Then
        replaceText(t) = findText(t) ' Set the `replaceText` to "quit" so that the loop will quit
        userHasQuit = True
      Else
        replaceText(t) = InputBox("Enter the text to replace", "Replace Text")
        t = t + 1
      End If
    Loop
End Sub

'''
'===================================================================================================================
'============== UpdateSingleWorkbook ===============================================================================
'===================================================================================================================

' ` UpdateSingleWorkbook` is a private subprocedure that is responsible for opening a workbook at the specified link
' and updating the old link to the new link in the active Excel workbook.

' This procedure is called by the `UpdateLinks` subprocedure for each external link that needs to be updated.
' It receives two arguments: `oldLink` and `newLink`, which are the original and updated links, respectively. It
' also receives a reference to the result string variable to record the result of the operation and the current
' workbook object.

' This procedure attempts to:

' 1. Open the new link (i.e., the modified link) as an Excel workbook.
' 2. If the workbook opens successfully:
'   a. It checks if the `oldLink` exists in the current workbook's links. If so, it replaces the old link
'      with the new link.
'   b. It then closes the newly opened workbook.
'   c. The result of this operation is recorded as "Updated Successfully". If the old link was not found, it records
'      "Old Link Not Found"
' 3. If there is an error opening the workbook, it records the result as "Error Opening Workbook".

' It is worth noting that this subroutine uses error handling to open the workbook and handles any errors
' by recording the result as "Error Opening Workbook" and setting the workbook object to Nothing.

' Finally, this procedure modifies the `result` variable with the result of the operation. This allows the
' `UpdateLinks` procedure to track the result of each link update.
'''
Private Sub UpdateSingleWorkbook(ByVal oldLink As String, ByVal newLink As String, ByRef result As String, ByVal curWB As Workbook)
    Dim wb As Workbook
    Dim links As Variant
    
    'Try to open the new workbook
    On Error Resume Next
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Set wb = Workbooks.Open(newLink, False, True)
    DoEvents
    Application.DisplayAlerts = True
    
    'If an error occurred, handle it
    If Err.Number <> 0 Then
        'Store the error number and description
        Dim errNumber As Long
        Dim errDescription As String
        errNumber = Err.Number
        errDescription = Err.Description
        
        'Clear the error and set the result to "Error Opening Workbook"
        Err.Clear
        result = "Error Opening Workbook: Error " & errNumber & " - " & errDescription

        'Close the new workbook after the ChangeLink operation
        Set wb = Nothing
    Else
        'If no error occurred, reset the error handler and update the link
        On Error GoTo 0
        
        'Get all links
        links = curWB.LinkSources(xlLinkTypeExcelLinks)
        
        'Only try to change the link if oldLink exists in links
        If Not IsEmpty(links) Then
            If IsInArray(oldLink, links) Then
                curWB.ChangeLink oldLink, newLink, xlLinkTypeExcelLinks
                result = "Updated Successfully"
            Else
                result = "Old Link Not Found"
            End If
        Else
            result = "No Links In Workbook"
        End If
    End If
    
    'Close the new workbook after the ChangeLink operation
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If

    Application.EnableEvents = True
End Sub

'''
'===================================================================================================================
'============== UpdateLinks ========================================================================================
'===================================================================================================================

' `UpdateLinks` is the main subprocedure from this script, and is responsible for updating the links in an active
' Excel workbook.

' This procedure calls the `AddFindReplaceText` Sub procedure to collect the find/replace text pairs from the user,
' where the user is repeatedly prompted to enter find/replace text pairs until they input "quit".

' If no find/replace text pairs are entered by the user, the procedure ends prematurely with a message to the user.
' If find/replace text pairs are provided, it gets all external links in the active workbook and then performs a
' find/replace operation on each of these links based on the user-provided text.

' For each external link in the active workbook, this procedure:

' 1. Iterates through the find/replace text pairs to modify each link.
' 2. Calls the `UpdateSingleWorkbook` Sub procedure to attempt to open the updated link, replace the old link
'    in the active workbook with the new link, and close the newly opened workbook. The result of this operation
'    is recorded.

' The original link, the updated link, and the result of each operation are stored in `resOld`, `resNew`, and
' `resMsg` arrays.

' Finally, this procedure creates a new worksheet in the active workbook named "VbaLinkUpdate". If a worksheet with
' this name already exists, it is deleted before the new one is created. This new sheet contains a table with three
' columns: "Original Link", "Updated Link", and "Result", populated with the data from `resOld`, `resNew`, and
' `resMsg` arrays.
'''
Sub UpdateLinks()
    Dim wb, curWB, persWB As Workbook
    Dim curWBName, persWBName As String
    Dim oldLink, newLink, result As String
    Dim i As Long
    Dim links, allLinks As Variant
    Dim resOld, resNew, resMsg As Variant
    
    Set curWB = ActiveWorkbook
    curWBName = ActiveWorkbook.Name
    Set persWB = Workbooks("PERSONAL.XLSB")
    persWBName = persWB.Name

    'Get the find/replace text -- see AddFindReplaceText above
    Call AddFindReplaceText

    
    'Check if the user didn't enter any find/replace text
    If findText(0) = "" Then
        MsgBox "No find/replace text entered, skipping link update."
        Exit Sub
    End If
    
    'Otherwise, proceed with the link update:
    
    'Get all external links
    links = ActiveWorkbook.LinkSources(xlExcelLinks)
    
    'Loop through all links
    For i = 1 To UBound(links)
        oldLink = links(i)
        newLink = oldLink ' Reset the newLink variable
        
        'Do find/replace on the string
        For j = 0 To UBound(findText)
          newLink = Replace(newLink, findText(j), replaceText(j))
        Next j
        
        'Update the link
        Call UpdateSingleWorkbook(oldLink, newLink, result, curWB)
        
        'Add the result to the results array
        If i > 1 Then
            ReDim Preserve resOld(0 To i - 1)
            ReDim Preserve resNew(0 To i - 1)
            ReDim Preserve resMsg(0 To i - 1)
        Else
            ReDim resOld(0 To 0)
            ReDim resNew(0 To 0)
            ReDim resMsg(0 To 0)
        End If
        
        'Update the value of the result arrays
        resOld(i - 1) = oldLink
        resNew(i - 1) = newLink
        resMsg(i - 1) = result
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
        .Range("A2").Resize(UBound(resOld, 1)).Value = Application.Transpose(resOld)
        .Range("B2").Resize(UBound(resNew, 1)).Value = Application.Transpose(resNew)
        .Range("C2").Resize(UBound(resMsg, 1)).Value = Application.Transpose(resMsg)
    End With
End Sub
