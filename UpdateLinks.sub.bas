'Globally-scoped variables to hold the find/replace text. These are used in the UpdateLinks Sub procedure,
'but are populated by the private AddFindReplaceText Sub procedure.
Dim findText, replaceText As Variant

'Helper function to check if a value is in an array
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

      ' check the lowercase version of the find text for "quit", and ensure t > 1
      If LCase(findText(t)) = "quit" And t > 0 Then
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
'============== UpdateSingleWorkbook ===============================================================================
'===================================================================================================================

' UpdateSingleWorkbook is a private Sub procedure that is responsible for opening a workbook at the specified link
' and updating the old link to the new link in the active Excel workbook.

' This procedure is called by the UpdateLinks Sub procedure for each external link that needs to be updated.
' It receives two arguments: oldLink and newLink, which are the original and updated links, respectively.

' This procedure attempts to:

'     1. Open the new link (i.e., the modified link) as an Excel workbook.
'     2. If the workbook opens successfully:
'           a. It replaces the old link with the new link in the active workbook.
'           b. It then closes the newly opened workbook.
'           c. The result of this operation is recorded as "Updated Successfully".
'     3. If there is an error opening the workbook, it records the result as "Error Opening Workbook".

' It is worth noting that this subroutine uses error handling to open the workbook and handles any errors by recording
' the result as "Error Opening Workbook" and setting the workbook object to Nothing.

' Finally, this procedure modifies the 'result' variable with the result of the operation. This allows the
' UpdateLinks procedure to track the result of each link update.
'''
Private Sub UpdateSingleWorkbook(ByVal oldLink As String, ByVal newLink As String, ByRef result As String, ByVal curWBName As String, ByVal persWBName As String)
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
        
        'Clear the error
        Err.Clear
        result = "Error Opening Workbook: Error " & errNumber & " - " & errDescription
        Set wb = Nothing
    Else
        'If no error occurred, reset the error handler and update the link
        On Error GoTo 0
        
        'Get all links
        links = ActiveWorkbook.LinkSources(xlLinkTypeExcelLinks)
        
        'Only try to change the link if oldLink exists in links
        If Not IsEmpty(links) Then
            If IsInArray(oldLink, links) Then
                ActiveWorkbook.ChangeLink oldLink, newLink, xlLinkTypeExcelLinks
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
    
    ' BEFORE LEAVING THIS sub, close all workbooks that are not the current WB
    Dim wbCount As Integer
    wbCount = Workbooks.Count
    For i = wbCount To 1 Step -1
        If Workbooks(i).Name <> curWBName And Workbooks(i).Name <> persWBName Then
            Workbooks(i).Close
        End If
    Next i
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
    Dim wb, curWB, persWB As Workbook
    Dim curWBName, persWBName As String
    Dim oldLink, newLink, result As String
    Dim i As Long
    Dim links, allLinks As Variant
    Dim resOld, resNew, resMsg As Variant
    
    Set curWB = ActiveWorkbook
    curWBName = ActiveWorkbook.Name
    Set persWB = Workbooks("PERSONAL.XLSB")
    newLink = persWB.Name

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
        
        If i = 1 Then
            MsgBox "newLink: " & newLink
        End If

        'Update the link
        Call UpdateSingleWorkbook(oldLink, newLink, result, curWBName, persWBName)
        
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

Sub testUpdateLinks()
UpdateSingleWorkbook "O:\STAFFHQ\SYMDATA\Actuarial\Reserving Applications\IBNR Allocation\2Q2022 Analysis\Allied 2Q2022 (not analyzed).xlsb", "O:\STAFFHQ\SYMDATA\Actuarial\Reserving Applications\IBNR Allocation\3Q2022 Analysis\Allied 3Q2022.xlsb", "result", ThisWorkbook.Name, "PERSONAL.XLSB"
End Sub
