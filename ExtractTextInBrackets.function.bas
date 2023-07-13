Function ExtractTextInBrackets(ByVal text As String) As String
    Dim regex As Object
    Dim matches As Object

    ' Create regular expression object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False
        .MultiLine = True
        .IgnoreCase = False
        .Pattern = "\[(.*?)\]"
    End With
    
    ' Execute the regular expression
    Set matches = regex.Execute(text)
    
    ' If there is a match, return it
    If matches.Count > 0 Then
        ExtractTextInBrackets = matches(0).SubMatches(0)
    Else
        ExtractTextInBrackets = ""
    End If
End Function

Function ExtractTextInBrackets_Test(Optional testString As String = "'c:\users\aw\[book1.xlsb]Sheet1'!A1") As String
    Dim testString As String
    Dim result As String

    result = ExtractTextInBrackets(testString)

    TestExtractTextInBrackets = result
End Function