Attribute VB_Name = "mdl_scanDoc"
Option Explicit

' Main subroutine to initiate the process
Sub IdentifyLikeTermsAndPronouns()
    On Error GoTo ErrorHandler
    
    Dim configFilePath As String
    Dim targetDocPath As String
    Dim selectedDoc As Document
    Dim likeTermGroups As Object ' Dictionary: GroupName -> Dictionary of Terms
    Dim termCounts As Object ' Dictionary: GroupName -> Dictionary of Term Counts and Pages
    Dim pronounsList As Variant
    Dim pronounsDict As Object ' Dictionary: Pronoun -> Dictionary of Count and Pages
    Dim output As String
    Dim reportFilePath As String
    Dim numberStyleIssues As Object ' Dictionary to store number style issues
    
    
    ' Step 1: Get the configuration file path from the user
    configFilePath = GetSelectedFilePath("Select Configuration File", "Text Files (*.txt)|*.txt")
    If configFilePath = "" Then Exit Sub ' User canceled
    
    ' Step 2: Get the document path from the user
    targetDocPath = GetSelectedFilePath("Select Word Document to Analyze", "Word Documents (*.docx;*.doc)|*.docx;*.doc")
    If targetDocPath = "" Then Exit Sub ' User canceled
    
    ' Step 3: Read the configuration file
    Set likeTermGroups = ReadConfigFile(configFilePath)
    If likeTermGroups Is Nothing Then
        MsgBox "Failed to read the configuration file.", vbCritical
        Exit Sub
    End If
    
    ' Step 4: Open the selected document
    Set selectedDoc = OpenDocument(targetDocPath)
    If selectedDoc Is Nothing Then
        MsgBox "Failed to open the selected document.", vbCritical
        Exit Sub
    End If
    
    ' Initialize number style issues dictionary
    Set numberStyleIssues = CreateObject("Scripting.Dictionary")
    numberStyleIssues.Add "small_numbers", CreateObject("Scripting.Dictionary") ' For numbers 1-9
    numberStyleIssues.Add "sentence_start", CreateObject("Scripting.Dictionary") ' For numbers at sentence start
        
    ' Step 5: Initialize pronouns
    pronounsList = Array("he", "she", "it", "they", "them", "his", "her", "their", "we", "us", "I", "you", "me", "him")
    Set pronounsDict = CreateObject("Scripting.Dictionary")
    Dim pronoun As Variant
    For Each pronoun In pronounsList
        Set pronounsDict(pronoun) = CreateObject("Scripting.Dictionary")
        pronounsDict(pronoun).Add "Count", 0
        pronounsDict(pronoun).Add "Pages", CreateObject("Scripting.Dictionary")
    Next pronoun
    
    ' Step 6: Count pronouns with page numbers
    Call CountPronounsWithPages(selectedDoc, pronounsDict)
    
    ' Step 7: Count like terms with page numbers
    Set termCounts = CountLikeTermsWithPages(selectedDoc, likeTermGroups)
    
    ' Step 8: Check for number style issues
    Call CheckNumberStyleIssues(selectedDoc, numberStyleIssues)
    
    ' Step 9: Prepare the output
    output = PrepareOutputWithPages(pronounsDict, termCounts, numberStyleIssues)
    
    ' Step 10: Display the results in a message box
    If output <> "" Then
        MsgBox output, vbInformation, "Analysis Results"
    Else
        MsgBox "No relevant terms found based on the analysis criteria.", vbInformation
    End If
    
    ' Step 11: Generate the summary report as a text file
    reportFilePath = GenerateReport(targetDocPath, output)
    If reportFilePath <> "" Then
        MsgBox "Summary report generated at:" & vbCrLf & reportFilePath, vbInformation, "Report Generated"
    Else
        MsgBox "Failed to generate the summary report.", vbCritical
    End If
    
    ' Optional: Close the document without saving
    ' selectedDoc.Close SaveChanges:=False
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
End Sub

' New function to check for number style issues
Sub CheckNumberStyleIssues(doc As Document, numberStyleIssues As Object)
    Dim findRange As Range
    Dim pageNumber As Long
    Dim numberText As String
    Dim lastPosition As Long
    Dim context As String
    
    ' Check for small numbers (1-9)
    Set findRange = doc.Content
    lastPosition = 0
    
    With findRange.Find
        .ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .MatchCase = False
        
        ' Pattern to find standalone single digits with word boundaries
        .text = "<[1-9]>"
        
        Do While .Execute
            ' Prevent infinite loop by checking if we're moving forward
            If findRange.Start <= lastPosition Then
                Exit Do
            End If
            lastPosition = findRange.Start
            
            ' Extract the actual number
            numberText = Trim(findRange.text)
            
            If IsNumeric(numberText) Then
                pageNumber = findRange.Information(wdActiveEndPageNumber)
                
                If Not numberStyleIssues("small_numbers").Exists(pageNumber) Then
                    Set numberStyleIssues("small_numbers")(pageNumber) = CreateObject("Scripting.Dictionary")
                End If
                
                context = GetSurroundingContext(findRange)
                numberStyleIssues("small_numbers")(pageNumber).Add _
                    numberStyleIssues("small_numbers")(pageNumber).count + 1, _
                    "Found '" & numberText & "' - should be spelled out. Context: " & context
            End If
            
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Check for numbers at sentence starts
    Set findRange = doc.Content
    lastPosition = 0
    
    With findRange.Find
        .ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .MatchCase = False
        
        ' Find numbers at start of content or after period + space
        .text = ". [0-9]@"
        
        Do While .Execute
            If findRange.Start <= lastPosition Then
                Exit Do
            End If
            lastPosition = findRange.Start
            
            ' Move past the period and space to get just the number
            Dim numStart As Long
            numStart = findRange.Start + 2
            
            ' Get the text starting from the number
            Dim testRange As Range
            Set testRange = doc.Range(numStart, findRange.End)
            numberText = ""
            
            ' Extract just the numeric portion
            Dim i As Long
            For i = 1 To Len(testRange.text)
                If IsNumeric(Mid(testRange.text, i, 1)) Then
                    numberText = numberText & Mid(testRange.text, i, 1)
                ElseIf numberText <> "" Then
                    Exit For
                End If
            Next i
            
            If numberText <> "" Then
                pageNumber = findRange.Information(wdActiveEndPageNumber)
                
                If Not numberStyleIssues("sentence_start").Exists(pageNumber) Then
                    Set numberStyleIssues("sentence_start")(pageNumber) = CreateObject("Scripting.Dictionary")
                End If
                
                context = GetSurroundingContext(findRange)
                numberStyleIssues("sentence_start")(pageNumber).Add _
                    numberStyleIssues("sentence_start")(pageNumber).count + 1, _
                    "Sentence starts with '" & numberText & "' - should be spelled out. Context: " & context
            End If
            
            findRange.Collapse wdCollapseEnd
        Loop
    End With
    
    ' Check for numbers at paragraph starts
    Set findRange = doc.Content
    lastPosition = 0
    
    With findRange.Find
        .ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .MatchCase = False
        
        ' Find numbers at paragraph starts
        .text = "^13[0-9]@"
        
        Do While .Execute
            If findRange.Start <= lastPosition Then
                Exit Do
            End If
            lastPosition = findRange.Start
            
            ' Skip the paragraph mark
            Dim parStart As Long
            parStart = findRange.Start + 1
            
            ' Get the text starting from after the paragraph mark
            Set testRange = doc.Range(parStart, findRange.End)
            numberText = ""
            
            ' Extract just the numeric portion
            For i = 1 To Len(testRange.text)
                If IsNumeric(Mid(testRange.text, i, 1)) Then
                    numberText = numberText & Mid(testRange.text, i, 1)
                ElseIf numberText <> "" Then
                    Exit For
                End If
            Next i
            
            If numberText <> "" Then
                pageNumber = findRange.Information(wdActiveEndPageNumber)
                
                If Not numberStyleIssues("sentence_start").Exists(pageNumber) Then
                    Set numberStyleIssues("sentence_start")(pageNumber) = CreateObject("Scripting.Dictionary")
                End If
                
                context = GetSurroundingContext(findRange)
                numberStyleIssues("sentence_start")(pageNumber).Add _
                    numberStyleIssues("sentence_start")(pageNumber).count + 1, _
                    "Sentence starts with '" & numberText & "' - should be spelled out. Context: " & context
            End If
            
            findRange.Collapse wdCollapseEnd
        Loop
    End With
End Sub

' New helper function to check for special number patterns
Function IsSpecialPattern(text As String) As Boolean
    IsSpecialPattern = False
End Function


' Helper function to extract number from text
Function ExtractNumber(text As String) As String
    Dim i As Long
    Dim result As String
    
    result = ""
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            result = result & Mid(text, i, 1)
        ElseIf result <> "" Then
            Exit For
        End If
    Next i
    
    ExtractNumber = result
End Function

' Helper function to extract the leading number from text
Function ExtractLeadingNumber(text As String) As String
    Dim i As Long
    Dim result As String
    
    result = ""
    For i = 1 To Len(text)
        If IsNumeric(Mid(text, i, 1)) Then
            result = result & Mid(text, i, 1)
        ElseIf result <> "" Then
            Exit For
        End If
    Next i
    
    ExtractLeadingNumber = result
End Function

' Helper function to check if a number is standalone (not part of a larger number)
Function IsStandaloneNumber(fullText As String, numberStr As String) As Boolean
    Dim pattern As String
    pattern = "^" & numberStr & "$"
    
    ' Remove any surrounding whitespace or punctuation
    fullText = Trim(fullText)
    fullText = Replace(Replace(fullText, ".", ""), ",", "")
    
    IsStandaloneNumber = (fullText = numberStr)
End Function

' Helper function to check if the range is in a special context
Function IsSpecialContext(rng As Range) As Boolean
    ' Check if within a heading
    If rng.Style Like "Heading*" Then
        IsSpecialContext = True
        Exit Function
    End If
    
    ' Check if within a list
    If rng.ListFormat.ListType <> wdListNoNumbering Then
        IsSpecialContext = True
        Exit Function
    End If
    
    ' Check if within brackets
    Dim paragraphText As String
    paragraphText = rng.Paragraphs(1).Range.text
    If InStr(paragraphText, "[") > 0 And InStr(paragraphText, "]") > 0 Then
        IsSpecialContext = True
        Exit Function
    End If
    
    ' Check if in table
    If rng.Information(wdWithInTable) Then
        IsSpecialContext = True
        Exit Function
    End If
    
    IsSpecialContext = False
End Function

' Helper function to get surrounding context
Function GetSurroundingContext(rng As Range) As String
    Dim contextRange As Range
    Set contextRange = rng.Duplicate
    
    On Error Resume Next
    ' Get surrounding words, but handle edge cases
    contextRange.MoveStart wdWord, -3
    If Err.Number <> 0 Then
        contextRange.MoveStart wdWord, 0
    End If
    Err.Clear
    
    contextRange.MoveEnd wdWord, 3
    If Err.Number <> 0 Then
        contextRange.MoveEnd wdWord, 0
    End If
    On Error GoTo 0
    
    ' Clean up the context text
    Dim contextText As String
    contextText = Trim(Replace(Replace(contextRange.text, vbCr, " "), vbLf, " "))
    ' Remove multiple spaces
    Do While InStr(contextText, "  ") > 0
        contextText = Replace(contextText, "  ", " ")
    Loop
    
    GetSurroundingContext = "..." & contextText & "..."
End Function


' Function to display a file dialog and return the selected file path
Function GetSelectedFilePath(prompt As String, fileFilter As String) As String
    Dim fd As FileDialog
    Dim selectedPath As String
    Dim filterParts() As String
    Dim i As Integer
    
    On Error GoTo DialogError
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = prompt
        .Filters.Clear
        
        ' Split the fileFilter string by "|" to get description and pattern pairs
        filterParts = Split(fileFilter, "|")
        
        ' Ensure that filterParts has even number of elements (description and pattern)
        If UBound(filterParts) Mod 2 <> 1 Then
            MsgBox "Invalid file filter format.", vbCritical
            GetSelectedFilePath = ""
            Exit Function
        End If
        
        ' Iterate through the filterParts array in pairs
        For i = 0 To UBound(filterParts) Step 2
            .Filters.Add Trim(filterParts(i)), Trim(filterParts(i + 1)), (i / 2) + 1
        Next i
        
        .AllowMultiSelect = False
        
        If .Show Then
            selectedPath = .SelectedItems(1)
        Else
            selectedPath = ""
        End If
    End With
    
    GetSelectedFilePath = selectedPath
    Exit Function

DialogError:
    MsgBox "Error initializing file dialog: " & Err.Description, vbCritical
    GetSelectedFilePath = ""
End Function

' Function to read the configuration file and return a dictionary of term groups
Function ReadConfigFile(filePath As String) As Object
    Dim fso As Object
    Dim ts As Object
    Dim line As String
    Dim groupName As String
    Dim termsPart As String
    Dim termsArray() As String
    Dim term As Variant ' Changed from String to Variant
    Dim likeTermGroups As Object ' Dictionary
    Dim termCollection As Object ' Dictionary
    
    On Error GoTo ReadError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "Configuration file does not exist: " & filePath, vbCritical
        Set ReadConfigFile = Nothing
        Exit Function
    End If
    
    Set ts = fso.OpenTextFile(filePath, 1) ' ForReading
    Set likeTermGroups = CreateObject("Scripting.Dictionary")
    
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        
        ' Skip empty lines or lines without brackets
        If Len(line) > 0 And InStr(line, "[") > 0 And InStr(line, "]") > 0 Then
            groupName = Trim(Left(line, InStr(line, "[") - 1))
            termsPart = Mid(line, InStr(line, "[") + 1, InStr(line, "]") - InStr(line, "[") - 1)
            termsArray = Split(termsPart, ",")
            
            Set termCollection = CreateObject("Scripting.Dictionary")
            For Each term In termsArray
                term = Trim(LCase(term))
                If Len(term) > 0 Then
                    Set termCollection(term) = CreateObject("Scripting.Dictionary")
                    termCollection(term).Add "Count", 0
                    termCollection(term).Add "Pages", CreateObject("Scripting.Dictionary")
                End If
            Next term
            
            If Not likeTermGroups.Exists(groupName) Then
                Set likeTermGroups(groupName) = termCollection ' Use Set keyword
            End If
        End If
    Loop
    
    ts.Close
    Set ReadConfigFile = likeTermGroups
    Exit Function

ReadError:
    MsgBox "Error reading configuration file: " & Err.Description, vbCritical
    Set ReadConfigFile = Nothing
End Function

' Function to open a Word document and return the Document object
Function OpenDocument(docPath As String) As Document
    On Error GoTo OpenError
    Set OpenDocument = Documents.Open(FileName:=docPath, ReadOnly:=True)
    Exit Function

OpenError:
    MsgBox "Error opening document: " & Err.Description, vbCritical
    Set OpenDocument = Nothing
End Function

' Subroutine to count pronouns and record page numbers
Sub CountPronounsWithPages(doc As Document, pronounsDict As Object)
    Dim pronoun As Variant
    Dim rng As Range
    Dim findRange As Range
    Dim pageNumber As Long
    
    For Each pronoun In pronounsDict.Keys
        ' Initialize the range to the start of the document
        Set findRange = doc.Content
        With findRange.Find
            .text = pronoun
            .MatchWildcards = False  ' Changed from True
            .MatchWholeWord = True   ' Added
            .MatchCase = False
            .Forward = True
            .Wrap = wdFindStop
        End With
        
        Do While findRange.Find.Execute
            ' Increment count
            pronounsDict(pronoun)("Count") = pronounsDict(pronoun)("Count") + 1
            
            ' Get page number
            pageNumber = findRange.Information(wdActiveEndPageNumber)
            If Not pronounsDict(pronoun)("Pages").Exists(pageNumber) Then
                pronounsDict(pronoun)("Pages").Add pageNumber, 1
            End If
            
            ' Move the range to after the found text
            findRange.Collapse wdCollapseEnd
        Loop
    Next pronoun
End Sub

' Function to count like terms and record page numbers
Function CountLikeTermsWithPages(doc As Document, likeTermGroups As Object) As Object
    Dim groupName As Variant
    Dim termsDict As Object
    Dim term As Variant
    Dim findRange As Range
    Dim pageNumber As Long
    
    Dim groupCounts As Object
    Set groupCounts = CreateObject("Scripting.Dictionary")
    
    For Each groupName In likeTermGroups.Keys
        Set termsDict = likeTermGroups(groupName)
        
        For Each term In termsDict.Keys
            ' Initialize
            Set findRange = doc.Content
            With findRange.Find
                .text = term
                .MatchWildcards = False  ' Changed from True
                .MatchWholeWord = True   ' Added
                .MatchCase = False
                .Forward = True
                .Wrap = wdFindStop
            End With
            
            Do While findRange.Find.Execute
                ' Increment count
                termsDict(term)("Count") = termsDict(term)("Count") + 1
                
                ' Get page number
                pageNumber = findRange.Information(wdActiveEndPageNumber)
                If Not termsDict(term)("Pages").Exists(pageNumber) Then
                    termsDict(term)("Pages").Add pageNumber, 1
                End If
                
                ' Move the range to after the found text
                findRange.Collapse wdCollapseEnd
            Loop
        Next term
        
        ' Check for active terms after counting
        Dim activeTermsCount As Long
        activeTermsCount = 0
        
        For Each term In termsDict.Keys
            If termsDict(term)("Count") > 0 Then
                activeTermsCount = activeTermsCount + 1
            End If
        Next term
        
        If activeTermsCount > 0 Then  ' Changed from >= 2 to > 0
            Set groupCounts(groupName) = termsDict
        End If
    Next groupName
    
    Set CountLikeTermsWithPages = groupCounts
End Function

' Function to prepare the output string based on pronoun counts and like term counts, including page numbers
Function PrepareOutputWithPages(pronounsDict As Object, termCounts As Object, numberStyleIssues As Object) As String
    Dim output As String
    Dim pronoun As Variant
    Dim groupName As Variant
    Dim term As Variant
    Dim pages As Variant
    Dim pageList As String
    Dim key As Variant
    
    output = ""
    
    ' Prepare Pronouns Found
    Dim pronounsFound As String
    pronounsFound = "Pronouns Found:" & vbCrLf
    Dim hasPronouns As Boolean
    hasPronouns = False
    For Each pronoun In pronounsDict.Keys
        If pronounsDict(pronoun)("Count") > 0 Then
            pronounsFound = pronounsFound & "- " & pronoun & ": " & pronounsDict(pronoun)("Count") & " occurrence(s) on page(s): "
            ' Collect page numbers
            pageList = ""
            For Each key In pronounsDict(pronoun)("Pages").Keys
                pageList = pageList & key & ", "
            Next key
            ' Remove trailing comma and space
            If Len(pageList) > 2 Then
                pageList = Left(pageList, Len(pageList) - 2)
            End If
            pronounsFound = pronounsFound & pageList & vbCrLf
            hasPronouns = True
        End If
    Next pronoun
    If hasPronouns Then
        output = output & pronounsFound & vbCrLf
    End If
    
    ' Prepare Like Terms Found
    Dim likeTermsFound As String
    likeTermsFound = "Like Terms Found:" & vbCrLf
    Dim hasLikeTerms As Boolean
    hasLikeTerms = False
    For Each groupName In termCounts.Keys
        likeTermsFound = likeTermsFound & groupName & ":" & vbCrLf
        For Each term In termCounts(groupName).Keys
            If termCounts(groupName)(term)("Count") > 0 Then
                likeTermsFound = likeTermsFound & "  - " & term & ": " & termCounts(groupName)(term)("Count") & " occurrence(s) on page(s): "
                ' Collect page numbers
                pageList = ""
                For Each key In termCounts(groupName)(term)("Pages").Keys
                    pageList = pageList & key & ", "
                Next key
                ' Remove trailing comma and space
                If Len(pageList) > 2 Then
                    pageList = Left(pageList, Len(pageList) - 2)
                End If
                likeTermsFound = likeTermsFound & pageList & vbCrLf
            End If
        Next term
        likeTermsFound = likeTermsFound & vbCrLf
        hasLikeTerms = True
    Next groupName
    If hasLikeTerms Then
        output = output & likeTermsFound
    End If
    
    ' Add number style issues section
    output = output & vbCrLf & "Number Style Issues:" & vbCrLf
    
    ' Small numbers (1-9) issues
    If numberStyleIssues("small_numbers").count > 0 Then
        output = output & vbCrLf & "Numbers 1-9 that should be spelled out:" & vbCrLf
        Dim pageNum As Variant
        Dim issue As Variant
        For Each pageNum In numberStyleIssues("small_numbers").Keys
            output = output & "Page " & pageNum & ":" & vbCrLf
            For Each issue In numberStyleIssues("small_numbers")(pageNum)
                output = output & "  - " & numberStyleIssues("small_numbers")(pageNum)(issue) & vbCrLf
            Next issue
        Next pageNum
    End If
    
    ' Numbers at sentence start issues
    If numberStyleIssues("sentence_start").count > 0 Then
        output = output & vbCrLf & "Numbers at the start of sentences:" & vbCrLf
        For Each pageNum In numberStyleIssues("sentence_start").Keys
            output = output & "Page " & pageNum & ":" & vbCrLf
            For Each issue In numberStyleIssues("sentence_start")(pageNum)
                output = output & "  - " & numberStyleIssues("sentence_start")(pageNum)(issue) & vbCrLf
            Next issue
        Next pageNum
    End If
    
    PrepareOutputWithPages = output
End Function

' Function to generate a summary report as a text file
Function GenerateReport(targetDocPath As String, reportContent As String) As String
    Dim fso As Object
    Dim ts As Object
    Dim reportPath As String
    Dim docFolder As String
    Dim docName As String
    Dim reportName As String
    
    On Error GoTo ReportError
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    docFolder = fso.GetParentFolderName(targetDocPath)
    docName = fso.GetBaseName(targetDocPath)
    reportName = docName & "_SummaryReport.txt"
    reportPath = fso.BuildPath(docFolder, reportName)
    
    Set ts = fso.CreateTextFile(reportPath, True) ' Overwrite if exists
    ts.WriteLine reportContent
    ts.Close
    
    GenerateReport = reportPath
    Exit Function

ReportError:
    MsgBox "Error generating report: " & Err.Description, vbCritical
    GenerateReport = ""
End Function

