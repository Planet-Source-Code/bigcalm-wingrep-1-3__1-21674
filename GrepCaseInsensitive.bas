Attribute VB_Name = "GrepCaseInsensitive"
Option Explicit
Option Base 0
Option Compare Text
' The two grep modules (GrepCase & GrepCaseInsensitive) are almost identical, the only
' difference being the Option Compare statement and the Function names within them.
' Everything else is the same.

Private Const MAXREGULAREXPRESSIONS = 1000

Public Function GrepCaseInSensitive(RegularExpression As String, FileName As String, _
        ByRef OutputFile As Variant, Optional InvertSelection As Boolean = False, _
        Optional ExactLineMatching As Boolean = False, _
        Optional DisplayFileName As GrepShowFileConstants = 0, _
        Optional LineNumber As Boolean = False, _
        Optional NumberOfLinesToOutputBeforeMatch As Long = 0, _
        Optional NumberOfLinesToOuputAfterMatch As Long = 0, _
        Optional JustCountThem As Boolean = False, _
        Optional Seperator As GrepSeperatorTypes = NoSeperator) As Long
Dim FileNumber As Long      ' Is the opened file number for input
Dim InputLine As String         ' Is the variable read by Line Input #
Dim OutputLine As String      ' Contains any output required to output file
Dim OutputThisLine As Boolean ' Internal variable to mark whether a line needs to be output
Dim RegExp() As String    ' Regular expressions
Dim Regularexpressions As Long ' Total in above array
Dim CloseOutputFile As Boolean ' Whether I want to close the output file when I've finished
Dim CurrentLine As Long ' Current Line in Input file
Dim BackLines() As String ' Buffer containing the previous lines read
Dim BackLinePointer As Long ' Pointer to above
Dim ForwardLines As Collection ' Contains buffered lines for read-ahead
Dim i As Long, j As Long    ' General all purpose looping control variables
Dim LinesSinceMatch As Long ' Number of lines since a match was found
Dim MatchingLines As Long ' Total number of matching lines found so far
Dim DisplayedFile As Boolean ' If DisplayFileName = Every File, whether we've displayed it yet
Dim InputLine2 As String ' Temporary string for read-ahead purposes.

    '-------------------------------------
    ' Initialize
    '-------------------------------------
    If NumberOfLinesToOutputBeforeMatch > 0 Then
        ReDim BackLines(NumberOfLinesToOutputBeforeMatch + 1)
        BackLinePointer = 1
    End If
    LinesSinceMatch = -1
    MatchingLines = 0
    DisplayedFile = False
    Set ForwardLines = New Collection
    
    '-------------------------------------
    ' Output File Handling
    '-------------------------------------

    ' Let's sort the Output file out first...
    ' I've held it as a variant so that if it's a STRING, we need to open it
    ' if it's a LONG, it's an already open file.
    If TypeName(OutputFile) = "String" Then
        On Error GoTo ErrHandler
        Open OutputFile For Output Access Write As #FileNumber
        OutputFile = FileNumber
        On Error GoTo 0
        CloseOutputFile = True
    Else
        CloseOutputFile = False
    End If
    
    '-------------------------------------
    ' Regular Expression Split
    '-------------------------------------
    
    Regularexpressions = 0
    ReDim RegExp((CountOccurrences(RegularExpression, ";") + 1))
    SplitStringIntoParts RegularExpression, ";"
    Do While Regularexpressions <= MAXREGULAREXPRESSIONS
        RegExp(Regularexpressions) = GetNextPartOfSplitString
        If Len(RegExp(Regularexpressions)) = 0 Then
            If Regularexpressions > 0 Then
                Regularexpressions = Regularexpressions - 1
                Exit Do
            End If
        End If
        If ExactLineMatching = False Then
            RegExp(Regularexpressions) = "*" & RegExp(Regularexpressions) & "*"
        End If
        Regularexpressions = Regularexpressions + 1
    Loop
    
    '-------------------------------------
    ' Input File Handling
    '-------------------------------------
    On Error GoTo ErrHandler2
    FileNumber = FreeFile
    Open FileName For Input Access Read As #FileNumber
    On Error GoTo 0
    CurrentLine = 0
    
    '-------------------------------------
    ' Main Loop
    '-------------------------------------
    
    Do While Not (EOF(FileNumber) And ForwardLines.Count = 0)
        '-------------------------------------
        ' Get Next Line
        '-------------------------------------
        If ForwardLines.Count > 0 Then
            InputLine = ForwardLines(1)
            ForwardLines.Remove (1)
        Else
            Line Input #FileNumber, InputLine
        End If
        
        '-------------------------------------
        ' Re-organise Back-Lines Buffer if necessary
        '-------------------------------------
        If NumberOfLinesToOutputBeforeMatch > 0 Then
            BackLines(BackLinePointer) = InputLine
            BackLinePointer = BackLinePointer + 1
            If BackLinePointer > NumberOfLinesToOutputBeforeMatch + 1 Then
                BackLinePointer = 1
            End If
        End If
        '-------------------------------------
        ' Initialize loop
        '-------------------------------------
        CurrentLine = CurrentLine + 1
        OutputThisLine = False
        '-------------------------------------
        ' Comparison!
        '-------------------------------------
        For i = 0 To Regularexpressions
            If InputLine Like RegExp(i) Then
                OutputThisLine = True
                Exit For
            End If
        Next
        '-------------------------------------
        ' Handling of regular expression parameters
        '-------------------------------------
        If InvertSelection = True Then
            OutputThisLine = Not OutputThisLine
        End If
        If OutputThisLine = True Then
            '-------------------------------------
            ' Ouput!
            '-------------------------------------
            MatchingLines = MatchingLines + 1
            If JustCountThem = False Then
                '-------------------------------------
                ' Output all lines in back buffer
                '-------------------------------------
                If NumberOfLinesToOutputBeforeMatch > 0 Then
                    OutputLine = ""
                    i = BackLinePointer
                    If CurrentLine < i Then
                        i = 1
                    End If
                    j = CurrentLine - NumberOfLinesToOutputBeforeMatch
                    If j <= 0 Then
                        j = 1
                    End If
                    Do While True
                        OutputLine = ""
                        If DisplayFileName = OncePerline Then
                            OutputLine = OutputLine & FileName & ": "
                        End If
                        If DisplayFileName = OncePerFile And DisplayedFile = False Then
                            OutputLine = OutputLine & FileName & ": "
                            Print #OutputFile, OutputLine
                            DisplayedFile = True
                            OutputLine = ""
                        End If
                        If LineNumber = True Then
                            OutputLine = OutputLine & Format(j, "0") & ": "
                        End If
                        j = j + 1
                        OutputLine = OutputLine & BackLines(i)
                        Print #OutputFile, OutputLine
                        i = i + 1
                        If i > NumberOfLinesToOutputBeforeMatch + 1 Then
                            i = 1
                        End If
                        If BackLinePointer = 1 Then
                            If i = NumberOfLinesToOutputBeforeMatch + 1 Then
                                Exit Do
                            End If
                        Else
                            If i = BackLinePointer - 1 Then
                                Exit Do
                            End If
                        End If
                    Loop
                End If
                '-------------------------------------
                ' Output THIS line
                '-------------------------------------
                OutputLine = ""
                If DisplayFileName = OncePerline Then
                    OutputLine = OutputLine & FileName & ": "
                End If
                If DisplayFileName = OncePerFile And DisplayedFile = False Then
                    OutputLine = OutputLine & FileName & ": "
                    Print #OutputFile, OutputLine
                    DisplayedFile = True
                    OutputLine = ""
                End If
                If LineNumber = True Then
                    OutputLine = OutputLine & Format(CurrentLine, "0") & ": "
                End If
                OutputLine = OutputLine & InputLine
                Print #OutputFile, OutputLine
                
                '-------------------------------------
                ' Output Lines AFTER match
                '-------------------------------------
                If NumberOfLinesToOuputAfterMatch > 0 Then
                    ' If there's already items in the collection we need to output them first
                    If ForwardLines.Count > 0 Then
                        For i = 1 To ForwardLines.Count
                            InputLine2 = ForwardLines(i)
                            OutputLine = ""
                            If DisplayFileName = OncePerline Then
                                OutputLine = OutputLine & FileName & ": "
                            End If
                            If DisplayFileName = OncePerFile And DisplayedFile = False Then
                                OutputLine = OutputLine & FileName & ": "
                                Print #OutputFile, OutputLine
                                DisplayedFile = True
                                OutputLine = ""
                            End If
                            If LineNumber = True Then
                                OutputLine = OutputLine & Format(CurrentLine + i, "0") & ": "
                            End If
                            OutputLine = OutputLine & InputLine2
                            Print #OutputFile, OutputLine
                        Next
                    End If
                    ' Read the lines into the collection
                    i = NumberOfLinesToOuputAfterMatch - ForwardLines.Count
                    Do While (i > 0 And (Not EOF(FileNumber)))
                        Line Input #FileNumber, InputLine2
                        ForwardLines.Add InputLine2
                        OutputLine = ""
                        If DisplayFileName = OncePerline Then
                            OutputLine = OutputLine & FileName & ": "
                        End If
                        If DisplayFileName = OncePerFile And DisplayedFile = False Then
                            OutputLine = OutputLine & FileName & ": "
                            Print #OutputFile, OutputLine
                            DisplayedFile = True
                            OutputLine = ""
                        End If
                        If LineNumber = True Then
                            OutputLine = OutputLine & Format(CurrentLine + ForwardLines.Count, "0") & ": "
                        End If
                        OutputLine = OutputLine & InputLine2
                        Print #OutputFile, OutputLine
                        i = i - 1
                    Loop
                End If
                '-------------------------------------
                ' Print Seperator
                '-------------------------------------
                If Seperator = SeperateMatches Then
                    OutputLine = SeperatorString
                    Print #OutputFile, OutputLine
                End If
            End If
            LinesSinceMatch = 0
        End If
    Loop ' End of main loop
    Close #FileNumber
    
    '-------------------------------------
    ' Anything that applies to the whole file is done here
    '-------------------------------------
    If Seperator = SeperateFiles And MatchingLines > 0 Then
        OutputLine = SeperatorString
        Print #OutputFile, OutputLine
    End If
    If JustCountThem = True Then
        If DisplayFileName = OncePerline Or (DisplayFileName = OncePerFile And MatchingLines > 0) Then
            OutputLine = FileName & ": " & Format(MatchingLines, "0")
            Print #OutputFile, OutputLine
        Else
            If CloseOutputFile = True Then
                OutputLine = Format(MatchingLines, "0")
                Print #OutputFile, OutputLine
            End If
        End If
    End If
    If CloseOutputFile = True Then
        Close #OutputFile
    End If
    '-------------------------------------
    ' Set Return value to be the number of matched lines found in the file.  Otherwise
    ' set to -1 if an error occurred
    '-------------------------------------
    GrepCaseInSensitive = MatchingLines
    Exit Function
ErrHandler:
    #If PartOfWinGrep = 1 Then
    MsgboxForWinGrep "Error " & Err.Number & " " & Err.Description & ": " & OutputFile
    #Else
    MsgBox "Error " & Err.Number & " " & Err.Description & ": " & OutputFile
    #End If
    GrepCaseInSensitive = -1
    Exit Function
ErrHandler2:
    #If PartOfWinGrep = 1 Then
      MsgboxForWinGrep "Error " & Err.Number & " " & Err.Description & ": " & FileName
    #Else
      MsgBox "Error " & Err.Number & " " & Err.Description & ": " & FileName
    #End If
    GrepCaseInSensitive = -1
    Exit Function
End Function

