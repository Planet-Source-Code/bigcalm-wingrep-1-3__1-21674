Attribute VB_Name = "FileFinder"
Option Explicit
Option Compare Text
Option Base 0
' This module contains procedures to help with file finding.
' Ideally should be bundled with FrmFindFiles, though you may find that you can use
' the code in this module seperately (you'll have to remove the function FileFindDialog)

' Recursively searches a directory tree trying to match files using a regular expression
' such as *.txt (or combinations thereof, e.g. "*.cls;*.bas")
' Any files that match are placed into a collection of strings.

' You will not be able to use the For Each method on this collection, (without a bit of kludging anyway).
' instead use:
' for i = 1 to collection.count
'      ...
' next
'
' If anyone wants, I'll turn this into an Event-driven class or User-control.

Private Const MaxRecursionDepth As Long = 512

Private FileColl As Collection
Public ffRunning As Boolean
Public ffStopping As Boolean

Public Function FileFindDialog(ByRef pCollection As Collection, Optional pStartPath As String = "")
    frmFindFiles.FindFilesForm pCollection, pStartPath
End Function

Public Function FindFiles(Path As String, RegExp As String, _
        Optional RecurseSubFolders As Boolean = True, _
        Optional OlderThan As Date, _
        Optional ProgressLabel As Label = Nothing) As Collection
Dim X As String
Dim i As Long

    If OlderThan = Empty Then
      OlderThan = Now
    End If
    If ffRunning = True Then
        Set FindFiles = Nothing
        Exit Function
    End If
    ffRunning = True
    ffStopping = False
   Set FileColl = New Collection
   DirectoryWalk Path, RegExp, 1, RecurseSubFolders, OlderThan, ProgressLabel
   Set FindFiles = FileColl
   If ffStopping = True Then
    Set FindFiles = Nothing
   End If
   ffRunning = False
   ffStopping = False
End Function

Public Function StopFindFiles()
    If ffRunning = True Then
        ffStopping = True
    End If
End Function

Private Sub DirectoryWalk(ByVal Path As String, ByVal RegularExpression As String, ByVal CurrentDepth As Long, _
        Optional ByVal RecurseSubFolders As Boolean = True, _
        Optional ByVal OlderThan As Date, _
        Optional ProgressLabel As Label = Nothing)
Dim PrivCollection As Collection
Dim File As String
Dim RegExp(10) As String
Dim Regularexpressions As Long
Dim i As Long
Dim Recurse As String

    If CurrentDepth > MaxRecursionDepth Then
      Exit Sub
    End If
    If Not (ProgressLabel Is Nothing) Then
        ProgressLabel.Caption = "Processing: " & Path
    End If
    DoEvents
    ' Read in regular expressions
    Regularexpressions = 0
    SplitStringIntoParts RegularExpression, ";"
    Do While Regularexpressions <= 10
        RegExp(Regularexpressions) = GetNextPartOfSplitString
        If Len(RegExp(Regularexpressions)) = 0 Then
            If Regularexpressions > 0 Then
                Regularexpressions = Regularexpressions - 1
                Exit Do
            End If
        End If
        Regularexpressions = Regularexpressions + 1
    Loop

    Set PrivCollection = New Collection
    ' Read in entries into directory
    If Right(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    File = Dir(Path & "*.*", vbDirectory)
    Do While File <> ""
        If File <> "." And File <> ".." Then
            If (GetAttr(Path & File) And vbDirectory) = vbDirectory Then
                PrivCollection.Add Path & File
            Else
                For i = 0 To Regularexpressions
                    If File Like RegExp(i) Then
                      If DateDiff("s", OlderThan, FileDateTime(Path & File)) < 0 Then
                        FileColl.Add Path & File
                        If Not (ProgressLabel Is Nothing) Then
                            ProgressLabel.Caption = "Found: " & Path & File
                        End If
                        DoEvents
                      End If
                    End If
                Next
            End If
        End If
        File = Dir
    Loop
    DoEvents
    ' Recursion
    If RecurseSubFolders = True And ffStopping = False Then
        For i = 1 To PrivCollection.Count
            Recurse = PrivCollection(i)
            Call DirectoryWalk(Recurse, RegularExpression, CurrentDepth + 1, RecurseSubFolders, OlderThan, ProgressLabel)
            If ffStopping = True Then
                Exit For
            End If
        Next
    End If
End Sub
