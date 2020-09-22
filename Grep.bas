Attribute VB_Name = "GrepAndOtherUsefuls"
Option Explicit
Option Base 0
Option Compare Text

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Const cMaxPath = &H104

Public Enum GrepShowFileConstants
    Never = 0
    OncePerline = 1
    OncePerFile = 2
End Enum

Public Enum GrepSeperatorTypes
    NoSeperator = 0
    SeperateMatches = 1
    SeperateFiles = 2
End Enum

Public Const SeperatorString = "____________________________________________________"
Public Const WinGrepVersion = "1.3"

Public Enum EShellShowConstants
    essSW_HIDE = 0
    essSW_MAXIMIZE = 3
    essSW_MINIMIZE = 6
    essSW_SHOWMAXIMIZED = 3
    essSW_SHOWMINIMIZED = 2
    essSW_SHOWNORMAL = 1
    essSW_SHOWNOACTIVATE = 4
    essSW_SHOWNA = 8
    essSW_SHOWMINNOACTIVE = 7
    essSW_SHOWDEFAULT = 10
    essSW_RESTORE = 9
    essSW_SHOW = 5
End Enum
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5        ' access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                ' file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                ' path not found
Private Const SE_ERR_OOM = 8                ' out of memory
Private Const SE_ERR_SHARE = 26

Global Const CB_ERR = -1
Global Const CB_FINDSTRING = &H14C
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' For split string purposes
Private mSplitLine As String ' These three vars are used to
Private mDelimiter As String ' split a delimiter seperated line up
Private mCurrentPos As Long

Public Function ShellEx( _
        ByVal sFIle As String, _
        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
        Optional ByVal sParameters As String = "", _
        Optional ByVal sDefaultDir As String = "", _
        Optional sOperation As String = "open", _
        Optional Owner As Long = 0 _
    ) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As String
    If (InStr(UCase$(sFIle), ".EXE") <> 0) Then
        eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFIle, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFIle, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
            sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy. Please try again in a moment."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed. Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "The file could not be opened due to time out. Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
            lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
            lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
            sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
                
        Err.Raise lErr, , App.EXEName & ".GShell: " & sErr
        ShellEx = False
    End If

End Function

Public Function GetWinTempDir() As String
Dim s As String, c As Long
    s = String$(cMaxPath, 0)
    c = GetTempPath(cMaxPath, s)
    GetWinTempDir = Left(s, c)
End Function

Public Function GetWinSystemDir() As String
Dim s As String, c As Long
    s = String$(cMaxPath, 0)
    c = GetSystemDirectory(s, cMaxPath)
    GetWinSystemDir = Left(s, c)
End Function


' I am ALWAYS using these two functions...
' Important: Cannot be used recursively!
' If you have VB6 you can use Split() instead.
Public Sub SplitStringIntoParts(pLine As String, pDelimiter)
    mSplitLine = pLine
    mDelimiter = pDelimiter
    mCurrentPos = 1
End Sub

Public Function GetNextPartOfSplitString() As String
Dim lCurrentPos As Long
    If mCurrentPos > Len(mSplitLine) Then
        GetNextPartOfSplitString = ""
    Else
        lCurrentPos = InStr(mCurrentPos, mSplitLine, mDelimiter)
        If lCurrentPos = 0 Then
            ' Get rest of line
            GetNextPartOfSplitString = Mid(mSplitLine, mCurrentPos, (Len(mSplitLine) - mCurrentPos) + 1)
            mCurrentPos = Len(mSplitLine) + 1
        Else
            GetNextPartOfSplitString = Mid(mSplitLine, mCurrentPos, (lCurrentPos - mCurrentPos))
            mCurrentPos = lCurrentPos + Len(mDelimiter)
        End If
    End If
End Function

' Search and Replace
' If you have VB6 you can use Replace()
Public Function QSAR(ByVal pString As String, ByVal pSearch As String, Optional ByVal pReplace As String = "", Optional pCompare As Long = vbBinaryCompare) As String
Dim lLen1 As Long
Dim lLen2 As Long
Dim lStartFind As Long
Dim lFoundLoc As Long
Dim ltmpString As String

    lLen1 = Len(pString)
    lLen2 = Len(pSearch)
    lStartFind = 1
    ltmpString = ""
    Do
        lFoundLoc = InStr(lStartFind, pString, pSearch, pCompare)
        If lFoundLoc = 0 Then
            Exit Do
        End If
        ltmpString = ltmpString & Mid(pString, lStartFind, lFoundLoc - lStartFind) & pReplace
        lStartFind = lFoundLoc + lLen2
    Loop
    ltmpString = ltmpString & Mid(pString, lStartFind, lLen1 - lStartFind + 1)
    QSAR = ltmpString
End Function

' Function to count occurences of one string within another
Public Function CountOccurrences(Text As String, Find As String) As Long
Dim CurrPos As Long
  CountOccurrences = -1
  CurrPos = 1
  Do
    CurrPos = InStr(CurrPos, Text, Find, vbTextCompare) + Len(Find)
    CountOccurrences = CountOccurrences + 1
  Loop Until CurrPos = 1
End Function

' called by frmGreta
Public Function Grep(RegularExpression As String, FileName As String, _
        ByRef OutputFile As Variant, Optional InvertSelection As Boolean = False, _
        Optional ExactLineMatching As Boolean = False, _
        Optional CaseSensitive As Boolean = False, _
        Optional DisplayFileName As GrepShowFileConstants = 0, _
        Optional LineNumber As Boolean = False, _
        Optional NumberOfLinesToOutputBeforeMatch As Long = 0, _
        Optional NumberOfLinesToOuputAfterMatch As Long = 0, _
        Optional JustCountThem As Boolean, _
        Optional Seperator As GrepSeperatorTypes = NoSeperator) As Long
    
    If CaseSensitive = True Then
        Grep = GrepCase.GrepCaseSensitive(RegularExpression, _
            FileName, OutputFile, InvertSelection, _
            ExactLineMatching, DisplayFileName, LineNumber, _
            NumberOfLinesToOutputBeforeMatch, NumberOfLinesToOuputAfterMatch, _
            JustCountThem, Seperator)
    Else
        Grep = GrepCaseInSensitive.GrepCaseInSensitive(RegularExpression, _
            FileName, OutputFile, InvertSelection, _
            ExactLineMatching, DisplayFileName, LineNumber, _
            NumberOfLinesToOutputBeforeMatch, NumberOfLinesToOuputAfterMatch, _
            JustCountThem, Seperator)
    End If
End Function

