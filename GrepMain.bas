Attribute VB_Name = "GrepMain"
Option Explicit
Option Base 0
Option Compare Text
' Command Line:
' grep [-y] [-v] [-n] [-s1|-s2] [-R|-e] [-l1|-l2] [-V=Version] [-c|-pre=<Pre Lines> -post=<Post Lines>]
'       [-z1|-z2] { <Regular Expression> | -reg=<Regular Expression> }
'       { -find=<Directory>;<File Types>; [-recurse] | <File List> } [ > <OutputFile> ]
' Anything enclosed in square brackets is optional, anything in curly is required.
' Calling grep with no arguments will display the main program.

' Complete flag explanations for 1.3
' Where "Command Line Only" appears, this argument is only available from the command
' line.
' I have tried to keep rough equivalencies with the UNIX version (i.e. -y Case, -c Count, -l Files, -n Line Numbers)
' but this is a bit tricky, 'cos I've added more functionality than the original.

' -e    - Edit a script file (Command Line only)
' -f="<File Name>" - Run a script file (Command line only)
' -R    - Force it to run the script straight away (Command Line only)
' -s1  - Silent running option.  1 = Completely silent.  Error messages are ignored (Command Line only)
' -s2  - Cut-down running option.  2 = Just display a small form for progress (Command Line only)
'
' > <OutputFile>    - Specifies an output file.
' -c    - Display a count of matching lines.  Varied output depending on the -l flag
' -find="<Directory;FileType1;FileType2...> -  Specifies that grep should find all files in Directory,
'           which match FileType1.  Seperate by semi-colons again.
' -l1   - Display filenames once per line
' -l2   - Display filenames once per file (if a match is found).
' -n    - Display the line number that the match occurred on.
' -pre=<Number>  - Display <Number> of lines before match is found.  If the -c flag is present, this has no effect.
' -post=<Number>    - Display <Number> of lines after match is found.  If the -c flag is present, this has no effect.
' -recurse  - If the -find option is specifed, this makes the macro recurse through subfolders as well
' -reg="<Regular Expressions">  - Regular expression to use.  Most of the time you don't need the -reg=, and can
'       specify it just as it is, e.g.   grep -y "Nokia" C:\tmp\x.txt.  More than one can be specified by putting semi-colons between them.
' -v    - Inverts selection - extracts all lines that don't match the regular expression.
' -V=<Version>  - specifies the version of WinGrep that the script was written for.
' -y    - Ignore case.  Default is case-sensitive matching(!)
' -z1   - Specifies a line of underscores needs to be printed after every match is found
' -z2   - Specifies a line of underscores needs to be printed after every file.


Public Silence As Long

Sub Main()
Dim ArgVal As String
Dim Argument As String
Dim ScriptFile As String
Dim tmpStr As String
Dim Run As Long
Dim Edit As Long

    Silence = 0
    ScriptFile = ""
    Run = 0
    Edit = 0
    ArgVal = Command
    
    If Len(ArgVal) > 0 Then
        SplitStringIntoParts ArgVal, " "
        Argument = GetNextPartOfSplitString
        Do While Len(Argument) > 0
            If Argument = "-s1" Then
                Silence = 1
            End If
            If Argument = "-s2" Then
                Silence = 2
            End If
            If Argument = "-R" Then
                Run = 1
            End If
            If Argument = "-e" Then
                Edit = 1
            End If
            If Mid(Argument, 1, 3) = "-f=" Then
                If Mid(Argument, 4, 1) = """" And Mid(Argument, Len(Argument), 1) <> """" Then
                    tmpStr = GetNextPartOfSplitString
                    Do While Len(tmpStr) > 0
                        If Mid(tmpStr, Len(tmpStr), 1) = """" Then
                            Exit Do
                        End If
                        Argument = Argument & tmpStr & " "
                        tmpStr = GetNextPartOfSplitString
                    Loop
                    ScriptFile = Argument
                Else
                    ScriptFile = Mid(Argument, 5, Len(Argument) - 5)
                End If
            End If
            Argument = GetNextPartOfSplitString
        Loop
        If Len(ScriptFile) > 0 Then
            If Edit = 0 Then
                frmGreta.Hide
                frmGreta.LoadGrepFile ScriptFile
                If Silence = 0 Then
                    frmGreta.Show
                Else
                    If Silence = 1 Then
                        Screen.MousePointer = vbArrowHourglass
                        frmGreta.GrepGo
                        Do While frmGreta.Running = True
                            DoEvents
                        Loop
                        Unload frmGreta
                        Screen.MousePointer = vbDefault
                    Else
                        frmGreta.ShowJustProgressStuff
                        frmGreta.Show
                        frmGreta.GrepGo
                        Do While frmGreta.Running = True
                            DoEvents
                        Loop
                        Unload frmGreta
                    End If
                End If
            Else
                frmGreta.Hide
                frmGreta.LoadGrepFile ScriptFile
                frmGreta.Show
            End If
        Else
            If Run = 0 Then
                frmGreta.txtCommand.Text = ArgVal
                frmGreta.Show
                frmGreta.lblProcess.Caption = "Initialising Script..."
                frmGreta.ConvertGrepStringToOptionBoxes ArgVal
            Else
                frmGreta.Hide
                frmGreta.txtCommand.Text = ArgVal
                frmGreta.ConvertGrepStringToOptionBoxes ArgVal
                frmGreta.GrepGo
                Do While frmGreta.Running = True
                    DoEvents
                Loop
                Unload frmGreta
            End If
        End If
    Else
        frmGreta.Show
    End If
End Sub

Public Function MsgboxForWinGrep(Prompt As String, Optional Buttons As VbMsgBoxStyle, Optional Title As String, Optional HelpFile As Variant, Optional Context As Variant) As VbMsgBoxResult
    If Silence = 1 Then
        If Buttons And vbOK = vbOK Then
            MsgboxForWinGrep = vbOK
        Else
            If Buttons And vbYes = vbYes Then
                MsgboxForWinGrep = vbYes
            End If
        End If
    Else
        MsgboxForWinGrep = MsgBox(Prompt, Buttons, Title, HelpFile, Context)
    End If
End Function
