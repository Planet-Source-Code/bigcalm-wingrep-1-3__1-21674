VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmGreta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinGrep version 1.3"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmGreta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCommandLine 
      Caption         =   "Command"
      Height          =   1095
      Left            =   0
      TabIndex        =   49
      Top             =   6000
      Width           =   8295
      Begin VB.CommandButton cmdSaveToHS3 
         Height          =   375
         Left            =   6720
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmGreta.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Assign Current Information To Hot Search Button 3"
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton cmdSaveToHS2 
         Height          =   375
         Left            =   5280
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmGreta.frx":0628
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Assign Current Information To Hot Search Button 2"
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton cmdSaveToHS1 
         Height          =   375
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmGreta.frx":080E
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Assign Current Information To Hot Search Button 1"
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   255
      End
      Begin VB.CommandButton cmdHotButton3 
         Caption         =   "Hot Search (&3)"
         Height          =   375
         Left            =   6960
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdHotButton2 
         Caption         =   "Hot Search (&2)"
         Height          =   375
         Left            =   5520
         TabIndex        =   34
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdHotButton1 
         Caption         =   "Hot Search (&1)"
         Height          =   375
         Left            =   4080
         TabIndex        =   32
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "Loa&d..."
         Height          =   375
         Left            =   2520
         TabIndex        =   30
         ToolTipText     =   "Load a previously saved Grep script"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveGrep 
         Caption         =   "S&ave..."
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         ToolTipText     =   "Save the current grep script for later use"
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdPropogateUp 
         Caption         =   "&Propogate Up"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Update all input boxes above from Command Line"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCommand 
         Height          =   285
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label lblCommand 
         Caption         =   "C&ommand Line:"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
   End
   Begin ComctlLib.ProgressBar pb1 
      Height          =   135
      Left            =   0
      TabIndex        =   43
      Top             =   7560
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   238
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Frame fraInputFiles 
      Caption         =   "Input Options"
      Height          =   2895
      Left            =   0
      TabIndex        =   40
      Top             =   1200
      Width           =   8295
      Begin VB.ComboBox cboFilesToFind 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmGreta.frx":09F4
         Left            =   1680
         List            =   "frmGreta.frx":0A0D
         TabIndex        =   17
         Top             =   2520
         Width           =   3855
      End
      Begin VB.CommandButton cmdBrowseFolders 
         Caption         =   "Browse.."
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.CheckBox chkRecurseSubFolders 
         Alignment       =   1  'Right Justify
         Caption         =   "Recu&rse SubFolders?"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   16
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtFolder 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   "C:\"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.CheckBox chkFindFiles 
         Alignment       =   1  'Right Justify
         Caption         =   "&Find Files"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdFindFiles 
         Caption         =   "Fi&nd..."
         Height          =   375
         Left            =   7200
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin ComctlLib.ListView lvwInputFiles 
         Height          =   1815
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "File Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Modified"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdChangeListView 
         Caption         =   "C&hange..."
         Height          =   375
         Left            =   7200
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton cmdRemoveFromListView 
         Caption         =   "Remo&ve"
         Height          =   375
         Left            =   7200
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdClearListView 
         Caption         =   "C&lear All"
         Height          =   375
         Left            =   7200
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdAddToListView 
         Caption         =   "&Add..."
         Height          =   375
         Left            =   7200
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblFilesToFind 
         Caption         =   "Files To Find:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblFolder 
         Caption         =   "Folder"
         Height          =   255
         Left            =   1200
         TabIndex        =   51
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label lblFilesToAnalyse 
         Caption         =   "Files to Analyse:"
         Height          =   615
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraOutput 
      Caption         =   "Output Options"
      Height          =   1935
      Left            =   0
      TabIndex        =   39
      Top             =   4080
      Width           =   8295
      Begin VB.ComboBox cboIncludeSeperator 
         Height          =   315
         ItemData        =   "frmGreta.frx":0A73
         Left            =   5160
         List            =   "frmGreta.frx":0A80
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cboDispFileNames 
         Height          =   315
         ItemData        =   "frmGreta.frx":0AB5
         Left            =   1680
         List            =   "frmGreta.frx":0AC2
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkOnlyACount 
         Alignment       =   1  'Right Justify
         Caption         =   "Display only a count of &matching lines"
         Height          =   255
         Left            =   5160
         TabIndex        =   24
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtPostLinesToOutput 
         Height          =   285
         Left            =   4200
         TabIndex        =   26
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtPreviousLinesToOutput 
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Text            =   "0"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton cmdBrowseOutputFile 
         Caption         =   "Browse..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   7200
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtOutputFile 
         Enabled         =   0   'False
         Height          =   405
         Left            =   2160
         TabIndex        =   19
         Top             =   360
         Width           =   5055
      End
      Begin VB.CheckBox chkOutputToFile 
         Alignment       =   1  'Right Justify
         Caption         =   "&Output to file?"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkLineNumbers 
         Alignment       =   1  'Right Justify
         Caption         =   "Show &Line Numbers?"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblSeperator 
         Caption         =   "Include Seperator?"
         Height          =   255
         Left            =   3600
         TabIndex        =   48
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lblDispFilenames 
         Caption         =   "Display Filenames?"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Post Lines"
         Height          =   255
         Left            =   4680
         TabIndex        =   46
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Previous Lines and "
         Height          =   255
         Left            =   2760
         TabIndex        =   45
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "If Match is found then output"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1560
         Width           =   2055
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Regular Expression"
      Height          =   1095
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   8295
      Begin VB.CheckBox chkExact 
         Alignment       =   1  'Right Justify
         Caption         =   "&Exact Line Matching?"
         Height          =   255
         Left            =   2520
         TabIndex        =   5
         Top             =   720
         Width           =   2295
      End
      Begin VB.CheckBox chkCase 
         Alignment       =   1  'Right Justify
         Caption         =   "Case &Sensitive?"
         Height          =   255
         Left            =   6240
         TabIndex        =   6
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox chkInvert 
         Alignment       =   1  'Right Justify
         Caption         =   "&Invert Selection?"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtReg 
         Height          =   405
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   6495
      End
      Begin VB.Label Label1 
         Caption         =   "Regular Expression"
         Height          =   375
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   7200
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog cdb1 
      Left            =   0
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblProcess 
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   7200
      Width           =   3495
   End
End
Attribute VB_Name = "frmGreta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 0

'WinGrep version 1.3
'
'For all those UNIX people out there, "grep" should need little
'introduction.  It is simply a powerful tool for extracting lines
'from files.  This is the GUI version for Windows.
'
'To those of you who don't know UNIX, "grep" looks for lines matching a
'regular expression in a file, and if it finds a matching line, then
'it outputs it to a new file.  For an explanation of "regular
'expressions " please examine the VB help on the " Like " operator."
'
'For example if I have a file containing the lines
'
'Good Morning
'and how
'are
'you?
'Good Afternoon
'and Goodbye!
'
'And put the following regular expression into WinGrep and run it
'"Good"
'then the following lines will be added into the output file:
'
'Good Morning
'Good Afternoon
'and Goodbye!
'
'Up to 10 Regular expressions may be specified in WinGrep by putting
'semicolons between the Regular expressions.
'For example, if I was looking for either the words "Good" or "you" then
'my regular expression would be: "Good;you"
'and my output would be:
'
'Good Morning
'you?
'Good Afternoon
'and Goodbye!
'
'Characters in pattern Matches in string:
'? Any single character.
'* Zero or more characters.
'# Any single digit (0-9).
'[charlist] Any single character in charlist.
'[!charlist] Any single character not in charlist.
'
'I know this help file is short - if you need any more information you
'can contact me at bigcalm@hotmail.com
'
'Bug fixes, Improvements, and Suggestions are most welcome too.
'
'
' New functionality for Version 1.1
' 1) Pre-Post line extraction
' 2) Counting facility
' 3) Find-Files is now finished.
' Numerous minor bug-fixes.

' New functionality for Version 1.2
' 1) UNIX Command Line style text box for those moving to WinGrep from Unix or Dos Grep
' 2) Load and Save searches
' 3) "HOT" (quick search) buttons for those regular searches (registry keys)
' 4) Improved file finding dialog on main form.
' 5) DisplayFilenames option changed + Seperator included.
' 6) Can be run from the command line and explorer.  i.e. It processes it's arguments, and now runs using a Sub Main() procedure.
' 7) Post-extraction works correctly now.
' 8) I've put some comments into the Grep functions, 'cos they were getting horribly complex.
' 9) Silent modes for command line running
' 10) Example files included.

' New functionality for Version 1.3 (bug fixes mainly).
' 1) If JustCountThem And DisplayFileNames = OncePerFile then don't have any output if matches = 0 for that file
' 2) More than 10 regular expressions!  I thought 10 was easily enough, but no.
' 3) Improved modularity using compiler directives.
' 4) Removed disgusting common dialog and listview hacks (and replaced with ones almost as bad, sigh).
' 5) Numerous bugs in WorkOutGrepCommand() and PropagateUp() fixed.
' 6) Move to using Windows temp directory for default output files, and add cleardown routine.
' 7) Multiple concurrent searches (in different applications).
' 8) Automatic registration of script files for use with explorer.


' Further improvements that people have asked for....
' 1) Awk! (They don't ask for this directly, but practically, an Awk scripting language is what they need).
'         - They can bog off unless they offer me a fat cheque or some seriously
'           good parsing code or source code for a compiler (not fussy which language, excepting Prolog and LISP).

#Const VBHardCore = False ' If you have Bruce McKinley's VBCore library,
    ' setting this compiler directive to True makes for better looking dialog boxes
    ' available from www.mvps.org/vb/hardcore
Public Running As Boolean
Private Stopping As Boolean
Private DoNotCallWorkOutGrep As Boolean

Private Sub cboDispFileNames_Click()
If DoNotCallWorkOutGrep = False Then
    WorkOutGrepCommand
End If
End Sub

Private Sub cboFilesToFind_Change()
    WorkOutGrepCommand
End Sub

Private Sub cboFilesToFind_Click()
    WorkOutGrepCommand
End Sub

Private Sub cboIncludeSeperator_Change()
    WorkOutGrepCommand
End Sub

Private Sub cboIncludeSeperator_Click()
    WorkOutGrepCommand
End Sub

Private Sub chkCase_Click()
    WorkOutGrepCommand
End Sub

Private Sub chkExact_Click()
    WorkOutGrepCommand
End Sub

Private Sub chkFindFiles_Click()
    If chkFindFiles.Value = 0 Then
        txtFolder.Enabled = False
        cmdBrowseFolders.Enabled = False
        chkRecurseSubFolders.Enabled = False
        lblFilesToFind.Enabled = False
        cboFilesToFind.Enabled = False
        lblFolder.Enabled = False
        lvwInputFiles.Enabled = True
        cmdAddToListView.Enabled = True
        cmdClearListView.Enabled = True
        cmdRemoveFromListView.Enabled = True
        cmdChangeListView.Enabled = True
        cmdFindFiles.Enabled = True
        lblFilesToAnalyse.Enabled = True

    Else
        txtFolder.Enabled = True
        cmdBrowseFolders.Enabled = True
        chkRecurseSubFolders.Enabled = True
        lblFilesToFind.Enabled = True
        cboFilesToFind.Enabled = True
        lblFolder.Enabled = True
        lvwInputFiles.Enabled = False
        cmdAddToListView.Enabled = False
        cmdClearListView.Enabled = False
        cmdRemoveFromListView.Enabled = False
        cmdChangeListView.Enabled = False
        cmdFindFiles.Enabled = False
        lblFilesToAnalyse.Enabled = False
    End If
    WorkOutGrepCommand
End Sub

Private Sub chkInvert_Click()
    WorkOutGrepCommand
End Sub

Private Sub chkLineNumbers_Click()
    WorkOutGrepCommand
End Sub

Private Sub chkOnlyACount_Click()
    If chkOnlyACount.Value = 1 Then
        chkLineNumbers.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        txtPreviousLinesToOutput.Enabled = False
        txtPostLinesToOutput.Enabled = False
        Label5.Enabled = False
    Else
        chkLineNumbers.Enabled = True
        Label3.Enabled = True
        Label4.Enabled = True
        txtPreviousLinesToOutput.Enabled = True
        txtPostLinesToOutput.Enabled = True
        Label5.Enabled = True
    End If
    WorkOutGrepCommand
End Sub

Private Sub chkOutputToFile_Click()
    If chkOutputToFile.Value = 1 Then
        txtOutputFile.Enabled = True
        cmdBrowseOutputFile.Enabled = True
    Else
        txtOutputFile.Enabled = False
        cmdBrowseOutputFile.Enabled = False
    End If
End Sub

Private Sub chkRecurseSubFolders_Click()
    WorkOutGrepCommand
End Sub

Private Sub cmdAddToListView_Click()
Dim i, j As Long
Dim ListItem As ListItem
Dim Path As String
Dim FileName As String
Dim ChkErr As Long

    On Error GoTo ErrHandler:
    cdb1.CancelError = True
    cdb1.DefaultExt = ".txt"
    cdb1.DialogTitle = "Add Files for Grep to search in..."
    cdb1.Filter = "All Files (*.*)|*.*|ASCII Text Files (*.txt)|*.txt|All Text Files (*.txt,*.doc,*.rtf)|*.txt;*.doc;*.rtf|C Files (*.h,*.c)|*.h;*.c|Rich Text Files (*.rtf)|*.rtf|Visual Basic Files (*.bas,*.vbp,*.ctl,*.cls;*.frm,*.pag,*.res)|*.bas,*.vbp,*.ctl,*.cls,*.frm,*.pag,*.res"
    cdb1.FilterIndex = 2
    cdb1.FileName = ""
    cdb1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    cdb1.MaxFileSize = 32000
    cdb1.ShowOpen
    If Len(cdb1.FileName) > 1000 Then
        Me.MousePointer = vbArrowHourglass
    End If
    i = InStr(cdb1.FileName, Chr(0))
    If i > 0 Then
        ' multiselect.  Whoever wrote the common dialog control decided that if you'd chosen multi-select
        ' that he'd return it in an odd way.  Specifically, the first "bit" of the string is the path (up to a NULL chr(0) character)
        ' then the filenames are spaced using NULL characters.  Yuk.
        Path = Mid(cdb1.FileName, 1, i - 1)
        j = InStr(i + 1, cdb1.FileName, Chr(0))
        Do While j > 0
            FileName = Path & "\" & Mid(cdb1.FileName, i + 1, j - i - 1)
            On Error Resume Next
            ChkErr = lvwInputFiles.ListItems(FileName).Text ' if error occurs here it is doesn't occur in the list?  Shut up, it works, ok?
            If Err Then
              Err.Clear
              Set ListItem = lvwInputFiles.ListItems.Add(, FileName, FileName)
              ListItem.SubItems(1) = FileLen(FileName)
              ListItem.SubItems(2) = Format(FileDateTime(FileName), "dd/mm/yy")
            End If
            On Error GoTo 0
            i = j
            j = InStr(i + 1, cdb1.FileName, Chr(0))
        Loop
        FileName = Path & "\" & Mid(cdb1.FileName, i + 1)
        On Error Resume Next
        ChkErr = lvwInputFiles.ListItems(FileName).Text ' if error occurs here it is doesn't occur in the list?  Shut up, it works, ok?
        If Err Then
          Err.Clear
          Set ListItem = lvwInputFiles.ListItems.Add(, FileName, FileName)
          ListItem.SubItems(1) = FileLen(FileName)
          ListItem.SubItems(2) = Format(FileDateTime(FileName), "dd/mm/yy")
        End If
        On Error GoTo 0
    Else
        ' single select
        FileName = cdb1.FileName
        On Error Resume Next
        ChkErr = lvwInputFiles.ListItems(FileName).Text ' if error occurs here it is doesn't occur in the list?  Shut up, it works, ok?
        If Err Then
          Err.Clear
          Set ListItem = lvwInputFiles.ListItems.Add(, FileName, FileName)
          ListItem.SubItems(1) = FileLen(FileName)
          ListItem.SubItems(2) = Format(FileDateTime(FileName), "dd/mm/yy")
        End If
        On Error GoTo 0
    End If
    If Len(cdb1.FileName) > 1000 Then
        Me.MousePointer = vbNormal
    End If
    WorkOutGrepCommand
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
       ' cancel was selected
    Else
        MsgboxForWinGrep Err.Description
    End If
    WorkOutGrepCommand
End Sub

Private Sub cmdBrowseFolders_Click()
Dim strDir As String
#If VBHardCore = True Then
    strDir = BrowseForFolder(Owner:=hWnd, DisplayName:="WinGrep", _
                              Options:=BIF_RETURNONLYFSDIRS, _
                              Title:="Select directory:", _
                              Root:=CSIDL_COMMON_DESKTOPDIRECTORY, _
                              Default:="C:\")
#Else
    strDir = frmChooseDir.GetDirectory
    If Len(strDir) > 0 Then
        txtFolder.Text = strDir
    End If
#End If
End Sub

Private Sub cmdBrowseOutputFile_Click()
    On Error GoTo ErrHandler
    cdb1.CancelError = True
    cdb1.DefaultExt = ".txt"
    cdb1.DialogTitle = "Save Grep Output As..."
    cdb1.FileName = "Output.txt"
    cdb1.Filter = "All Files(*.*)|*.*|Text File(*.txt)|*.txt|Rich Text File (*.rtf)|*.rtf)"
    cdb1.FilterIndex = 2
    cdb1.Flags = cdlOFNExplorer Or cdlOFNLongNames Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    cdb1.ShowSave
    txtOutputFile.Text = cdb1.FileName
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub cmdChangeListView_Click()
Dim i, j As Long
Dim ListItem As ListItem
Dim Path As String
Dim FileName As String
    On Error GoTo ErrHandler:
    cdb1.CancelError = True
    cdb1.DefaultExt = ".txt"
    cdb1.DialogTitle = "Add Files for Grep to search in..."
    cdb1.Filter = "All Files (*.*)|*.*|ASCII Text Files (*.txt)|*.txt|All Text Files (*.txt,*.doc,*.rtf)|*.txt;*.doc;*.rtf|C Files (*.h,*.c)|*.h;*.c|Rich Text Files (*.rtf)|*.rtf|Visual Basic Files (*.bas,*.vbp,*.ctl,*.cls;*.frm,*.pag,*.res)|*.bas,*.vbp,*.ctl,*.cls,*.frm,*.pag,*.res"
    cdb1.FilterIndex = 2
    cdb1.FileName = ""
    cdb1.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist
    cdb1.MaxFileSize = 32000
    cdb1.ShowOpen
    If Len(cdb1.FileName) > 1000 Then
        Me.MousePointer = vbArrowHourglass
    End If
    lvwInputFiles.ListItems.Clear
    i = InStr(cdb1.FileName, Chr(0))
    If i > 0 Then
        ' multiselect.  Whoever wrote the common dialog control decided that if you'd chosen multi-select
        ' that he'd return it in an odd way.  Specifically, the first "bit" of the string is the path (up to a NULL chr(0) character)
        ' then the filenames are spaced using NULL characters.  Yuk.
        Path = Mid(cdb1.FileName, 1, i - 1)
        j = InStr(i + 1, cdb1.FileName, Chr(0))
        Do While j > 0
            FileName = Path & "\" & Mid(cdb1.FileName, i + 1, j - i - 1)
            Set ListItem = lvwInputFiles.ListItems.Add(, FileName, FileName)
            ListItem.SubItems(1) = FileLen(FileName)
            ListItem.SubItems(2) = Format(FileDateTime(FileName), "dd/mm/yy")
            i = j
            j = InStr(i + 1, cdb1.FileName, Chr(0))
        Loop
        FileName = Path & "\" & Mid(cdb1.FileName, i + 1)
        Set ListItem = lvwInputFiles.ListItems.Add(, FileName, FileName)
        ListItem.SubItems(1) = FileLen(FileName)
        ListItem.SubItems(2) = Format(FileDateTime(FileName), "dd/mm/yy")
    Else
        ' single select
        Set ListItem = lvwInputFiles.ListItems.Add(, cdb1.FileName, cdb1.FileName)
        ListItem.SubItems(1) = FileLen(cdb1.FileName)
        ListItem.SubItems(2) = Format(FileDateTime(cdb1.FileName), "dd/mm/yy")
    End If
    If Len(cdb1.FileName) > 1000 Then
        Me.MousePointer = vbNormal
    End If
    WorkOutGrepCommand
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
       ' cancel was selected
    Else
        MsgboxForWinGrep Err.Description
    End If
    WorkOutGrepCommand
End Sub

Private Sub cmdClearListView_Click()
    lvwInputFiles.ListItems.Clear
    WorkOutGrepCommand
End Sub

Private Sub cmdClose_Click()
    If Running = True Then
        Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdFindFiles_Click()
Dim colFindFile As Collection
Dim i, j As Long
Dim ListItem As ListItem
Dim Path As String
Dim FileName As String
Dim ChkErr As Long

    FileFindDialog colFindFile, cdb1.InitDir
    If Not (colFindFile Is Nothing) Then
        ' Need to insert the contents of the find collection into listview
        If colFindFile.Count > 1000 Then
            Me.MousePointer = vbArrowHourglass
        End If
        For i = 1 To colFindFile.Count
            FileName = colFindFile(i)
            On Error Resume Next
            ChkErr = lvwInputFiles.ListItems(FileName).Text ' if error occurs here it is doesn't occur in the list?  Shut up, it works, ok?
            If Err Then
              Err.Clear
              Set ListItem = lvwInputFiles.ListItems.Add(, FileName, FileName)
              ListItem.SubItems(1) = FileLen(FileName)
              ListItem.SubItems(2) = Format(FileDateTime(FileName), "dd/mm/yy")
            End If
            On Error GoTo 0
        Next
        If colFindFile.Count > 1000 Then
            Me.MousePointer = vbDefault
        End If
    End If
    WorkOutGrepCommand
End Sub

Private Sub cmdHotButton1_Click()
Dim tmpStr As String
    If Running = True Then
        Exit Sub
    End If
    tmpStr = GetSetting("WinGrep", "HotButtons", "HotButton1", "")
    If Len(tmpStr) = 0 Then
        Exit Sub
    End If
    cmdHotButton1.ToolTipText = tmpStr
    txtCommand.Text = tmpStr
    cmdPropogateUp_Click
    cmdOk_Click
End Sub

Private Sub cmdHotButton2_Click()
Dim tmpStr As String
    If Running = True Then
        Exit Sub
    End If
    tmpStr = GetSetting("WinGrep", "HotButtons", "HotButton2", "")
    If Len(tmpStr) = 0 Then
        Exit Sub
    End If
    txtCommand.Text = tmpStr
    cmdHotButton2.ToolTipText = tmpStr
    cmdPropogateUp_Click
    cmdOk_Click
End Sub

Private Sub cmdHotButton3_Click()
Dim tmpStr As String
    If Running = True Then
        Exit Sub
    End If
    tmpStr = GetSetting("WinGrep", "HotButtons", "HotButton3", "")
    If Len(tmpStr) = 0 Then
        Exit Sub
    End If
    txtCommand.Text = tmpStr
    cmdHotButton3.ToolTipText = tmpStr
    cmdPropogateUp_Click
    cmdOk_Click
End Sub

Private Sub cmdLoad_Click()
Dim FileName As String

    cdb1.Filter = "All Files (*.*)|*.*|WinGrep " & WinGrepVersion & " Files (*.gre)|*.gre"
    cdb1.FilterIndex = 2
    cdb1.DefaultExt = "*.gre"
    FileName = App.Path
    If Mid(FileName, Len(FileName)) <> "\" Then
        FileName = FileName & "\"
    End If
    cdb1.InitDir = FileName
    cdb1.CancelError = True
    cdb1.FileName = cmdSaveGrep.Tag
    cdb1.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNFileMustExist
    cdb1.DialogTitle = "Load Grep Script From..."
    On Error GoTo ErrHandler
    cdb1.ShowOpen
    LoadGrepFile cdb1.FileName
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
       ' cancel was selected
    Else
        MsgboxForWinGrep Err.Description
    End If
End Sub

Private Sub cmdOk_Click()
    GrepGo
End Sub

Private Sub cmdPropogateUp_Click()
    ConvertGrepStringToOptionBoxes txtCommand.Text
End Sub

Private Sub cmdRemoveFromListView_Click()
Dim ListItem As ListItem
Dim mCol As New Collection ' I'm beginning to hate list views.  Why do you have to delete like this?  Is there a better way?
    For Each ListItem In lvwInputFiles.ListItems
        If ListItem.Selected = True Then
            mCol.Add ListItem
        End If
    Next
    For Each ListItem In mCol
        lvwInputFiles.ListItems.Remove ListItem.Index
    Next
    WorkOutGrepCommand
End Sub

Private Sub cmdReset_Click()
Dim tmpStr As String

    If Running = True Or DoNotCallWorkOutGrep = True Then
        Exit Sub
    End If
    DoNotCallWorkOutGrep = True
    txtReg.Text = ""
    chkInvert.Value = 0
    chkExact.Value = 0
    chkCase.Value = 0
    lvwInputFiles.ListItems.Clear
    chkOutputToFile.Value = 0
    txtOutputFile.Text = ""
    txtOutputFile.Enabled = False
    cmdBrowseOutputFile.Enabled = False
    cboDispFileNames.ListIndex = 0
    chkLineNumbers.Value = 0
    txtPostLinesToOutput.Text = 0
    txtPreviousLinesToOutput.Text = 0
    chkLineNumbers.Enabled = True
    chkOnlyACount.Value = 0
    Label3.Enabled = True
    Label4.Enabled = True
    txtPreviousLinesToOutput.Enabled = True
    txtPostLinesToOutput.Enabled = True
    Label5.Enabled = True
    cboIncludeSeperator.ListIndex = 0
    txtFolder.Enabled = False
    cmdBrowseFolders.Enabled = False
    chkRecurseSubFolders.Enabled = False
    lblFilesToFind.Enabled = False
    cboFilesToFind.Enabled = False
    lvwInputFiles.Enabled = True
    cmdAddToListView.Enabled = True
    cmdClearListView.Enabled = True
    cmdRemoveFromListView.Enabled = True
    cmdChangeListView.Enabled = True
    cmdFindFiles.Enabled = True
    lblFilesToAnalyse.Enabled = True
    lblFolder.Enabled = False
    txtFolder.Text = "C:\"
    cmdOk.Enabled = True
    cboFilesToFind.ListIndex = 0
    chkRecurseSubFolders.Value = 1
    chkFindFiles.Value = 0
    cmdHotButton1.ToolTipText = GetSetting("WinGrep", "HotButtons", "HotButton1", "")
    cmdHotButton2.ToolTipText = GetSetting("WinGrep", "HotButtons", "HotButton2", "")
    cmdHotButton3.ToolTipText = GetSetting("WinGrep", "HotButtons", "HotButton3", "")
    cmdOk.Caption = "Ok"
    Me.Caption = "WinGrep Version " & WinGrepVersion
    DoNotCallWorkOutGrep = False
    WorkOutGrepCommand
End Sub

Private Sub cmdSaveGrep_Click()
Dim FileName As String
Dim FileNum As Long
Dim OutputString As String

    If cmdOk.Enabled = False Then
        Exit Sub
    End If
    cdb1.Filter = "All Files (*.*)|*.*|WinGrep " & WinGrepVersion & " Files (*.gre)|*.gre"
    cdb1.FilterIndex = 2
    cdb1.DefaultExt = "*.gre"
    FileName = App.Path
    If Mid(FileName, Len(FileName)) <> "\" Then
        FileName = FileName & "\"
    End If
    cdb1.InitDir = FileName
    cdb1.CancelError = True
    cdb1.FileName = cmdLoad.Tag
    cdb1.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNNoReadOnlyReturn Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    cdb1.DialogTitle = "Save Grep Script To..."
    On Error GoTo ErrHandler
    cdb1.ShowSave
    
    FileNum = FreeFile
    Open cdb1.FileName For Output Access Write As #FileNum
    OutputString = txtCommand.Text
    Print #FileNum, OutputString
    Close #FileNum
    cmdSaveGrep.Tag = cdb1.FileName
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
       ' cancel was selected
    Else
        MsgboxForWinGrep Err.Description
    End If
End Sub

Private Sub cmdSaveToHS1_Click()
    SaveSetting "WinGrep", "HotButtons", "HotButton1", txtCommand.Text
    cmdHotButton1.ToolTipText = GetSetting("WinGrep", "HotButtons", "HotButton1", "")
End Sub

Private Sub cmdSaveToHS2_Click()
    SaveSetting "WinGrep", "HotButtons", "HotButton2", txtCommand.Text
    cmdHotButton2.ToolTipText = GetSetting("WinGrep", "HotButtons", "HotButton2", "")
End Sub

Private Sub cmdSaveToHS3_Click()
    SaveSetting "WinGrep", "HotButtons", "HotButton3", txtCommand.Text
    cmdHotButton3.ToolTipText = GetSetting("WinGrep", "HotButtons", "HotButton3", "")
End Sub

Private Sub Form_Load()
    Running = False
    pb1.Visible = False
    DoNotCallWorkOutGrep = False
    cmdReset_Click
    WorkOutGrepCommand
    RegisterIfNecessary
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Running = True Then
        Cancel = 1
        If Silence = 2 Then
          Stopping = True
          If FileFinder.ffRunning = True Then
              FileFinder.StopFindFiles
          End If
        End If
    Else
      Me.Hide
      ClearDownTemporaryFiles
    End If
End Sub

Private Sub lvwInputFiles_DblClick()
Dim lRet As Long
    On Error GoTo ErrHandler
    If Not (lvwInputFiles.SelectedItem Is Nothing) Then
        lRet = ShellEx(lvwInputFiles.SelectedItem.Text, essSW_SHOWNORMAL, , , , 0)
    End If
    Exit Sub
ErrHandler:
    Debug.Print Err.Description
    WorkOutGrepCommand
End Sub

Private Sub lvwInputFiles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cmdRemoveFromListView_Click
    End If
    WorkOutGrepCommand
End Sub

Private Sub lvwInputFiles_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim i As Long
Dim ListItem As ListItem
Dim ChkErr As Long
Dim FileName As String

    If Data.GetFormat(vbCFFiles) = True Then
        For i = 1 To Data.Files.Count
            On Error Resume Next
            FileName = Data.Files(i)
            ChkErr = lvwInputFiles.ListItems(FileName).Text ' if error occurs here it is doesn't occur in the list?  Shut up, it works, ok?
            If Err Then
              Err.Clear
              Set ListItem = lvwInputFiles.ListItems.Add(, FileName, FileName)
              ListItem.SubItems(1) = FileLen(FileName)
              ListItem.SubItems(2) = Format(FileDateTime(FileName), "dd/mm/yy")
            End If
            On Error GoTo 0
        Next
    End If
    WorkOutGrepCommand
End Sub

Public Sub ConvertGrepStringToOptionBoxes(GrepString As String)
Dim Argument As String
Dim InFileList As Boolean
Dim tmpStr As String
Dim ListItem As ListItem
Dim Version As String
Dim i As Long

    cmdReset_Click
    chkRecurseSubFolders.Value = 0
    chkCase.Value = 1
    DoNotCallWorkOutGrep = True
    Me.MousePointer = vbHourglass
    InFileList = False
    SplitStringIntoParts GrepString, " "
    Argument = GetNextPartOfSplitString
    Do While Len(Argument) > 0
        Select Case Argument
            Case "grep"
                
            Case "-v"
                chkInvert.Value = 1
            Case "-x"
                chkExact.Value = 1
            Case "-y"
                chkCase.Value = 0
            Case "-l1"
                cboDispFileNames.ListIndex = 1
            Case "-l2"
                cboDispFileNames.ListIndex = 2
            Case "-z1"
                cboIncludeSeperator.ListIndex = 1
            Case "-z2"
                cboIncludeSeperator.ListIndex = 2
            Case "-n"
                chkLineNumbers.Value = 1
            Case "-c"
                chkOnlyACount.Value = 1
            Case Else
                If Argument Like "-pre=*" Then
                    txtPreviousLinesToOutput.Text = Mid(Argument, 6)
                End If
                If Argument Like "-post=*" Then
                    txtPostLinesToOutput.Text = Mid(Argument, 7)
                End If
                If Argument Like "-V=*" Then
                    Version = Mid(Argument, 4)
                End If
                If Argument Like "-reg=*" Then
                    Do While Mid(Argument, Len(Argument), 1) <> """"
                        tmpStr = GetNextPartOfSplitString
                        If Len(tmpStr) > 0 Then
                            Argument = Argument & " " & tmpStr
                        Else
                            Exit Do
                        End If
                    Loop
                    txtReg.Text = Mid(Argument, 7, Len(Argument) - 7)
                    Exit Do
                End If
                If Mid(Argument, 1, 1) <> "-" Then
                    Do While Mid(Argument, Len(Argument), 1) <> """"
                        tmpStr = GetNextPartOfSplitString
                        If Len(tmpStr) > 0 Then
                            Argument = Argument & " " & tmpStr
                        Else
                            Exit Do
                        End If
                    Loop
                    txtReg.Text = Mid(Argument, 2, Len(Argument) - 2)
                    Exit Do
                End If
        End Select
        Argument = GetNextPartOfSplitString
    Loop
    Argument = GetNextPartOfSplitString
    Do While Len(Argument) > 0
        If Argument Like "-find=*" Then
            chkFindFiles.Value = 1
            chkFindFiles_Click
            Do While Mid(Argument, Len(Argument), 1) <> ";" And Len(Argument) > 0
                Argument = Argument & " " & GetNextPartOfSplitString
            Loop
            i = InStr(7, Argument, ";", vbTextCompare)
            If i > 0 Then
                If Mid(Argument, i - 1, 1) = """" Then
                    txtFolder.Text = Mid(Argument, 8, i - 9)
                Else
                    txtFolder.Text = Mid(Argument, 7, i - 7)
                End If
                cboFilesToFind.Text = Mid(Argument, i + 1)
            End If
        Else
            If Argument = "-recurse" Then
                chkRecurseSubFolders.Value = 1
            Else
                If Argument = ">" Then
                    chkOutputToFile.Value = 1
                    txtOutputFile.Text = ""
                    Do
                      txtOutputFile.Text = txtOutputFile.Text & " " & GetNextPartOfSplitString
                    Loop Until Mid(txtOutputFile.Text, Len(txtOutputFile.Text), 1) = """"
                    txtOutputFile.Text = Mid(txtOutputFile.Text, 3, Len(txtOutputFile.Text) - 3)
                    Exit Do
                End If
                If Argument = "<NoFilesSpecified>" Then
                Else
                    If Mid(Argument, 1, 1) = """" Then
                      Do While Mid(Argument, Len(Argument), 1) <> ";" And Len(Argument) > 0
                          Argument = Argument & " " & GetNextPartOfSplitString
                      Loop
                      On Error Resume Next
                      Argument = Mid$(Argument, 2, Len(Argument) - 3)
                      Set ListItem = lvwInputFiles.ListItems.Add(, Argument, Argument)
                      ListItem.SubItems(1) = FileLen(Argument)
                      ListItem.SubItems(2) = Format(FileDateTime(Argument), "dd/mm/yy")
                      On Error GoTo 0
                    Else
                      On Error Resume Next
                      Set ListItem = lvwInputFiles.ListItems.Add(, Argument, Argument)
                      ListItem.SubItems(1) = FileLen(Argument)
                      ListItem.SubItems(2) = Format(FileDateTime(Argument), "dd/mm/yy")
                      On Error GoTo 0
                    End If
                End If
            End If
        End If
        Argument = GetNextPartOfSplitString
    Loop
    DoNotCallWorkOutGrep = False
    Me.MousePointer = vbNormal
    WorkOutGrepCommand
End Sub

Private Sub WorkOutGrepCommand()
Dim GrepString As String
Dim i As Long
Dim tmpStr As String

    If DoNotCallWorkOutGrep = True Then
        Exit Sub
    End If
    Me.MousePointer = vbHourglass
    DoNotCallWorkOutGrep = True
    DoEvents
    GrepString = "grep " & "-V=" & WinGrepVersion & " "
    If chkInvert.Value = 1 Then
        GrepString = GrepString & "-v "
    End If
    If chkExact.Value = 1 Then
        GrepString = GrepString & "-x "
    End If
    If chkCase.Value = 0 Then
        GrepString = GrepString & "-y "
    End If
    Select Case cboDispFileNames.ListIndex
        Case 1
            GrepString = GrepString & "-l1 "
        Case 2
            GrepString = GrepString & "-l2 "
    End Select
    Select Case cboIncludeSeperator.ListIndex
        Case 1
            GrepString = GrepString & "-z1 "
        Case 2
            GrepString = GrepString & "-z2 "
    End Select
    If chkLineNumbers.Value = 1 Then
        GrepString = GrepString & "-n "
    End If
    If chkOnlyACount.Value = 1 Then
        GrepString = GrepString & "-c "
    End If
    If IsNumeric(txtPreviousLinesToOutput.Text) Then
        If txtPreviousLinesToOutput.Text > 0 Then
            GrepString = GrepString & "-pre=" & txtPreviousLinesToOutput.Text & " "
        End If
    End If
    If IsNumeric(txtPostLinesToOutput.Text) Then
        If txtPostLinesToOutput.Text > 0 Then
            GrepString = GrepString & "-post=" & txtPostLinesToOutput.Text & " "
        End If
    End If
    GrepString = GrepString & "-reg=""" & txtReg.Text & """" & " "
    If chkFindFiles.Value = 1 Then
        GrepString = GrepString & "-find=""" & txtFolder.Text & """;"
        SplitStringIntoParts cboFilesToFind.Text, ";"
        tmpStr = GetNextPartOfSplitString
        Do While Len(tmpStr) > 0
            GrepString = GrepString & tmpStr & ";"
            tmpStr = GetNextPartOfSplitString
        Loop
        If chkRecurseSubFolders.Value = 1 Then
            GrepString = GrepString & " -recurse "
        Else
            GrepString = GrepString & " "
        End If
    Else
        If lvwInputFiles.ListItems.Count >= 1 Then
            For i = 1 To lvwInputFiles.ListItems.Count
                GrepString = GrepString & """" & lvwInputFiles.ListItems(i).Text & """; "
            Next
        Else
            GrepString = GrepString & "<NoFilesSpecified> "
        End If
        If chkRecurseSubFolders.Value = 1 Then
            GrepString = GrepString & "-recurse "
        End If
    End If
    
    If Len(txtOutputFile.Text) > 0 Then
        GrepString = GrepString & "> " & """" & txtOutputFile.Text & """"
    End If
    txtCommand.Text = GrepString
    Me.MousePointer = vbNormal
    DoNotCallWorkOutGrep = False
End Sub

Public Sub LoadGrepFile(FileName As String)
Dim InputString As String
Dim FileNum As Long

    On Error GoTo ErrHandler
    FileNum = FreeFile
    Open FileName For Input Access Read As #FileNum
    Line Input #FileNum, InputString
    txtCommand.Text = InputString
    Close #FileNum
    cmdLoad.Tag = cdb1.FileName
    cmdPropogateUp_Click
    Me.Caption = "WinGrep Version " & WinGrepVersion & "(" & cdb1.FileName & ")"
    Exit Sub
ErrHandler:
    If Err.Number = cdlCancel Then
       ' cancel was selected
    Else
        MsgboxForWinGrep Err.Description
    End If
End Sub

Public Sub GrepGo()
Dim ListItem As ListItem
Dim FileNumber As Long
Dim OutputFile As String
Dim lRet As Long
Dim lPrev As Long
Dim lPost As Long
Dim Cancel As Boolean
Dim MatchingLines As Long
Dim TotalMatches As Long
Dim AColl As Collection
Dim i As Long

    '--------------------------------------
    ' Sort out running/stopping
    '--------------------------------------
    If Running = True Then
        Stopping = True
        If FileFinder.ffRunning = True Then
            FileFinder.StopFindFiles
        End If
        Exit Sub
    Else
        Running = True
        Stopping = False
        cmdOk.Caption = "Stop"
    End If
    
    '--------------------------------------
    ' Find files
    '--------------------------------------
    If chkFindFiles.Value = 1 Then
        lvwInputFiles.ListItems.Clear
        Set AColl = FindFiles(txtFolder.Text, cboFilesToFind.Text, -chkRecurseSubFolders.Value, , lblProcess)
        If AColl Is Nothing Then
            cmdOk.Caption = "Ok"
            Running = False
            Stopping = False
            lblProcess.Caption = ""
            Exit Sub
        End If
        For i = 1 To AColl.Count
            Set ListItem = lvwInputFiles.ListItems.Add(, AColl(i), AColl(i))
            ListItem.SubItems(1) = FileLen(AColl(i))
            ListItem.SubItems(2) = Format(FileDateTime(AColl(i)), "dd/mm/yy")
        Next
    End If
    If lvwInputFiles.ListItems.Count = 0 Then
        cmdOk.Caption = "Ok"
        Running = False
        Stopping = False
        lblProcess.Caption = ""
        Exit Sub
    End If

    '--------------------------------------
    ' Run initialization
    '--------------------------------------
    pb1.Visible = True
    pb1.Min = 0
    pb1.Max = lvwInputFiles.ListItems.Count
    pb1.Value = 0
    Me.MousePointer = vbArrowHourglass
    lblProcess.Caption = "Initialising..."
    If IsNumeric(txtPreviousLinesToOutput.Text) Then
        lPrev = txtPreviousLinesToOutput.Text
    Else
        lPrev = 0
    End If
    If IsNumeric(txtPostLinesToOutput.Text) Then
        lPost = txtPostLinesToOutput.Text
    Else
        lPost = 0
    End If
    TotalMatches = 0
    
    '--------------------------------------
    ' Output file handling
    '--------------------------------------
    OutputFile = txtOutputFile.Text
    If Len(OutputFile) = 0 Then
        OutputFile = GetWinTempDir & "WinGrep" & Me.hWnd & ".txt"
    End If
    On Error GoTo ErrHandler
    FileNumber = FreeFile
    Open OutputFile For Output Access Write As #FileNumber
    On Error GoTo 0
    Cancel = False
    
    If Stopping = True Then
        Close #FileNumber
        GoTo StopNow  ' Apologies to all the good programmers out there for this(!)
    End If
    '--------------------------------------
    ' Main Grep Loop
    '--------------------------------------
    For Each ListItem In lvwInputFiles.ListItems
        lblProcess.Caption = "Analysing " & ListItem.Text
        pb1.Value = pb1.Value + 1
        DoEvents
        If Stopping = True Then
            Cancel = True
            Exit For
        End If
        If ListItem.Text = OutputFile Then
            ' Trying to read the Output file as an input file. Yipe.  We don't want this to happen
        Else
            '--------------------------------------
            MatchingLines = Grep(txtReg.Text, ListItem.Text, FileNumber, _
                -chkInvert.Value, -chkExact.Value, -chkCase.Value, _
                cboDispFileNames.ListIndex, -chkLineNumbers.Value, _
                lPrev, lPost, -chkOnlyACount.Value, cboIncludeSeperator.ListIndex)
            '--------------------------------------
            If MatchingLines = -1 Then
                If MsgboxForWinGrep("An error occurred while processing a file " & ListItem.Text & ".  Do you want to continue? (File will be ignored)", vbYesNo) = vbNo Then
                    Cancel = True
                    Exit For
                End If
            Else
                TotalMatches = TotalMatches + MatchingLines
            End If
        End If
    Next
    '--------------------------------------
    ' Output file tidying
    '--------------------------------------
    If chkOnlyACount.Value = 1 And cboDispFileNames.ListIndex = 0 Then
        Print #FileNumber, Format(TotalMatches, "0")
    End If
    Close #FileNumber
    If Cancel = True Then
        Running = False
        pb1.Visible = False
        Me.MousePointer = vbDefault
        lblProcess.Caption = ""
        cmdOk.Caption = "Ok"
        If Silence = 2 Then
          Unload Me
        End If
        Exit Sub
    End If
    '--------------------------------------
    ' Launch file
    '--------------------------------------
    On Error GoTo ErrHandler2
    lblProcess.Caption = "Launching File..."
    If Len(txtOutputFile.Text) = 0 Then
        lRet = ShellEx(OutputFile, essSW_SHOWNORMAL, , , , 0)
    Else
        MsgboxForWinGrep "Wrote output to " & OutputFile
    End If
    '--------------------------------------
    ' Tidy Up
    '--------------------------------------
    pb1.Value = 0
    pb1.Visible = False
    lblProcess.Caption = ""
    On Error GoTo 0
    Running = False
    cmdOk.Caption = "Ok"
    Me.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Me.MousePointer = vbDefault
    MsgboxForWinGrep "Could not open output file"
    Running = False
    pb1.Visible = False
    cmdOk.Caption = "Ok"
    If Silence = 2 Then
      Unload Me
    End If
    Exit Sub
ErrHandler2:
    Me.MousePointer = vbDefault
    MsgboxForWinGrep Err.Description
    Running = False
    pb1.Visible = False
    cmdOk.Caption = "Ok"
    If Silence = 2 Then
      Unload Me
    End If
    Exit Sub
StopNow:
    Me.MousePointer = vbDefault
    Running = False
    Stopping = False
    pb1.Visible = False
    cmdOk.Caption = "Ok"
    If Silence = 2 Then
      Unload Me
    End If
End Sub


Private Sub txtFolder_Change()
    If Len(Dir(txtFolder.Text, vbDirectory)) > 0 Then
        cmdOk.Enabled = True
    Else
        cmdOk.Enabled = False
    End If
    WorkOutGrepCommand
End Sub

Private Sub txtFolder_LostFocus()
    If Mid(txtFolder.Text, Len(txtFolder.Text)) <> "\" Then
        txtFolder.Text = txtFolder.Text & "\"
    End If
End Sub

Private Sub txtOutputFile_Change()
    WorkOutGrepCommand
End Sub

Private Sub txtPostLinesToOutput_Change()
    WorkOutGrepCommand
End Sub

Private Sub txtPreviousLinesToOutput_Change()
    WorkOutGrepCommand
End Sub

Private Sub txtReg_Change()
    WorkOutGrepCommand
End Sub

Public Sub ShowJustProgressStuff()
Dim Control As Control
    On Error Resume Next
    For Each Control In Me.Controls
        Control.Visible = False
    Next
    On Error GoTo 0
    Me.Width = pb1.Width
    Me.Height = 0
    Me.Height = Me.Height + lblProcess.Height + pb1.Height
    lblProcess.Top = 0
    lblProcess.Left = 0
    lblProcess.Visible = True
    pb1.Top = lblProcess.Top + lblProcess.Height
    pb1.Left = 0
    pb1.Visible = True
End Sub

' Tidy up routine for temporary files - removes any WinGrep temporary files older than one day.
Public Sub ClearDownTemporaryFiles()
Dim FileColl As Collection
Dim i As Long
  Set FileColl = FindFiles(GetWinTempDir, "WinGrep*.txt", False, DateAdd("s", -86400, Now))
  For i = 1 To FileColl.Count
    Kill FileColl(i)
  Next
End Sub

Public Sub RegisterIfNecessary()
Dim Registered As Boolean
Dim OutputFile As String
Dim FileNumber As Long

  If FileExists(App.Path & "\" & App.EXEName & ".exe") Then
    Registered = GetSetting("WinGrep", "Registered", "Registered", False)
    If Registered = False Then
      ' Yeah I know, I can't be bothered to do all the API hacking to do this properly
      ' Don't copy my method please!
      OutputFile = App.Path & "\Greta13.reg"
      FileNumber = FreeFile
      Open OutputFile For Output As #FileNumber
      Print #FileNumber, "REGEDIT4"
      Print #FileNumber, ""
      Print #FileNumber, "[HKEY_CLASSES_ROOT\grefile]"
      Print #FileNumber, "@=""Grep Script File"""
      Print #FileNumber, """EditFlags""=hex:00,00,00,00"
      Print #FileNumber, """AlwaysShowExt""="""""
      Print #FileNumber, ""
      Print #FileNumber, "[HKEY_CLASSES_ROOT\grefile\Shell]"
      Print #FileNumber, "@="""""
      Print #FileNumber, ""
      Print #FileNumber, "[HKEY_CLASSES_ROOT\grefile\Shell\open]"
      Print #FileNumber, ""
      Print #FileNumber, "[HKEY_CLASSES_ROOT\grefile\Shell\open\command]"
      Print #FileNumber, "@=""\""" & QSAR(App.Path, "\", "\\") & "\\" & App.EXEName & ".exe\"" -s2 -f=\""%1\"""""
      Print #FileNumber, ""
      Print #FileNumber, "[HKEY_CLASSES_ROOT\grefile\Shell\edit]"
      Print #FileNumber, "@=""Edit"""
      Print #FileNumber, ""
      Print #FileNumber, "[HKEY_CLASSES_ROOT\grefile\Shell\edit\command]"
      Print #FileNumber, "@=""\""" & QSAR(App.Path, "\", "\\") & "\\" & App.EXEName & ".exe\"" -e -f=\""%1"""""
      Print #FileNumber, ""
      Print #FileNumber, "[HKEY_CLASSES_ROOT\grefile\DefaultIcon]"
      Print #FileNumber, "@=""" & QSAR(GetWinSystemDir(), "\", "\\") & "\\shell32.dll,43"""
      Print #FileNumber, ""
      Close #FileNumber
      ShellEx OutputFile
      SaveSetting "WinGrep", "Registered", "Registered", True
    End If
  End If
End Sub
Public Function FileExists(Path As String) As Boolean
    FileExists = Len(Dir(Path)) > 0
End Function

