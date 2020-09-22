VERSION 5.00
Begin VB.Form frmFindFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find Files"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   Icon            =   "frmFindFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox drv1 
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   3495
   End
   Begin VB.CheckBox chkRecurseSubFolders 
      Alignment       =   1  'Right Justify
      Caption         =   "Recurse SubFolders?"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Value           =   1  'Checked
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Now"
      Default         =   -1  'True
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox cboRegularExpression 
      Height          =   315
      ItemData        =   "frmFindFiles.frx":0442
      Left            =   0
      List            =   "frmFindFiles.frx":045B
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin VB.DirListBox dirLB1 
      Height          =   1890
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label lblProcessUpdate 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   3840
      Width           =   3495
   End
End
Attribute VB_Name = "frmFindFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Option Compare Text
'  - FireClaw.  bigcalm@hotmail.com
Private Running As Boolean
Private mCollection As Collection

Private Sub cmdCancel_Click()
    If Running = False Then
        Unload Me
    End If
End Sub

Private Sub cmdFind_Click()
    If Running = False Then
        Running = True
        Me.Height = Me.Height + lblProcessUpdate.Height
        Set mCollection = FindFiles(dirLB1.Path, cboRegularExpression.Text, -chkRecurseSubFolders.Value, , lblProcessUpdate)
        Running = False
        Me.Height = Me.Height - lblProcessUpdate.Height
        Unload Me
    End If
End Sub

Private Sub drv1_Change()
Dim OldPath As String
    OldPath = dirLB1.Path
    On Error GoTo ErrHandler
    dirLB1.Path = drv1.Drive
    Exit Sub
ErrHandler:
    #If PartOfWinGrep = 1 Then
      MsgboxForWinGrep "Error changing drive to " & drv1.Drive & ": " & Err.Description
    #Else
      MsgBox "Error changing drive to " & drv1.Drive & ": " & Err.Description
    #End If
    dirLB1.Path = OldPath
    drv1.Drive = Mid(OldPath, 1, 3)
    Exit Sub
End Sub

Private Sub Form_Load()
    Running = False
    Set mCollection = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Running = True Then
        Cancel = 1
    End If
End Sub

' After calling this function pCollection should be filled in with a Collection of Strings (Filenames with Paths)
' Usage:
' dim x as collection
' findfilesform(x)
' debug.print x.count  (will show number of elements in form).
' If the user cancels, the collection will be set to nothing
Public Sub FindFilesForm(ByRef pCollection As Collection, Optional pStartPath As String = "")
    If Len(pStartPath) > 0 Then
        dirLB1.Path = pStartPath
    End If
    Me.Show vbModal
    Set pCollection = mCollection
End Sub
