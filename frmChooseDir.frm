VERSION 5.00
Begin VB.Form frmChooseDir 
   Caption         =   "Choose Directory..."
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmChooseDir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.DriveListBox drv1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   4335
   End
End
Attribute VB_Name = "frmChooseDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Option Compare Text

Dim MyPath As String

Private Sub cmdCancel_Click()
    MyPath = ""
    Unload Me
End Sub

Private Sub cmdOk_Click()
    MyPath = Dir1.Path & "\"
    Unload Me
End Sub


Private Sub drv1_Change()
Dim OldPath As String
    OldPath = Dir1.Path
    On Error GoTo ErrHandler
    Dir1.Path = drv1.Drive
    Exit Sub
ErrHandler:
    #If PartOfWinGrep = 1 Then
      MsgboxForWinGrep "Error changing drive to " & drv1.Drive & ": " & Err.Description
    #Else
      MsgBox "Error changing drive to " & drv1.Drive & ": " & Err.Description
    #End If
    Dir1.Path = OldPath
    drv1.Drive = Mid(OldPath, 1, 3)
    Exit Sub
End Sub

Public Function GetDirectory(Optional pStartPath As String = "C:\") As String
    If Len(Dir(pStartPath)) = 0 Then
        pStartPath = "C:\"
    End If
    Dir1.Path = pStartPath
    drv1.Drive = Mid(pStartPath, 1, 3)
    Me.Show
    GetDirectory = MyPath
End Function
