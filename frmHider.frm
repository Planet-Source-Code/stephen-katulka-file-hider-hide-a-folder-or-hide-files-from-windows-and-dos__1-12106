VERSION 5.00
Begin VB.Form frmHider 
   Caption         =   "Folder Hider"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "frmHider.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton cmdSetPass 
      Caption         =   "Set Password"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   5040
      Width           =   5535
   End
   Begin VB.ListBox Hidden 
      Height          =   4350
      ItemData        =   "frmHider.frx":08CA
      Left            =   3000
      List            =   "frmHider.frx":08CC
      TabIndex        =   3
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "&Show"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   4440
      Width           =   2535
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   2535
   End
   Begin VB.DirListBox Dirs 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Password"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   5400
      Width           =   2655
   End
End
Attribute VB_Name = "frmHider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================
'=                                                                           =
'=                  Folder Hider v1.0 by Stephen Katulka                     =
'=                                                                           =
'=============================================================================
' I made this program as a favor to a friend whose boyfriend was prying through
' her files. This is my first submission, and isn't very well documented, but
' doesn't do anything too complicated. It shows how to manipulate the
' FileSystemObject attributes.


Private Sub cmdHide_Click()
On Error GoTo errHandler
Dim pass As String
Open "C:\pltQr01w.sys" For Input As #2
Input #2, pass
Close #2
hideFile:
If txtPassword.Text = pass Then
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(Dirs.Path)
        'set the the "Hidden" bit in the attribute byte
        f.Attributes = 2
    Open "C:\flg.log" For Append As #1
    Write #1, Dirs.Path
    Close #1
    Set f = Nothing
    Set fs = Nothing
    Dirs.Refresh
Else
    MsgBox "Invalid Password"
    txtPassword.SetFocus
End If
Form_Load
Dirs.Refresh
Exit Sub
errHandler:
Open "C:\pltQr01w.sys" For Append As #3
pass = InputBox("What would you like to use as your initial password?", "Password Select")
Write #3, pass
Close #3
GoTo hideFile
Exit Sub
End Sub

Private Sub cmdShow_Click()
On Error GoTo errHandler
Dim pass As String
Open "C:\pltQr01w.sys" For Input As #2
Input #2, pass
Close #2
If txtPassword.Text = pass Then
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(Hidden.List(Hidden.ListIndex))
    If f <> Drive1.Drive Then
        'Turn off the "Hidden" bit in the attribute byte
        f.Attributes = 0
    End If
    Set f = Nothing
    Set fs = Nothing
    Hidden.RemoveItem (Hidden.ListIndex)
    If Hidden.ListCount > 0 Then
        Open "C:\flg.log" For Output As #1
        Dim i As Integer
        For i = 0 To Hidden.ListCount - 1
            Write #1, Hidden.List(i)
        Next
        Close #1
    Else
        Set fs = CreateObject("Scripting.FileSystemObject")
        fs.DeleteFile "C:\flg.log"
        Set fs = Nothing
    End If
Else
    MsgBox "Invalid Password"
    txtPassword.SetFocus
End If
    Form_Load
    Dirs.Refresh
Exit Sub
errHandler:
Exit Sub
End Sub

Private Sub Drive1_Change()
    Dirs.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Hidden.Clear
    Dim file As String
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(Dirs.Path)
If fs.FileExists("C:\flg.log") Then
    Open "C:\flg.log" For Input As #1
While Not EOF(1)
    Input #1, file
    Hidden.AddItem file
Wend
    Close #1
End If
End Sub
