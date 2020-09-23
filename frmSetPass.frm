VERSION 5.00
Begin VB.Form frmSetPass 
   Caption         =   "Set Password"
   ClientHeight    =   1605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   Icon            =   "frmSetPass.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1605
   ScaleWidth      =   4020
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change Password"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3855
   End
   Begin VB.TextBox txtNewConfirm 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.TextBox txtNewPass 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtOldPass 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm New Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "New Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password:"
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   165
      Width           =   1095
   End
End
Attribute VB_Name = "frmSetPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
Dim inputPass As String
Open "C:\pltQr01w.sys" For Input As #2
Input #2, inputPass
Close #2
If inputPass = txtOldPass.Text Then
    If txtNewPass.Text = txtNewConfirm.Text Then
        Open "C:\pltQr01w.sys" For Output As #2
        Write #2, txtNewPass.Text
        Close #2
    Else
        MsgBox "New passwords do not match"
    End If
Else
    MsgBox "Invalid old password"
End If
End Sub
