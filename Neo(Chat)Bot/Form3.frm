VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H80000004&
   Caption         =   "Select"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form3"
   ScaleHeight     =   4320
   ScaleWidth      =   4260
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3240
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4035
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   180
      Pattern         =   "*.mp3"
      TabIndex        =   4
      Top             =   2340
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "Ok"
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   855
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOk_Click()
File1.Path = Dir1.Path
If File1.ListCount <> 0 Then
    For tel = 1 To File1.ListCount
        File1.ListIndex = tel - 1
        
        
        
        If Len(Dir1.Path) > 3 Then
            Form1.List1.AddItem Dir1.Path & "\" & File1.FileName
        Else
           'Exit For
            'MsgBox "You can't add a drive, only folders", vbOKOnly, "Error"
           'Exit Sub
        Form1.List1.AddItem Dir1.Path & File1.FileName
        End If
    Next tel
            Unload Me
Else
    MsgBox "No files were found in specific folder", vbOKOnly, "Error"
    Unload Me
End If
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub
