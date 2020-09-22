VERSION 5.00
Begin VB.Form frmUsers 
   Caption         =   "User Database"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form2"
   ScaleHeight     =   3840
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   2565
      ItemData        =   "frmUsers.frx":0000
      Left            =   120
      List            =   "frmUsers.frx":0010
      TabIndex        =   5
      Top             =   720
      Width           =   4455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Current User To Grant Access To:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmUsers.Hide
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
List1.Text = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        KeyAscii = 0
    List1.AddItem Text1.Text
    Text1.Text = ""
    End If
End Sub
