VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   2520
   ClientLeft      =   2490
   ClientTop       =   2835
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Channt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Quit 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2040
      Width           =   1005
   End
   Begin VB.TextBox UserNamet 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   150
      TabIndex        =   0
      Top             =   315
      Width           =   2055
   End
   Begin VB.TextBox PassWordt 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   150
      PasswordChar    =   "§"
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton Login 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Channel"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   75
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   675
      Width           =   855
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Login_Click()
    If UserNamet.text = "" Then
        MsgBox "Must Enter a Username to Login", vbOKOnly
        Exit Sub
    End If
    Username = UserNamet.text
    Password = PassWordt.text
    Select Case UserNamet.text
        Case "Anonymous", "anonymous", "Guest", "guest":
            BattleMain.ProceedLogin
            Unload Me
            Exit Sub
    End Select
    If Password <> "" Then
        BattleMain.ProceedLogin
    Else
        MsgBox "Must Enter a Password to Login", vbOKOnly
        Exit Sub
    End If
FrmLogin.Hide
End Sub
Private Sub Quit_Click()
Unload Me
End Sub


