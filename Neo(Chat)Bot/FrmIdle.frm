VERSION 5.00
Begin VB.Form FrmIdle 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anti Idle Message Setup"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Send Idle Message"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   4320
         TabIndex        =   2
         Text            =   "/me is an Atomic Bot - by At0m and Neogenic"
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ok"
         Height          =   375
         Left            =   4320
         TabIndex        =   1
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "Anti Idle Message"
         Height          =   495
         Left            =   4320
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmIdle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Open "Idle.ini" For Output As 1
Unload Me
End Sub

Private Sub vsFlexArray1_SelChange()

End Sub
