VERSION 5.00
Begin VB.Form FrmChannel 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Channel"
   ClientHeight    =   4245
   ClientLeft      =   2040
   ClientTop       =   2130
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   2685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   375
      Left            =   960
      TabIndex        =   7
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
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
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   2520
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Quit"
      Height          =   330
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Join"
      Height          =   330
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
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
      Height          =   2190
      ItemData        =   "FrmChannel.frx":0000
      Left            =   120
      List            =   "FrmChannel.frx":000D
      TabIndex        =   0
      Top             =   240
      Width           =   2460
   End
   Begin VB.TextBox txtchannel 
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
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   2520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Channel"
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
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Channel Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2535
   End
End
Attribute VB_Name = "FrmChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    BattleMain.Text1.text = "/Channel " & txtchannel.text
    BattleMain.Send_Click
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
List1.AddItem Text1.text
End Sub

Private Sub List1_Click()
    txtchannel.text = List1.List(List1.ListIndex)
End Sub

