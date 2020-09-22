VERSION 5.00
Begin VB.Form FrmRealm 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server Selector"
   ClientHeight    =   3075
   ClientLeft      =   2310
   ClientTop       =   3165
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   2685
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtrealm 
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
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   75
      TabIndex        =   4
      Top             =   2280
      Width           =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   330
      Left            =   825
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.ListBox List1 
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
      ForeColor       =   &H000000FF&
      Height          =   1740
      ItemData        =   "FrmRealm.frx":0000
      Left            =   75
      List            =   "FrmRealm.frx":0010
      TabIndex        =   0
      Top             =   375
      Width           =   2565
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   915
      TabIndex        =   1
      Top             =   75
      Width           =   855
   End
End
Attribute VB_Name = "FrmRealm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    RealmServer = txtrealm.text
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub List1_Click()
    txtrealm.text = List1.List(List1.ListIndex)
End Sub
