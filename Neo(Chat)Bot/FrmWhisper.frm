VERSION 5.00
Begin VB.Form FrmWhisper 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3585
   ForeColor       =   &H80000004&
   LinkTopic       =   "Form1"
   ScaleHeight     =   1605
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000C0&
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "User to whisper:"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "FrmWhisper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    BattleMain.Text1.text = "/w " & Text2.text & " " & Text1.text
    BattleMain.Send_Click
    BattleMain.Text3.text = "You Whispered: " & Text1.text & " To: " & Text2.text & vbCrLf
    BattleMain.Text3.SelStart = 0
    BattleMain.Text3.SelColor = QBColor(8)
End Sub

Private Sub Command2_Click()
FrmWhisper.Hide
End Sub
