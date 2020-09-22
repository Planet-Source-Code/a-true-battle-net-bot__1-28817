VERSION 5.00
Begin VB.Form FrmBinCon 
   BackColor       =   &H80000004&
   Caption         =   "Binary Connection"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   Icon            =   "FrmBinCon.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Launch"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Launch"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Launch"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Launch Warcraft Game Client, MAKE SURE THE BINARY CONNECTION RECEIVER IS OPEN!!!"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Launch Starcraft Game Client, MAKE SURE THE BINARY CONNECTION RECEIVER IS OPEN!!!"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Launch Binary Connection Receiver, MAKE SURE REGISTRY FILES HAVE BEEN IMPORTED!!!"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "FrmBinCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Shell "C:\Program Files\Atomic Binary\Atomic Binary Receiver.exe"
MsgBox "Binary Connection Receiver open, Closing Atomic Bot..."
Unload BattleMain
End Sub

Private Sub Command2_Click()
Shell "C:\Program Files\starcraft\starcraft.exe"
End Sub

Private Sub Command3_Click()
Shell "C:\Program Files\Warcraft II BNE\Warcraft II BNE.exe"
End Sub

