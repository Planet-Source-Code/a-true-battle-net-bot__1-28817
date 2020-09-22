VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Binary Connection"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   4275
      Left            =   0
      ScaleHeight     =   4215
      ScaleWidth      =   5565
      TabIndex        =   3
      Top             =   0
      Width           =   5625
      Begin VB.ListBox Lstlog 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2040
         TabIndex        =   8
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   600
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "6112"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Txtlocip 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Text            =   "useast.battle.net"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Port :"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Server:"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox picon 
      Height          =   615
      Index           =   2
      Left            =   7920
      Picture         =   "Form11.frx":0442
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7320
      Top             =   2400
   End
   Begin VB.PictureBox picon 
      Height          =   615
      Index           =   1
      Left            =   7920
      Picture         =   "Form11.frx":0853
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picon 
      Height          =   615
      Index           =   0
      Left            =   7920
      Picture         =   "Form11.frx":0C67
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSWinsockLib.Winsock Ws2 
      Index           =   0
      Left            =   8160
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Ws1 
      Index           =   0
      Left            =   7800
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'This Code Was Fully Built By Asaf Yarkoni Isreal Nov 2000
'Please Feel Free to Learn From it Use It Or Just Haveing it
'For Comments Questions And Just for .. Please Contact me at yarkoni1@netvision.net.il




Dim WsI1(100) As Boolean
Dim WsI2(100) As Boolean

Dim TotalIn As Long
Dim TotalOut As Long
 

 



Private Sub ConStat_KillCon(Index As Integer)
                  'Kill the connection when the user has press the 'Kill Button'
Ws1(Index).Close
Ws1_Close Index
End Sub

Private Sub Form_Load()

Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
Ws1(0).LocalPort = 6112
Ws1(0).Listen                 ' Listen on the Termianl port --3389

ConStat(0).Top = 0 - ConStat(0).Height + 50
Vs1.Max = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub LblMnu_Click(Index As Integer)
Select Case Index
Case 0
FrmSecurity.Show vbModal, Me              'Show the security menu
Case 1
MsgBox "Not Yet Avail..."                 'Show the Prop Menu
Case 2
FrmAbout.Show vbModal
End Select
End Sub

Private Sub LblLeftC_Click()

End Sub

Private Sub Timer1_Timer()
Dim TmpK As Long
LblConnected.Caption = Ws1.Count - 1              'Update The counters
LblLeftC.Caption = 101 - Ws1.Count

TmpK = TotalIn / 1000
LblTotalIn = TmpK & "k"

TmpK = TotalOut / 1000
LblTotalOUT = TmpK & "k"
End Sub

Private Sub Vs1_Change()
PicLay.Top = -Vs1.Value       ' Move our PictureBox Space
End Sub

Private Sub Ws1_Close(Index As Integer)
On Error Resume Next
' Disconnect all Var involeved With The Index Above
ConStat(Index).ConDir = 3
Delay 1

Unload Ws1(Index)
Unload Ws2(Index)

Unload ConStat(Index)
  


WsI1(Index) = False
WsI2(Index) = False
End Sub

Private Sub Ws1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next

If ChkIp(Ws1(Index).RemoteHostIP) = False Then Exit Sub 'Check For Security Settings

For i = 1 To 100
If WsI1(i) = False Then
WsI1(i) = True
Exit For
End If
Next i
Load Ws2(i)                       'Load the Core Winsock Control
Ws2(i).Connect Txtlocip, 6112     'Establish Connection To The Terminal Sever
Do Until Ws2(i).State = 7         ' Wait for it to connect to you terminal server
DoEvents
v = v + 1
Loop

Load Ws1(i)
Ws1(i).Accept requestID
 

Load ConStat(i)                  'Load the Monitor User Control

ConStat(i).IPAd = "Ip: " & Ws1(i).RemoteHostIP
ConStat(i).IndNum = i
ConStat(i).Top = ConStat(i - 1).Top + ConStat(i - 1).Height - 50
ConStat(i).Visible = True



PicLay.Height = ConStat(i).Height * ConStat.Count

 
Vs1.Max = PicLay.Height / 4                          'Adjust the Slide Bar
 

Lstlog.AddItem "#" & i & " " & Date & "-" & Time & " - " & Ws1(i).RemoteHostIP    'write the log ***


End Sub

 
Private Sub Ws1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
' Transfer All data from the Client to the host

ConStat(Index).ConDir = 0
Dim an As Variant
Ws1(Index).GetData an

 
Ws2(Index).SendData an
 

End Sub

Private Sub Ws1_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
TotalIn = TotalIn + bytesSent
End Sub

Private Sub Ws2_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next
'Transfer All data from the Server to the client

ConStat(Index).ConDir = 1
Dim an As Variant
Ws2(Index).GetData an
Ws1(Index).SendData an


End Sub

Private Sub Ws2_SendProgress(Index As Integer, ByVal bytesSent As Long, ByVal bytesRemaining As Long)
TotalOut = TotalOut + bytesSent

End Sub
