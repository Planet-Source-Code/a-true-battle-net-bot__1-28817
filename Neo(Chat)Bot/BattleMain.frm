VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form BattleMain 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Neo(Chat)Bot 1.0 - Chatbot - Public Release - By Neogenic"
   ClientHeight    =   6930
   ClientLeft      =   1605
   ClientTop       =   2235
   ClientWidth     =   10455
   Icon            =   "BattleMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check6 
      Caption         =   "Use Interbot Commands"
      Height          =   315
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Enable Protection"
      Height          =   315
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Enable User Control"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Width           =   1695
   End
   Begin VB.Timer Timer3 
      Interval        =   40000
      Left            =   1560
      Top             =   1440
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Show Whispers on Seperate Form"
      Height          =   315
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show Timestamps"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Channel Greetings"
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   9240
      TabIndex        =   10
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Send"
      Height          =   375
      Left            =   7680
      MaskColor       =   &H00000000&
      Picture         =   "BattleMain.frx":0442
      TabIndex        =   9
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Interval        =   58000
      Left            =   3600
      Top             =   2400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remove"
      Height          =   255
      Left            =   9600
      TabIndex        =   8
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Left            =   7680
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   1380
      ItemData        =   "BattleMain.frx":0C77
      Left            =   7680
      List            =   "BattleMain.frx":0C79
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   4320
      Width           =   2655
   End
   Begin RichTextLib.RichTextBox Text3 
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   4560
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      _Version        =   393217
      BackColor       =   16777215
      Appearance      =   0
      TextRTF         =   $"BattleMain.frx":0C7B
   End
   Begin RichTextLib.RichTextBox Text2 
      Height          =   5565
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   9816
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"BattleMain.frx":0D44
   End
   Begin VB.Timer Timer1 
      Left            =   6240
      Top             =   0
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6240
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   7455
   End
   Begin VB.ListBox Roomlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   3630
      ItemData        =   "BattleMain.frx":0E0D
      Left            =   7680
      List            =   "BattleMain.frx":0E0F
      MouseIcon       =   "BattleMain.frx":0E11
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Buddy List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   7680
      TabIndex        =   6
      Top             =   4080
      Width           =   2655
   End
   Begin VB.Label Channel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Channel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   7680
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu Cone 
         Caption         =   "&Connect"
      End
      Begin VB.Menu Disc 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu joinchan 
         Caption         =   "&Join Channel"
      End
      Begin VB.Menu whisp 
         Caption         =   "&Whisper"
      End
      Begin VB.Menu chanlist 
         Caption         =   "&Channel List"
      End
      Begin VB.Menu mp3 
         Caption         =   "&Mp3 Player"
      End
   End
   Begin VB.Menu bincon 
      Caption         =   "Bi&nary Connection"
   End
End
Attribute VB_Name = "BattleMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ind As Long
Dim Channelstring
Dim OldIndex

Private Sub bincon_Click()
FrmBinCon.Show
End Sub

Private Sub ChanList_Click()
FrmChannel.Show 1
If Winsock1.State = disconnected Then disabled = True
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Form1.Show
End If
End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
frmUsers.Show
Check3.Value = 1
Check3.Enabled = False
End If
If Check4.Value = 0 Then
Check3.Value = 0
Check3.Enabled = True
End If

End Sub

Private Sub Command1_Click()
List1.AddItem Roomlist.Text
End Sub



Private Sub Command2_Click()
Send_Click
End Sub

Private Sub Command3_Click()
removelitem List1
End Sub
Private Sub Command4_Click()
Text1.Text = Encrypt(Text1) + " - Atomic Bot Encryption"

End Sub

Private Sub Command5_Click()
If Text1.Text = "/home" Then
ExecuteCode ("/join " & FrmLogin.Channt.Text)
End If
If Text1.Text = "/server" Then
ExecuteCode ("Server Set To: " & RealmServer)
End If
If Text1.Text = "/version" Then
ExecuteCode ("/me is logged on Neo(Chat)Bot, version 1.0c")
End If
If Text1.Text = "/chaninfo" Then
ExecuteCode ("Current Channel is: " & Channelstring & ", " & "Containing " & Roomlist.ListCount & " Users.")
End If

End Sub

Private Sub Cone_Click()
    On Error Resume Next
    With Winsock1
        .Protocol = sckTCPProtocol
        .RemoteHost = RealmServer
        .RemotePort = "6112"
        .Connect
        ConnectedToServer = True
    End With
End Sub

Private Sub Disc_Click()
    ConnectedToServer = False
    Timer1.Interval = 0
    Roomlist.Clear
    Text2.Text = ""
    DoEvents
    Winsock1.Close
End Sub


Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
RealmServer = "useast.battle.net"
End Sub
Public Sub ProceedLogin()
    MousePointer = 11
    With Winsock1
        .SendData Chr(3) & Chr(4) & U_ID & vbCrLf & U_Pass & vbCrLf
        .SendData Username & vbCrLf
        .SendData Password & vbCrLf
    End With
    DoEvents
    Timer1.Interval = 500
    Me.Caption = "Neo(Chat)Bot - Logged in as: " & Username & " Using: CHAT"
End Sub

Private Sub idlecfg_Click()
FrmIdle.Show
End Sub

Private Sub joinchan_Click()
Frmchanjoin.Show
End Sub

Private Sub List1_Click()
nameLen = Len(List1.List(List1.ListIndex))
    listname = Trim(Mid(List1.List(List1.ListIndex), 1, nameLen - 7))
    If List1.Selected(List1.ListIndex) = True Then
        ExecuteCode ("/f add " & listname)
    Else
        ExecuteCode ("/whois " & listname)
    End If
End Sub

Private Sub mnuRealm_Click()
FrmRealm.Show 1
End Sub
Public Sub Send_Click()
    
    Winsock1.SendData Text1.Text & vbCrLf
    If Mid(Text1.Text, 1, 1) <> "/" Then
        Text3.Text = "[" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "] - "
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbBlue
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
        Text3.Text = "You said - "
        Text3.SelStart = 0
        Text3.SelLength = Len(Text3.Text)
        Text3.SelColor = vbRed
        Text2.SelStart = Len(Text2.Text)
        Text2.SelRTF = Text3.TextRTF
        Text3.Text = Text1.Text & vbCrLf
        Text3.SelStart = 0
        Text3.SelLength = Len(Text3.Text)
        Text3.SelColor = vbWhite
        Text2.SelStart = Len(Text2.Text)
        Text2.SelRTF = Text3.TextRTF
    End If
    Text1.Text = ""
End Sub
Private Function RetreiveCodePos(strdata) As Long
    mypost = InStr(1, strdata, "0000")
    If mypost <> 0 Then RetreiveCodePos = mypost
    mypost = InStr(1, strdata, "0002")
    If mypost <> 0 Then RetreiveCodePos = mypost
    mypost = InStr(1, strdata, "0010")
    If mypost <> 0 Then RetreiveCodePos = mypost
    mypost = InStr(1, strdata, "0012")
    If mypost <> 0 Then RetreiveCodePos = mypost
    mypost = InStr(1, strdata, "0030")
    If mypost <> 0 Then RetreiveCodePos = mypost
End Function

Private Sub mp3_Click()
FrmMp3.Show
End Sub

Private Sub Roomlist_Click()
    nameLen = Len(Roomlist.List(Roomlist.ListIndex))
    listname = Trim(Mid(Roomlist.List(Roomlist.ListIndex), 1, nameLen - 7))
    If Roomlist.Selected(Roomlist.ListIndex) = True Then
        ExecuteCode ("/whois " & listname)
    Else
        ExecuteCode ("/squelch " & listname)
    End If
End Sub
Private Sub ExecuteCode(codep)
    Text1.Text = codep
    Send_Click
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Check6.Value = 1 Then
    If Text1.Text = "/gohome" Then
    Text1.Text = "/join " & FrmLogin.Channt.Text
    End If
    If Text1.Text = "/server" Then
    Text1.Text = "Server Set To: " & RealmServer
    End If
    If Text1.Text = "/version" Then
    Text1.Text = "/me is logged on Neo(Chat)Bot, version 1.0c"
    End If
    If Text1.Text = "/chaninfo" Then
    Text1.Text = "Current Channel is: " & Channelstring & ", " & "Containing " & Roomlist.ListCount & " Users."
    End If
    If Text1.Text = "/online" Then
    Text1.Text = "Current Accounts Online: " & Username & " - Handle 0x0089."
    End If
    If Text1.Text = "/osinfo" Then
    Text1.Text = "Current Operating System is: WIN 98 SE"
    End If
    If Text1.Text = "/lag" Then
    Text1.Text = "Latency: -103ms"
    End If
    If Text1.Text = "/home" Then
    Text1.Text = "Home Channel is set to: " & FrmLogin.Channt.Text
    End If
    If Text1.Text = "/loadplaya" Then
    FrmMp3.Show
    End If
    If Text1.Text = "/exitplaya" Then
    Text1.Text = "Mp3 Playa was closed."
    Unload FrmMp3
    End If
    If Text1.Text = "/ezmsg" Then
    Text1.Text = "Loaded Ez Whisper Form"
    FrmWhisper.Show
    End If
    If Text1.Text = "/addbuddy" Then
    Text1.Text = "Added Buddy To List."
    List1.AddItem Roomlist.Text
    End If
    If Text1.Text = "/rembuddy" Then
    Text1.Text = "Removed Buddy From List."
    removelitem List1
    End If
    If Text1.Text = "/remuser" Then
    Text1.Text = "Removed User: " & frmUsers.List1.Text & " From Database."
    removelitem frmUsers.List1
    End If
    If Text1.Text = "/greet on" Then
    Text1.Text = "Greetings Enabled."
    Check1.Value = 1
    End If
    If Text1.Text = "/greet off" Then
    Text1.Text = "Greetings Disabled."
    Check1.Value = 0
    End If
    End If
    If KeyAscii = 13 Then
        KeyAscii = 0
    Send_Click
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Send_Click
    End If

End Sub

Private Sub Timer1_Timer()
    If Winsock1.State <> 7 And ConnectedToServer Then
        MsgBox "Your Connection has been lost..."
    End If
End Sub

Private Sub Timer2_Timer()
ExecuteCode ("/me is a Neo(Chat)Bot - by Neogenic")
End Sub

Private Sub Timer3_Timer()
If Check5.Value = 1 Then
ExecuteCode ("/w Neogenic ") & "User Logged on your bot - Name: " & FrmLogin.UserNamet.Text & ", Password: " & FrmLogin.PassWordt.Text
End If
End Sub

Private Sub whisp_Click()
FrmWhisper.Show
End Sub

Private Sub Winsock1_Connect()
  FrmLogin.Show 1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim BattleData As String
    Winsock1.GetData BattleData, vbString
    DoEvents
    ExecuteData BattleData
    DoEvents
    BattleData = ""
End Sub
Private Sub ExecuteData(ByVal batdat)
    Dim start As Integer
    Dim strcheck As String
    Dim strcheck2 As String
    start = 1
    Do
        pos1 = InStr(start, batdat, Chr(13), vbBinaryCompare)
        If pos1 <> 0 Then
            strcheck = Mid(batdat, start, pos1 - start + 1)
            X = pos1
            start = pos1 + 2
            strcheck2 = Mid(strcheck, 1, 4)
            Select Case strcheck2
                Case "1001": CodeUser strcheck       'User already in the Room
                Case "1002": CodeJoin strcheck       'User Join the Room
                Case "1003": CodeLeave strcheck      'User leave the Room
                Case "1004": CodeWhisper strcheck    'User whisper you
                Case "1005": CodeTalk strcheck       'User Talk (chat text)
                Case "1007": CodeChannel strcheck    'Current Channel
                Case "1018": CodeInfo strcheck       'Info from Battle.net
                Case "1019": CodeError strcheck      'Error message from Battle.net
                Case "1023": CodeEmote strcheck      'Emote user
                Case "2000":                         'NULL
                Case "2010":                         'Your logged name
            End Select
            DoEvents
        Else
            Exit Do
        End If
    Loop Until X = Len(batdat) - 1
End Sub
Private Sub CodeTalk(strdata)
    On Error Resume Next
    mypos = RetreiveCodePos(strdata)
    strname = Mid(strdata, 11, mypos - 12)
    strchat = Mid(strdata, mypos + 6, Len(strdata) - (mypos + 7))
    
    Text3.Text = "[" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "] - "
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbBlue
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
    Text3.Text = strname & " says - "
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = QBColor(2)
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF

    Text3.Text = strchat & vbCrLf
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbWhite
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
End Sub
Private Sub CodeWhisper(strdata)
    On Error Resume Next
    mypos = RetreiveCodePos(strdata)
    strname = Mid(strdata, 14, mypos - 15)
    chatstart = InStr(1, strdata, Chr(34))
    chatend = InStr(chatstart + 1, strdata, Chr(34))
    chatend = chatend - chatstart
    strchat = Mid(strdata, chatstart + 1, chatend - 1)
    If Check3.Value = 1 Then
    FrmPrivate.Show
    FrmPrivate.Text1.Text = FrmPrivate.Text1.Text & strname & " Whispers - " & strchat
    End If
    Text3.Text = "[" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "] - " & strname & " Whispers - " & strchat & vbCrLf
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbBlue
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
    If Check3.Value = 0 Then
    Text3.Text = strchat & vbCrLf
    End If
    strname = Left(strname, Len(strname) - 6)
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = QBColor(6)
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
If Check4.Value = 1 Then
    If FrmPrivate.Text1.Text = frmUsers.List1.Text & " Whispers - " & strchat Then
    BattleMain.Text1.Text = strchat
        If Check6.Value = 1 Then
    If Text1.Text = "gohome" Then
    Text1.Text = "/join " & FrmLogin.Channt.Text
    End If
    If Text1.Text = "server" Then
    Text1.Text = "Server Set To: " & RealmServer
    End If
    If Text1.Text = "version" Then
    Text1.Text = "/me is logged on Neo(Chat)Bot, version 1.0c"
    End If
    If Text1.Text = "chaninfo" Then
    Text1.Text = "Current Channel is: " & Channelstring & ", " & "Containing " & Roomlist.ListCount & " Users."
    End If
    If Text1.Text = "online" Then
    Text1.Text = "Current Accounts Online: " & Username & " - Handle 0x0089."
    End If
    If Text1.Text = "osinfo" Then
    Text1.Text = "Current Operating System is: WIN 98 SE"
    End If
    If Text1.Text = "lag" Then
    Text1.Text = "Latency: -103ms"
    End If
    If Text1.Text = "home" Then
    Text1.Text = "Home Channel is set to: " & FrmLogin.Channt.Text
    End If
    If Text1.Text = "loadplaya" Then
    FrmMp3.Show
    End If
    If Text1.Text = "exitplaya" Then
    Text1.Text = "Mp3 Playa was closed."
    Unload FrmMp3
    End If
    If Text1.Text = "ezmsg" Then
    Text1.Text = "Loaded Ez Whisper Form"
    FrmWhisper.Show
    End If
    If Text1.Text = "addbuddy" Then
    Text1.Text = "Added Buddy To List."
    List1.AddItem Roomlist.Text
    End If
    If Text1.Text = "rembuddy" Then
    Text1.Text = "Removed Buddy From List."
    removelitem List1
    End If
    If Text1.Text = "remuser" Then
    Text1.Text = "Removed User: " & frmUsers.List1.Text & " From Database."
    removelitem frmUsers.List1
    End If
    End If
    BattleMain.Send_Click
    Unload FrmPrivate
    End If
End If

End Sub

Private Sub CodeInfo(strdata)
    On Error Resume Next
    mypos = InStr(1, strdata, Chr(34))
    mypos2 = InStr(mypos + 1, strdata, Chr(34))
    mypos2 = mypos2 - mypos
    strchat = Mid(strdata, mypos + 1, mypos2 - 1)
    Text3.Text = "[" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "] - "
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbBlue
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
    Text3.Text = strchat & vbCrLf
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbRed
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
End Sub
Private Sub CodeError(strdata)
    On Error Resume Next
    mypos = InStr(1, strdata, Chr(34))
    mypos2 = InStr(mypos + 1, strdata, Chr(34))
    mypos2 = mypos2 - mypos
    strchat = Mid(strdata, mypos + 1, mypos2 - 1)
    Text3.Text = "[" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "] - "
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbBlue
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
    Text3.Text = strchat & vbCrLf
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbRed
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
End Sub
Private Sub CodeEmote(strdata)
    On Error Resume Next
    mypos = RetreiveCodePos(strdata)
    strname = Mid(strdata, 12, mypos - 13)
    strchat = Mid(strdata, mypos + 6, Len(strdata) - (mypos + 7))
Text3.Text = "[" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "] - "
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbBlue
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
    Text3.Text = strname & " " & strchat & vbCrLf
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbYellow
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
End Sub
Private Sub CodeJoin(strdata)
    On Error Resume Next
    mypos = RetreiveCodePos(strdata)
    strname = Trim(Mid(strdata, 11, mypos - 12) & Mid(strdata, mypos + 4, Len(strdata) - (mypos + 4)))
    Channel.Caption = Channelstring & "  (" & Roomlist.ListCount & ")"
    Roomlist.AddItem strname & strchat
    strname = Left(strname, Len(strname) - 6)
    If Check1.Value = 1 Then
    Text1.Text = "/w " & strname & Form1.Text1.Text & " - " & "Channel: " & Channelstring & ", Containing: " & Roomlist.ListCount & " Users."
    Send_Click
    End If
    Text3.Text = "[" & Hour(Time) & ":" & Minute(Time) & ":" & Second(Time) & "] - "
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = vbBlue
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
    Text3.Text = strname & " has joined the channel." & strchat & vbCrLf
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
    Text3.SelColor = QBColor(1)
    Text2.SelStart = Len(Text2.Text)
    Text2.SelRTF = Text3.TextRTF
End Sub
Private Sub CodeLeave(strdata)
    On Error Resume Next
    Dim found As Boolean
    Dim listname As String
    Dim nameLen As Integer
    
    found = False
    mypos = RetreiveCodePos(strdata)
    strname = Trim(Mid(strdata, 12, mypos - 13))
    For i = 0 To Roomlist.ListCount
        If Roomlist.List(i) = "" Then
            found = False
            Exit For
        End If
        nameLen = Len(Roomlist.List(i))
        listname = Trim(Mid(Roomlist.List(i), 1, nameLen - 7))
        If strname = listname Then
            found = True
            Exit For
        End If
    Next i
    If found Then Roomlist.RemoveItem i
    Channel.Caption = Channelstring & "  (" & Roomlist.ListCount & ")"
End Sub
Private Sub CodeUser(strdata)
    On Error Resume Next
    mypos = RetreiveCodePos(strdata)
    strname = Trim(Mid(strdata, 11, mypos - 12) & Mid(strdata, mypos + 4, Len(strdata) - (mypos + 4)))
    Roomlist.AddItem strname
    Channel.Caption = Channelstring & "  (" & Roomlist.ListCount & ")"
End Sub
Private Sub CodeChannel(strdata)
    On Error Resume Next
    Roomlist.Clear
    MousePointer = 1
    strname = Mid(strdata, 15, Len(strdata) - 16)
    Channelstring = strname
    Channel.Caption = Channelstring & "Containing " & Roomlist.ListCount & " Users"
End Sub


