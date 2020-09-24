VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IChat Professional"
   ClientHeight    =   6435
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9120
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrStatus 
      Interval        =   200
      Left            =   6120
      Top             =   15
   End
   Begin VB.ListBox lstUsers 
      Height          =   5325
      Left            =   6720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   2280
   End
   Begin MSWinsockLib.Winsock MyC 
      Left            =   7440
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   100
      LocalPort       =   8
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   270
      Left            =   5805
      TabIndex        =   5
      Top             =   6060
      Width           =   810
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7905
      Top             =   50
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtMsgSend 
      Height          =   285
      Left            =   930
      TabIndex        =   4
      Top             =   6060
      Width           =   4830
   End
   Begin RichTextLib.RichTextBox Messages 
      Height          =   5580
      Left            =   210
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   450
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   9843
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmServer.frx":0ECA
   End
   Begin MSWinsockLib.Winsock Connector 
      Index           =   0
      Left            =   8520
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   100
   End
   Begin VB.Label lblState 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3975
      TabIndex        =   9
      Top             =   240
      Width           =   45
   End
   Begin VB.Label Label4 
      Caption         =   "Your Status:"
      Height          =   195
      Left            =   2955
      TabIndex        =   8
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "Users:"
      Height          =   300
      Left            =   6735
      TabIndex        =   7
      Top             =   510
      Width           =   2250
   End
   Begin VB.Label Label2 
      Caption         =   "Send:"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Server Offline"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1155
      TabIndex        =   1
      Top             =   240
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "IChat Status:"
      Height          =   210
      Left            =   225
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnustart 
         Caption         =   "&Start Server"
      End
      Begin VB.Menu mnustop 
         Caption         =   "S&top Server"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucon 
         Caption         =   "&Connect To Server"
      End
      Begin VB.Menu mnudisc 
         Caption         =   "&Disconnect"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSendF 
         Caption         =   "Send a file or a photo..."
      End
      Begin VB.Menu mnusep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuquit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu usrMenu 
      Caption         =   "User Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuprivate 
         Caption         =   "&Private Message"
      End
      Begin VB.Menu mnuSendFile 
         Caption         =   "&Send a file  or a photo..."
      End
      Begin VB.Menu mnusep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnukick 
         Caption         =   "&Kick"
      End
      Begin VB.Menu mnukickbyID 
         Caption         =   "Kick by &ID"
      End
      Begin VB.Menu mnusep17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuuseInfo 
         Caption         =   "&User Information"
      End
      Begin VB.Menu mnuUIDID 
         Caption         =   "User Information by I&D"
      End
   End
   Begin VB.Menu mnuchat 
      Caption         =   "&Chat Menu"
      Visible         =   0   'False
      Begin VB.Menu mnusavechat 
         Caption         =   "&Save chat as..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuabout 
         Caption         =   "About IChat Professional"
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IPLookUp As String
Private FreePl() As Boolean
Public totalPlaces As Long
Private Nick() As String

'XP UPGRADE AUTO-INSERT
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub cmdSend_Click()
MyC.SendData txtMsgSend.Text
txtMsgSend = ""
txtMsgSend.SetFocus
End Sub

Private Sub Connector_Close(Index As Integer)
Cone(Index).active = False
End Sub

Private Sub Connector_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Select Case Index
Case 0
Load Connector(totalPlaces + 1)
totalPlaces = totalPlaces + 1
Connector(totalPlaces).LocalPort = 100 + totalPlaces
Connector(totalPlaces).Accept requestID
ReDim Preserve Cone(totalPlaces)
Cone(totalPlaces).IP = Connector(totalPlaces).RemoteHostIP
Cone(totalPlaces).active = True
wait:
DoEvents
If Connector(totalPlaces).State <> 7 Then Exit Sub
End Select
End Sub

Private Sub Connector_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim data As String
Call Connector(Index).GetData(data, vbString)
DoEvents

If Cone(Index).Name = "" Then
    'First time - add name
    Cone(Index).Name = data
    SendNewList 'Send the new user-list
    Exit Sub
End If

If Left(data, 1) = Chr(1) Then
    'SYSTEM MESSAGE
    D = Len(Messages.Text)
    Messages.Text = Messages.Text & Right(data, Len(data) - 1) & vbNewLine
    Messages.SelStart = D
    Messages.SelLength = Len(data) - 1
    Messages.SelBold = True
    Messages.SelColor = RGB(0, 0, 255)
    Messages.SelStart = Len(Messages.Text)
    For i = 1 To totalPlaces
        If Cone(i).active Then Connector(i).SendData data
    Next i
    Exit Sub
End If

If Left(data, 2) = "/P" Then
    Dim dest As String
    'PRIVATE MESSAGE
    For i = 4 To Len(data)
        If Mid(data, i, 1) = " " Then dest = Mid(data, 4, i - 4): Exit For
    Next i
    If dest = "" Then Exit Sub
    Connector(FindID(dest)).SendData "From " & Cone(Index).Name & ": " & Right(data, Len(data) - 4 - Len(dest))
    Exit Sub
End If

If Left(data, 2) = "/I" Then
    Connector(Index).SendData "/I " & Cone(FindID(Right(data, Len(data) - 3))).IP
    Exit Sub
End If

'Simple Message
data = Cone(Index).Name & ": " & data
' Send Messages to all network connected users
For i = 1 To totalPlaces
If Cone(i).active Then
Connector(i).SendData data
End If
DoEvents
Next i
End Sub

Private Sub Connector_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
MsgBox "Server Winsock Error: " & Description & " (" & Number & ")"
End Sub
Public Sub SendNewList()
For i = 1 To totalPlaces
If Cone(i).active Then v = v & Cone(i).Name & "|"
Next i
v = Chr(2) & v
On Error Resume Next
For i = 1 To totalPlaces
If Cone(i).active Then Connector(i).SendData v
Next i
End Sub

Private Sub Form_Load()
Load frmSender
On Error GoTo inuse
frmSender.WNS.Listen
Exit Sub
inuse:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub lstUsers_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Then Exit Sub
If ListIndex = -1 Then Exit Sub
PopupMenu usrMenu
End Sub

Private Sub Messages_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 2 Then Exit Sub
PopupMenu mnuchat
End Sub

Private Sub mnuabout_Click()
frmAbout.ShowAbout (True)
frmAbout.Show vbModal
End Sub

Private Sub mnucon_Click()
Dim servIP As String, uN As String
Call frmCon.GetInfo(servIP, uN)
If servIP = "" Then Exit Sub
MyC.Close
MyC.RemoteHost = servIP
redo:
On Error Resume Next
MyC.Connect
If MyC.State = 0 Then GoTo inuse
On Error GoTo 0
wait:
DoEvents
If MyC.State <> 7 Then GoTo wait
MyC.SendData uN
Exit Sub
inuse:
MyC.Close
MyC.LocalPort = MyC.LocalPort + 1
On Error GoTo 0
GoTo redo
End Sub

Private Sub mnudisc_Click()
MyC.Close
lstUsers.Clear
End Sub

Public Sub mnukick_Click()
On Error GoTo youarenotahost
Dim x As String, y As Long
If MyC.RemoteHost <> "localhost" Then GoTo youarenotahost
If lstUsers.ListIndex = -1 Then Exit Sub
x = lstUsers.List(lstUsers.ListIndex)
y = FindID(x, False)
If y = 0 Then GoTo youarenotahost
Cone(y).active = False
Connector(y).Close
SendNewList
Exit Sub
youarenotahost:
MsgBox "You are not the host of this chat and therefore are not qualified to kick somebody." & vbNewLine & "Or, the person you tried to kick was not found in the list"
End Sub

Private Sub mnukickbyID_Click()
On Error GoTo error
If MyC.RemoteHost <> "localhost" Then GoTo error
x = InputBox("Insert userID to kick:")
If IsNumeric(x) = False Then Exit Sub
If x = "" Then Exit Sub
Cone(x).active = False
Connector(x).Close
SendNewList 'Update User List
Exit Sub
error:
MsgBox "This option is not availble because you are not the admin of this chat"
End Sub

Private Sub mnuprivate_Click()
If txtMsgSend = "" Then MsgBox "Type the message in first": Exit Sub
txtMsgSend = "/P " & lstUsers.List(lstUsers.ListIndex) & " " & txtMsgSend
cmdSend_Click
End Sub

Private Sub mnuquit_Click()
If Connector(0).State <> sckListening Then GoTo ex
If totalPlaces = 1 Then GoTo ex
CDW.Show vbModal
ex:
Messages.Text = Messages.Text & "Shutdown of chat servers in progress..." & vbNewLine
On Error Resume Next
For i = 0 To totalPlaces
Connector(i).Close
Unload Connector(i)
DoEvents
Next i
End
End Sub

Private Sub mnusavechat_Click()
CD.Filter = "RTF Files|*.rtf"
CD.ShowSave
If CD.FileName = "" Then Exit Sub
Messages.SaveFile CD.FileName
CD.FileName = ""
End Sub

Private Sub mnusendF_Click()
frmSender.Show
End Sub

Private Sub mnuSendFile_Click()
If lstUsers.ListIndex = -1 Then Exit Sub
MyC.SendData "/P " & lstUsers.List(lstUsers.ListIndex) & " I would be attempting to send you a file soon (automatic message)"
DoEvents
For i = 1 To 1000
DoEvents
Next i
IPLookUp = ""
MyC.SendData "/I " & lstUsers.List(lstUsers.ListIndex)
wait:
DoEvents
If IPLookUp = "" Then GoTo wait
frmSender.Show
Call InputBox("Please copy this IP address (you will need to paste it in the next box):", , IPLookUp)
IPLookUp = ""
frmSender.cmdConnect_Click
End Sub

Private Sub mnustart_Click()
lstUsers.Clear
Dim Ho As String
Dim nickN As String
Ho = "localhost"
Call frmCon.GetInfo(Ho, nickN)
If Ho = "" Then Exit Sub
'Destory Previous server Components
ReDim Cone(0)
totalPlaces = 0
Messages.Text = ""
'Invisible
'nickN = vbNullChar & nickN
Connector(0).Listen
MyC.Close
MyC.RemoteHost = Ho
wait1:
DoEvents
If Connector(0).State <> sckListening Then GoTo wait1
MyC.Connect
wait:
DoEvents
If MyC.State <> 7 Then GoTo wait
MyC.SendData nickN
End Sub

Private Sub mnustop_Click()
If totalPlaces > 1 Then
CDW.Show vbModal
GoTo down
Else
down:
Connector(0).Close
On Error Resume Next
For i = 1 To totalPlaces
Connector(i).Close
DoEvents
Unload Connector(i)
Next i
End If
End Sub

Private Sub mnuUIDID_Click()
If MyC.RemoteHost <> "localhost" Then GoTo adminerror
id = InputBox("Insert user ID:")
If id = "" Then Exit Sub
If IsNumeric(id) = False Then Exit Sub
frmUInfo.ShowUI id
frmUInfo.Show
Exit Sub
adminerror:
MsgBox "This option is not avalible because you are not a host of this chat"
End Sub

Private Sub mnuuseInfo_Click()
On Error GoTo uerror
If MyC.RemoteHost <> "localhost" Then GoTo net
Dim x As String, y As Long
x = lstUsers.List(lstUsers.ListIndex)
y = FindID(x, False)
frmUInfo.ShowUI y
frmUInfo.Show
Exit Sub
net:
x = lstUsers.List(lstUsers.ListIndex)
IPLookUp = ""
MyC.SendData "/I " & x
wait:
DoEvents
If IPLookUp = "" Then GoTo wait
Call frmNUInfo.ShowUI(x, IPLookUp)
IPLookUp = ""
frmNUInfo.Show
End Sub

Private Sub MyC_Close()
c = Len(Messages.Text)
Messages.Text = Messages.Text & "You have been disconnected from this chat" & vbNewLine
Messages.SelStart = c
Messages.SelLength = Len("You have been disconnected from this chat")
Messages.SelBold = True
mnudisc_Click
End Sub

Private Sub MyC_DataArrival(ByVal bytesTotal As Long)
Dim data As String
Dim x As Long
Dim y As String
Call MyC.GetData(data, vbString)
DoEvents
If Left(data, 1) <> Chr(1) And Left(data, 1) <> Chr(2) Then
    If Left(data, 2) <> "/I" Then
    'SIMPLE MESSAGE
    Messages.Text = Messages.Text & Replace(data, Chr(1), "") & vbNewLine
    Else
    'REPLY TO IP-Challange
    IPLookUp = Right(data, Len(data) - 3)
    End If
ElseIf Left(data, 1) = Chr(2) Then
    'LIST-REFRESH
    lstUsers.Clear
    data = Right(data, Len(data) - 1)
    x = InStr(data, "|")
    Do Until x = 0
        y = Left(data, x - 1)
        If Left(y, 1) = Chr(1) Then GoTo done
        lstUsers.AddItem y
done:
        data = Right(data, Len(data) - x)
        x = InStr(data, "|")
    Loop
Else
'SYSTEM MESSAGE
D = Len(Messages.Text)
Messages.Text = Messages.Text & Right(data, Len(data) - 1) & vbNewLine
Messages.SelStart = D
Messages.SelLength = Len(data) - 1
Messages.SelBold = True
Messages.SelColor = RGB(0, 0, 255)
Messages.SelStart = Len(Messages.Text)
End If
End Sub

Private Sub MyC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Static ErrorCount
Select Case Number
Case sckAddressInUse
'In Use or Not availible
MyC.Close
MyC.LocalPort = MyC.LocalPort + 1
ErrorCount = ErrorCount + 1
If ErrorCount = 30 Then MsgBox "Connection aborted. Too many Address In Use errors.": Exit Sub
On Error Resume Next
MyC.Connect
Case Else
MsgBox "Local Winsock Error: " & Description & " (" & Number & ")"
End Select
End Sub

Private Sub tmrStatus_Timer()
Dim s As Label, v As Label
Set s = lblStatus
Set v = lblState
Select Case Connector(0).State
Case sckClosed
s = "Server Offline"
mnustart.Enabled = True
mnustop.Enabled = False
Case sckListening
s = "Server Online"
mnustart.Enabled = False
mnustop.Enabled = True
End Select
Select Case MyC.State
Case sckClosed
v = "Not Connected to anywhere"
mnucon.Enabled = True
mnudisc.Enabled = False
Case sckConnected
v = "Connected to " & MyC.RemoteHostIP
mnudisc.Enabled = True
mnucon.Enabled = False
Case sckConnectionPending
v = "Connecting to " & MyC.RemoteHostIP
mnudisc.Enabled = True
mnucon.Enabled = False
End Select
End Sub

Private Sub txtMsgSend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdSend_Click
End Sub

Private Sub Form_Initialize()
'XP UPGRADE AUTO-INSERT
InitCommonControls
End Sub
