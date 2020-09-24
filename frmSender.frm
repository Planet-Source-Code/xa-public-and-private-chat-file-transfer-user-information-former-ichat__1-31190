VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSender 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transfer"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   Icon            =   "frmSender.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About File Transfer"
      Height          =   330
      Left            =   2835
      TabIndex        =   6
      Top             =   1440
      Width           =   2070
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   330
      Left            =   2835
      TabIndex        =   5
      Top             =   1800
      Width           =   2070
   End
   Begin VB.Timer tmrState 
      Interval        =   200
      Left            =   4500
      Top             =   630
   End
   Begin VB.CommandButton cmdReciever 
      Caption         =   "Enable File Reciever"
      Enabled         =   0   'False
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   1800
      Width           =   2640
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Send a file over the net"
      Height          =   330
      Left            =   105
      TabIndex        =   0
      Top             =   1440
      Width           =   2640
   End
   Begin MSWinsockLib.Winsock BNS 
      Left            =   4455
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   50
   End
   Begin MSWinsockLib.Winsock WNS 
      Left            =   4455
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   50
   End
   Begin VB.Label BState 
      AutoSize        =   -1  'True
      Caption         =   "Preparing Services"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Status 
      Caption         =   "Welcome To File Transfer"
      Height          =   225
      Left            =   225
      TabIndex        =   3
      Top             =   825
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   $"frmSender.frx":0ECA
      Height          =   975
      Left            =   180
      TabIndex        =   1
      Top             =   75
      Width           =   4710
   End
End
Attribute VB_Name = "frmSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private File As String
Private Reset As Boolean


'XP UPGRADE AUTO-INSERT
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub BNS_Close()
On Error Resume Next
BNS.Close
End Sub

Private Sub BNS_DataArrival(ByVal bytesTotal As Long)
Dim leng As String, x As String
leng = FileLen(File)
Open File For Binary Access Read As #1
If leng > 102400 Then
'LARGE FILE > 10 KB
Dim y As String, c As Long, iB As Long
Do Until c = leng
If c + 102400 <= leng Then
'Single 124000 to take
y = Input(102400, 1)
c = c + 102400
Else
y = Input(leng - c, 1)
c = leng
End If
x = x & y
DoEvents
BState = "Opening file " & c & " bytes out of " & leng
DoEvents
Loop
Else
' SMALL FILE - Input everything
x = Input(LOF(1), 1)
End If
' Done with input
BNS.SendData x
Close #1
End Sub

Private Sub BNS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
BNS.Close
End Sub

Private Sub cmdAbout_Click()
frmAbout.ShowAbout (False)
frmAbout.Show vbModal
End Sub

Public Sub cmdConnect_Click()
Dim host As String
host = InputBox("Insert destination IP:")
If host = "" Then Exit Sub
File = InputBox("Insert file to transfer:")
If File = "" Then Exit Sub
BNS.Close
BNS.RemoteHost = host
BNS.Connect
wait:
DoEvents
If BNS.State <> 7 Then GoTo wait
On Error GoTo error
BNS.SendData FileLen(File)
Exit Sub
error:
MsgBox Err.Description
BNS.Close
End Sub

Private Sub cmdExit_Click()
Me.Hide
End Sub

Private Sub cmdReciever_Click()
Select Case cmdReciever.Caption
Case "Enable File Reciever"
WNS.Listen
Case "Disable File Reciever"
WNS.Close
End Select
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
App.Title = "File Transfer by HISoft"
End Sub

Private Sub tmrState_Timer()
Select Case WNS.State
Case 2
Status = "Reciever Listening"
Status.ForeColor = RGB(55, 175, 55)
If cmdReciever.Enabled = False Then cmdReciever.Enabled = True
If LCase(Left(cmdReciever.Caption, 6)) = "disabl" Then GoTo doagain
cmdReciever.Caption = "Disable File Reciever"
Case 6 Or 7
Status = "Recieving Data"
cmdReciever.Enabled = False
Status.ForeColor = vbRed
Case 0
Status = "Reciever Disabled"
If cmdReciever.Enabled = False Then cmdReciever.Enabled = True
Status.ForeColor = vbBlue

If LCase(Left(cmdReciever.Caption, 6)) = "enable" Then GoTo doagain
cmdReciever.Caption = "Enable File Reciever"
End Select
doagain:
Select Case BNS.State
Case 0
BState = ""
cmdConnect.Enabled = True
Case sckConnecting
BState = "Connecting..."
cmdConnect.Enabled = False
Case 7
cmdConnect.Enabled = False
End Select
End Sub

Private Sub WNS_ConnectionRequest(ByVal requestID As Long)
WNS.Close
WNS.Accept requestID
Reset = True
End Sub

Private Sub WNS_DataArrival(ByVal bytesTotal As Long)
Static BC As Long
Static Second As Boolean
Dim x As String, xC As Long
If Reset Then BC = 0: Second = False: x = "": Reset = False
If Not Second Then
Call WNS.GetData(xC, vbLong, bytesTotal)
If xC = 0 Then WNS.Close: Exit Sub
Second = True
BC = xC
Dim res As VbMsgBoxResult
Me.Show
res = MsgBox("Accept incoming file (" & BC & " bytes in length; from " & WNS.RemoteHostIP & ")?", vbQuestion + vbYesNo)
If res = vbNo Then WNS.Close: Exit Sub
WNS.SendData "ok"
Else
If bytesTotal <> BC Then Exit Sub
x = InputBox("Insert path where to save the file:")
Open x For Binary Access Write As #1
Dim y As String
Call WNS.GetData(y, vbString, bytesTotal)
Put #1, , y
Close #1
WNS.Close
WNS.Listen
MsgBox "Done"
End If
End Sub

Private Sub WNS_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
WNS.Close
WNS.Listen
End Sub

Private Sub Form_Initialize()
'XP UPGRADE AUTO-INSERT
InitCommonControls
End Sub
