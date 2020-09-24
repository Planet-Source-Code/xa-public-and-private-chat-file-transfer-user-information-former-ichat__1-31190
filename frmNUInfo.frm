VERSION 5.00
Begin VB.Form frmNUInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Information"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmNUInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPing 
      Caption         =   "Ping User"
      Height          =   240
      Left            =   225
      TabIndex        =   11
      Top             =   2175
      Width           =   1500
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   405
      Left            =   2790
      TabIndex        =   7
      Top             =   2385
      Width           =   1845
   End
   Begin VB.CommandButton cmdTrace 
      Caption         =   "Launch"
      Height          =   240
      Left            =   1770
      TabIndex        =   5
      Top             =   1620
      Width           =   750
   End
   Begin VB.Label lblPort 
      Caption         =   "Is not disclosed. (ICHAT SEVER)"
      Height          =   195
      Left            =   1875
      TabIndex        =   10
      Top             =   1365
      Width           =   2535
   End
   Begin VB.Label lblIP 
      Height          =   180
      Left            =   1875
      TabIndex        =   9
      Top             =   1020
      Width           =   2070
   End
   Begin VB.Label lblNick 
      Height          =   210
      Left            =   1875
      TabIndex        =   8
      Top             =   705
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "User Options:"
      Height          =   195
      Left            =   195
      TabIndex        =   6
      Top             =   1950
      Width           =   4050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "TraceRoute Query:"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1635
      Width           =   1365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Port"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   1330
      Width           =   285
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "IP:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   1025
      Width           =   195
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Chat Nickname:"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   675
      TabIndex        =   0
      Top             =   90
      Width           =   3540
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmNUInfo.frx":08CA
      Top             =   105
      Width           =   480
   End
End
Attribute VB_Name = "frmNUInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private UID As Long
'XP UPGRADE AUTO-INSERT
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Public Sub ShowUI(ByVal Name As String, IP As String)
lblNick = Name
lblIP = IP
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdKick_Click()
Dim res As VbMsgBoxResult
res = MsgBox("Are you sure you want to kick " & lblNick & "?", vbQuestion + vbYesNo)
If res = vbNo Then Exit Sub
frmServer.lstUsers.ListIndex = UID - 1
frmServer.mnukick_Click
Unload Me
End Sub

Private Sub cmdPing_Click()
Call Shell("ping " & frmServer.Connector(UID).RemoteHostIP, vbNormalFocus)
End Sub

Private Sub cmdTrace_Click()
Call Shell("tracert " & lblIP, vbNormalFocus)
End Sub

Private Sub Label8_Click()

End Sub

Private Sub Form_Initialize()
'XP UPGRADE AUTO-INSERT
InitCommonControls
End Sub

Private Sub lblNick_Click()

End Sub
