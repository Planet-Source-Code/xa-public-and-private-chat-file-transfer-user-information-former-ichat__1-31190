VERSION 5.00
Begin VB.Form CDW 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please wait."
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCount 
      Interval        =   1000
      Left            =   4185
      Top             =   825
   End
   Begin VB.CommandButton cmdShut 
      Caption         =   "Shut Down"
      Enabled         =   0   'False
      Height          =   600
      Left            =   1830
      TabIndex        =   2
      Top             =   795
      Width           =   2310
   End
   Begin VB.Label lblTimeLeft 
      Caption         =   "60"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1065
      TabIndex        =   1
      Top             =   825
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   $"CDW.frx":0000
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4500
   End
End
Attribute VB_Name = "CDW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'XP UPGRADE AUTO-INSERT
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub Command1_Click()

End Sub

Private Sub cmdShut_Click()
Unload Me
End Sub

Private Sub Form_Load()
frmServer.MyC.SendData Chr(1) & "This server is beeing shutdown. For 60 seconds, all connections are safe. After 60 secs the host has the option to manually turn off this server. This means that you will lose your connection. Please chose a next host to start the server now, and join his server."
End Sub

Private Sub tmrCount_Timer()
lblTimeLeft = lblTimeLeft - 1
If lblTimeLeft = 0 Then
cmdShut.Enabled = True
tmrCount.Enabled = False
End If
End Sub

Private Sub Form_Initialize()
'XP UPGRADE AUTO-INSERT
InitCommonControls
End Sub
