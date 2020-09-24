VERSION 5.00
Begin VB.Form frmCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect To..."
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6015
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4680
      TabIndex        =   5
      Top             =   555
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   315
      Left            =   4665
      TabIndex        =   4
      Top             =   195
      Width           =   1245
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Private - cannot see in the user list"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Using this, you are not listed in the user list, and cannot be sent private messanges."
      Top             =   1290
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Clear - can see you in the user list"
      Height          =   240
      Left            =   1080
      TabIndex        =   2
      ToolTipText     =   "Using this, everyone can see you in the chat. You can recieve private messages."
      Top             =   1080
      Value           =   -1  'True
      Width           =   3180
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1215
      TabIndex        =   1
      Top             =   705
      Width           =   2025
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   225
      Width           =   2040
   End
   Begin VB.Label Label2 
      Caption         =   "My Nickname:"
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   780
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Remote Host:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   270
      Width           =   1035
   End
End
Attribute VB_Name = "frmCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Cancel As Boolean


'XP UPGRADE AUTO-INSERT
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub cmdCancel_Click()
Cancel = True
Me.Hide
End Sub

Private Sub cmdOK_Click()
If txtHost <> "" And txtNick <> "" Then Me.Hide
End Sub

Public Sub GetInfo(HostName As String, UserName As String)
If HostName <> "" Then txtHost.Text = HostName: txtHost.Enabled = False
If UserName <> "" Then txtNick.Text = UserName: txtNick.Enabled = False
Me.Show vbModal
If Cancel Then HostName = "": UserName = "": Cancel = False: GoTo quit
HostName = txtHost
UserName = IIf(Option1.Value = True, txtNick, vbNullChar & txtNick)
quit:
Unload Me
End Sub

Private Sub txtNick_KeyPress(KeyAscii As Integer)
If KeyAscii = Asc(" ") Then KeyAscii = 0
End Sub

Private Sub Form_Initialize()
'XP UPGRADE AUTO-INSERT
InitCommonControls
End Sub
