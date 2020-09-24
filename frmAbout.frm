VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About File Transfer"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmAIC 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2070
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   4530
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Copyright (c) 2001-2002 HISoft (Andrei O. Lisovoi)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1695
         Width           =   4290
      End
      Begin VB.Label Label6 
         Caption         =   "This software utilizes parts of File Transfer by HISoft."
         Height          =   240
         Left            =   150
         TabIndex        =   9
         Top             =   1320
         Width           =   4020
      End
      Begin VB.Label Label5 
         Caption         =   "IChat Professional allows you to host a chat server, or join one. You could send files over, in addition to simple chat."
         Height          =   525
         Left            =   135
         TabIndex        =   8
         Top             =   885
         Width           =   4140
      End
      Begin VB.Label lblIV 
         Caption         =   "Version: "
         Height          =   255
         Left            =   915
         TabIndex        =   7
         Top             =   660
         Width           =   3240
      End
      Begin VB.Label Label4 
         Caption         =   "IChat Professional"
         BeginProperty Font 
            Name            =   "Westminster"
            Size            =   24
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   900
         TabIndex        =   6
         Top             =   195
         Width           =   3585
      End
      Begin VB.Image imgLOGO 
         Height          =   720
         Left            =   120
         Picture         =   "frmAbout.frx":0ECA
         Top             =   105
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   315
      Left            =   2850
      TabIndex        =   4
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   225
      Picture         =   "frmAbout.frx":1D94
      ToolTipText     =   "This is a Microsoft trademark. It has nothing to do with this program."
      Top             =   150
      Width           =   720
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright (c) 2002 HISoft (Andrei O. Lisovoi)"
      Height          =   225
      Left            =   1110
      TabIndex        =   3
      Top             =   1920
      Width           =   3330
   End
   Begin VB.Label Label2 
      Caption         =   "This piece of software, allows you to send binary files anywhere in the world, with quite good reading and saving speeds."
      Height          =   660
      Left            =   1125
      TabIndex        =   2
      Top             =   1035
      Width           =   3330
   End
   Begin VB.Label lblVersion 
      Caption         =   "Automatic Version add-in"
      Height          =   255
      Left            =   1140
      TabIndex        =   1
      Top             =   765
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "File Transfer"
      BeginProperty Font 
         Name            =   "Westminster"
         Size            =   27.75
         Charset         =   0
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1140
      TabIndex        =   0
      Top             =   150
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'XP UPGRADE AUTO-INSERT
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Sub cmdClose_Click()
Unload Me
End Sub
Public Sub ShowAbout(IChat As Boolean)
frmAIC.Visible = IChat
End Sub
Private Sub Form_Load()
lblVersion = App.Major & "." & App.Minor & "." & App.Revision
lblIV = lblIV & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Initialize()
'XP UPGRADE AUTO-INSERT
InitCommonControls
End Sub
