VERSION 5.00
Begin VB.Form frmmsg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Message Box"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3270
   Icon            =   "frmmsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   218
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Windows_XP_Controls.xp_canvas xp_canvas2 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3413
      Caption         =   "Message Box"
      Fixed_Single    =   -1  'True
      Begin Windows_XP_Controls.xptopbuttons xpclose 
         Height          =   315
         Left            =   2850
         ToolTipText     =   "Close"
         Top             =   90
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin Windows_XP_Controls.xpcmdbutton cmdexit 
         Cancel          =   -1  'True
         Default         =   -1  'True
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   1440
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         Caption         =   "Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   10
         Left            =   2520
         Top             =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "This is a Fixed Single Window."
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remember!                        It is fun and easy to generate message boxes."
         Height          =   615
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   1935
      End
      Begin VB.Image imgexclamation 
         Height          =   480
         Left            =   240
         Picture         =   "frmmsg.frx":000C
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgcritical 
         Height          =   480
         Left            =   240
         Picture         =   "frmmsg.frx":0CD6
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imgquestion 
         Height          =   480
         Left            =   240
         Picture         =   "frmmsg.frx":19A0
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image imginformation 
         Height          =   480
         Left            =   240
         Picture         =   "frmmsg.frx":266A
         Top             =   600
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmmsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    xp_canvas2.Height = Me.ScaleHeight
    xp_canvas2.Width = Me.ScaleWidth
    cmdexit.State = Default_
    Beep
    xp_canvas2.make_trans Me
    'Sets window on top
    xp_canvas2.AlwaysOnTop Me, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmmain.xp_canvas1.SetFocus
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Me.SetFocus
End Sub

Private Sub xpclose_Click()
    Unload Me
End Sub
