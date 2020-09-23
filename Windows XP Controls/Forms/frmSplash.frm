VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF0000&
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   3600
      Top             =   2400
   End
   Begin VB.PictureBox picpgb2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   240
      ScaleHeight     =   19
      ScaleMode       =   0  'User
      ScaleWidth      =   415
      TabIndex        =   1
      Top             =   3720
      Width           =   6225
   End
   Begin VB.Image imgpgb1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      Picture         =   "frmSplash.frx":000C
      Top             =   3000
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows XP Controls"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   4005
   End
   Begin VB.Image imgLogo 
      Height          =   1785
      Left            =   240
      Picture         =   "frmSplash.frx":03B2
      Stretch         =   -1  'True
      Top             =   555
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows XP Controls"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   4005
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim distance As Integer

Private Sub Form_Load()
    distance = 4
    Horizontal Me, RGB(131, 166, 244), RGB(33, 120, 224)
    picpgb2.PaintPicture imgpgb1, 0, 0, 4, 19, 0, 0, 4, 19
    picpgb2.PaintPicture imgpgb1, 4, 0, picpgb2.Width - 9, 19, 4, 0, 10, 19
    picpgb2.PaintPicture imgpgb1, picpgb2.Width - 5, 0, 5, 19, 14, 0, 5, 19
End Sub

Private Sub Form_Terminate()
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    For i = 1 To 2
        picpgb2.PaintPicture imgpgb1.Picture, distance, 4, 8, 12, 23, 5, 8, 12
        distance = distance + 10
    Next i
    If distance > picpgb2.Width - 5 Then
        Timer1.Enabled = False
        Unload Me
        Load frmmain
        frmmain.Show
    End If
End Sub
