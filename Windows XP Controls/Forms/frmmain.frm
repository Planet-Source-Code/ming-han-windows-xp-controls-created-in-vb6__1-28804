VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmmain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Windows XP Controls"
   ClientHeight    =   6390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6030
   Icon            =   "frmmain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   Begin Windows_XP_Controls.xp_canvas xp_canvas1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   11245
      Caption         =   "Windows XP Controls"
      Icon            =   "frmmain.frx":1CFA
      Begin Windows_XP_Controls.xphelp xphelp1 
         Height          =   315
         Left            =   4575
         Top             =   90
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin Windows_XP_Controls.xptopbuttons xpclose 
         Height          =   315
         Left            =   5610
         ToolTipText     =   "Close"
         Top             =   90
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin Windows_XP_Controls.xptopbuttons xpmr 
         Height          =   315
         Left            =   5265
         ToolTipText     =   "Maximized"
         Top             =   90
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Value           =   1
      End
      Begin Windows_XP_Controls.xptopbuttons xpmin 
         Height          =   315
         Left            =   4920
         ToolTipText     =   "Minimized"
         Top             =   90
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         Value           =   2
      End
      Begin PicClip.PictureClip pc2 
         Left            =   3600
         Top             =   0
         _ExtentX        =   582
         _ExtentY        =   29104
         _Version        =   393216
         Rows            =   50
         Picture         =   "frmmain.frx":2294
      End
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   1
         Left            =   5160
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   16711935
         _Version        =   393216
      End
      Begin PicClip.PictureClip pcimg 
         Index           =   1
         Left            =   3360
         Top             =   4200
         _ExtentX        =   13547
         _ExtentY        =   847
         _Version        =   393216
         Cols            =   16
         Picture         =   "frmmain.frx":14716
      End
      Begin PicClip.PictureClip pcimg 
         Index           =   0
         Left            =   3360
         Top             =   3960
         _ExtentX        =   13547
         _ExtentY        =   847
         _Version        =   393216
         Cols            =   16
         Picture         =   "frmmain.frx":20768
      End
      Begin MSComctlLib.ImageList ImageList1 
         Index           =   0
         Left            =   4440
         Top             =   3240
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   16711935
         _Version        =   393216
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   660
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   1164
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1(1)"
         HotImageList    =   "ImageList1(0)"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   5400
         Top             =   960
      End
      Begin Windows_XP_Controls.xpgroupbox xpgroupbox1 
         Height          =   975
         Left            =   1440
         TabIndex        =   6
         Top             =   3720
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1720
         Caption         =   "xpgroupbox"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         Begin Windows_XP_Controls.xpcheckbox xpcheckbox2 
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   2
            Caption         =   "Mixed State"
            ForeColor       =   33023
         End
         Begin Windows_XP_Controls.xpcheckbox xpcheckbox1 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777215
            ForeColor       =   49152
         End
      End
      Begin Windows_XP_Controls.xpgroupbox xpgroupbox2 
         Height          =   1815
         Left            =   3600
         TabIndex        =   1
         Top             =   1800
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3201
         Caption         =   "Message Box Options"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   12648447
         Begin Windows_XP_Controls.xpradiobutton op4 
            Height          =   255
            Left            =   240
            TabIndex        =   2
            Top             =   960
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "vb_exclamation"
         End
         Begin Windows_XP_Controls.xpradiobutton op3 
            Height          =   255
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "vb_question"
         End
         Begin Windows_XP_Controls.xpradiobutton op2 
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "vb_critical"
         End
         Begin Windows_XP_Controls.xpradiobutton op1 
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Value           =   -1  'True
            Caption         =   "vb_information"
         End
         Begin Windows_XP_Controls.xpcmdbutton cmdmsg 
            Height          =   375
            Left            =   360
            TabIndex        =   17
            Top             =   1320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Caption         =   "Message Box"
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
      End
      Begin VB.PictureBox pictxt 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   4080
         ScaleHeight     =   495
         ScaleWidth      =   975
         TabIndex        =   9
         Top             =   3600
         Width           =   975
         Begin VB.TextBox txt1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Text            =   "xptextbox"
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmmain.frx":2C7BA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1095
         Index           =   3
         Left            =   3360
         TabIndex        =   15
         Top             =   4200
         Width           =   2535
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Windows XP Controls"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   21
         Top             =   2040
         Width           =   2085
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hope it is useful! Other themes can also be applied."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   3840
         Width           =   1245
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "There may be furthur updates on this project."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "The distance between each of the xptopbuttons is 2 pixels while the distance between the close button and the edge is 6 pixels."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   855
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   1815
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   3135
      End
      Begin VB.Image Image1 
         Height          =   330
         Left            =   5400
         Top             =   480
         Width           =   330
      End
      Begin VB.Label lbllink1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Here! ( Highly recommended for downloading)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         MouseIcon       =   "frmmain.frx":2C855
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   6000
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   330
         Left            =   5280
         Top             =   480
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   360
         Picture         =   "frmmain.frx":2C9A7
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows XP Controls includes:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   2520
         Width           =   2205
      End
      Begin VB.Label lblinfo 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmmain.frx":2D271
         Height          =   855
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   120
         X2              =   5760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblinfo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "This project was designed according to the Windows XP visual guidelines at:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Index           =   4
         Left            =   3360
         TabIndex        =   14
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         Height          =   495
         Index           =   1
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Image watermark 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   0
         Left            =   3600
         Picture         =   "frmmain.frx":2D2FC
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   2250
      End
      Begin VB.Image watermark 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   1
         Left            =   120
         Picture         =   "frmmain.frx":2DC02
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   2250
      End
      Begin VB.Image watermark 
         Appearance      =   0  'Flat
         Height          =   2250
         Index           =   2
         Left            =   480
         Picture         =   "frmmain.frx":2E7F2
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   2250
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/----------------------------------------------------\
'/Descriptions: Creation of Windows XP windows and    \
'/              controls in Visual Basic              \
'/Created by: Teh Ming Han (teh_minghan@hotmail.com)  \
'/Special thanks: Chris Yates (cyates@neo.rr.com)     \
'/                for trans_colour module             \
'/                                                    \
'/REMEMBER TO VOTE!                                   \
'/                                                    \
'/If you use this code in your program please give me \
'/credit and e-mail me (teh_minghan@hotmail.com) and  \
'/tell me about your program.                         \
'/------Hope you find it useful!---------2001---------\
'/----------------------------------------------------\

Dim n As Integer
'function for hyperlink
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdmsg_Click()
    Load frmmsg
    frmmsg.Show
    'Check for correct icon to be displayed
    With frmmsg
        .imginformation.Visible = False
        .imgquestion.Visible = False
        .imgcritical.Visible = False
        .imgexclamation.Visible = False
        If op1.Value = True Then
            .imginformation.Visible = True
        ElseIf op2.Value = True Then
            .imgcritical.Visible = True
        ElseIf op3.Value = True Then
            .imgquestion.Visible = True
        ElseIf op4.Value = True Then
            .imgexclamation.Visible = True
        End If
    End With
End Sub

Private Sub Form_GotFocus()
xp_canvas1.SetFocus
End Sub

Private Sub Form_Resize()
    xp_canvas1.Height = Me.ScaleHeight
    xp_canvas1.Width = Me.ScaleWidth
    xp_canvas1.make_trans Me
    'Makes the topbuttons move
    xpclose.Left = Me.Width - 405
    xpmr.Left = Me.Width - 750
    xpmin.Left = Me.Width - 1095
    xphelp1.Left = Me.Width - 1440
End Sub

Private Sub lblinfo_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbllink1.BackStyle = 0
End Sub

Private Sub lbllink1_Click()
    'hyperlink
    Dim ret&
    ret = ShellExecute(Me.hwnd, "Open", _
        "http://www.microsoft.com/hwdev/windowsxp/downloads/", _
        "", "", 1)
End Sub

Private Sub lbllink1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbllink1.BackStyle = 1
End Sub

Private Sub Timer1_Timer()
    'Timer for animated Windows XP logo
    On Error Resume Next
    For a = 1 To 3
        Image1.Picture = pc2.GraphicCell(n)
        n = n + 1
    Next a
    If n > 50 Then
        n = 0
        Image1.Picture = pc2.GraphicCell(0)
    End If
End Sub

Private Sub Form_Load()
    Image1.Picture = pc2.GraphicCell(0)
    xptxt txt1(0), pictxt, RGB(240, 232, 224), Normal
    
    'assign toolbar picture
    Dim i As Integer
    For i = 1 To 16
        ImageList1(0).ListImages.Add i, , pcimg(0).GraphicCell(i - 1)
        ImageList1(1).ListImages.Add i, , pcimg(1).GraphicCell(i - 1)
    Toolbar1.Buttons.Item(i).Image = i
    Next i
    
    'sets xp backcolour
    xpcheckbox1.BackColor = RGB(236, 233, 216)
    xpcheckbox2.BackColor = RGB(236, 233, 216)
    xpgroupbox1.BackColor = RGB(236, 233, 216)
    xpgroupbox2.BackColor = RGB(236, 233, 216)
    op1.BackColor = RGB(236, 233, 216)
    op2.BackColor = RGB(236, 233, 216)
    op3.BackColor = RGB(236, 233, 216)
    op4.BackColor = RGB(236, 233, 216)
End Sub

Private Sub watermark_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbllink1.BackStyle = 0
End Sub

Private Sub xp_canvas1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbllink1.BackStyle = 0
End Sub

Private Sub xpclose_Click()
    End
End Sub

Private Sub xpmin_Click()
    Me.WindowState = 1
End Sub

Private Sub xpmr_Click()
    'Change state and button face
    If Me.WindowState = 0 Then
        xpmr.Value = RestoreB
        Me.WindowState = 2
        xpmr.ToolTipText = "Restore"
    ElseIf Me.WindowState = 2 Then
        xpmr.Value = MaxB
        Me.WindowState = 0
        xpmr.ToolTipText = "Maximize"
    End If
End Sub
