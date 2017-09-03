VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form view 
   Caption         =   "观测"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   11280
   Begin VB.Timer Timer1 
      Left            =   8520
      Top             =   2760
   End
   Begin VB.CommandButton Command3 
      Caption         =   "跳转"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   3840
      Width           =   735
   End
   Begin VB.TextBox jump 
      Height          =   390
      Left            =   4920
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开"
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox url 
      Height          =   390
      Left            =   1200
      TabIndex        =   7
      Text            =   "D:\4我的相关------------------照片等\joke_essay\1463666139048.mp4"
      Top             =   4320
      Width           =   3495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   600
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开文件"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "跳转(秒)："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3840
      TabIndex        =   5
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label video_long 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00;00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3240
      TabIndex        =   4
      Top             =   3960
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "视频总长："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2280
      TabIndex        =   3
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "输入地址："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "视频播放："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1050
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6840
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   12065
      _cy             =   6641
   End
End
Attribute VB_Name = "view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Sub Form_Load()
Timer1.Interval = 1

'view窗体大小
Me.Height = 5400
Me.Width = 7100


End Sub

Private Sub Form_resize()

If Me.WindowState = 0 Then
    Me.WindowState = 0
    'view窗体位置
    Me.Left = 0
    Me.Top = 0

End If

End Sub



Private Sub Timer1_Timer()
If GetAsyncKeyState(vbKeyF9) Then
MsgBox "F9"
End If
End Sub



'------延时函数Delay
Sub Delay(Seconds&)
t& = Timer
Delay: DoEvents
If Timer < t + Seconds Then GoTo Delay
End Sub

Private Sub Command1_Click()
' 设置“CancelError”为 True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' 设置标志
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' 设置过滤器
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    ' 指定缺省的过滤器
    CommonDialog1.FilterIndex = 2
    ' 显示“打开”对话框
    CommonDialog1.ShowOpen
    ' 显示选定文件的名字
    url = CommonDialog1.FileName
   ' MsgBox CommonDialog1.filename '显示路径
   WindowsMediaPlayer1.url = url
   '视频路径设置为资源管理器打开的

    '延时
    Delay (1.5)
    '输出视频总长
    video_long.Caption = WindowsMediaPlayer1.currentMedia.durationString
   
    Exit Sub
ErrHandler:
    ' 用户按了“取消”按钮
    Exit Sub

End Sub

Private Sub Command2_Click()
WindowsMediaPlayer1.url = url
'延时
Delay (1.5)
'输出视频总长
video_long.Caption = WindowsMediaPlayer1.currentMedia.durationString

End Sub

Private Sub Command3_Click()
'跳转播放位置
WindowsMediaPlayer1.Controls.currentPosition = jump.Text
End Sub

