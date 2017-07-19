VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15630
   LinkTopic       =   "Form2"
   ScaleHeight     =   10305
   ScaleWidth      =   15630
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Text            =   "Text2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox key_con 
      Height          =   375
      Left            =   2400
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label11 
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
      Left            =   1680
      TabIndex        =   10
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label Label9 
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
      Left            =   1560
      TabIndex        =   9
      Top             =   3000
      Width           =   525
   End
   Begin VB.Label Label8 
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
      Left            =   1680
      TabIndex        =   8
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label Label7 
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
      Left            =   1560
      TabIndex        =   7
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label Label6 
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
      Left            =   1680
      TabIndex        =   6
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A ："
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
      Left            =   720
      TabIndex        =   5
      Top             =   3720
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A ："
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
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A ："
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
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A ："
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
      Left            =   840
      TabIndex        =   2
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A ："
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
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "行为记录"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8040
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----- key_con 键盘控制实现
Private Sub key_con_KeyDown(KeyCode As Integer, Shift As Integer)

'Dim text$
'text = Text3.text
'ElseIf KeyCode = vbKeyUp Then
'Text3.text = "" & Text3.text & "您按了上键"
  
  If KeyCode = vbKeyC And Shift = 1 Then
     MsgBox "你按下的是Shift键+字母C键组合，即输入大写字母C"
     
  ElseIf KeyCode = vbKeyUp Then
  

  
ElseIf KeyCode = vbKeyDown Then
'停止
    WindowsMediaPlayer1.Controls.stop
ElseIf KeyCode = vbKeySpace Then
'暂停
    WindowsMediaPlayer1.Controls.pause
ElseIf KeyCode = vbKeyReturn Then
'播放
     WindowsMediaPlayer1.Controls.play

ElseIf KeyCode = vbKeyLeft Then
'加速后退"
WindowsMediaPlayer1.Controls.fastReverse
ElseIf KeyCode = vbKeyRight Then
'加速播放
WindowsMediaPlayer1.Controls.fastForward
ElseIf KeyCode = vbKeyA Then
MsgBox "您按了A键!"

ElseIf KeyCode = vbKeyAdd Then
'打开记录页
    Form2.Visible = True
   
    
    


  End If
  
  key_con.Text = ""
  Form1.key_con.SetFocus
  
  
End Sub
