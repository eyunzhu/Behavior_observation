VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   Caption         =   "��Ϊ�۲�"
   ClientHeight    =   10545
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   15930
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   12240
      TabIndex        =   30
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   12240
      TabIndex        =   29
      Top             =   5280
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   12240
      TabIndex        =   28
      Top             =   4680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   12240
      TabIndex        =   27
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   12240
      TabIndex        =   26
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   15840
      TabIndex        =   14
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   11640
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��ת"
      Height          =   375
      Left            =   17160
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   14280
      TabIndex        =   9
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox key_con 
      Height          =   390
      Left            =   18720
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���벥��"
      Height          =   375
      Left            =   17160
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox url 
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11640
      TabIndex        =   2
      Text            =   "D:\4�ҵ����------------------��Ƭ��\joke_essay\ü��ѩ_����.mp4"
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F5 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11640
      TabIndex        =   25
      Top             =   6000
      Width           =   525
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F4 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11640
      TabIndex        =   24
      Top             =   5400
      Width           =   525
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F3 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11640
      TabIndex        =   23
      Top             =   4800
      Width           =   525
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ϊ��ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10560
      TabIndex        =   22
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F1 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11640
      TabIndex        =   21
      Top             =   3600
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F2 ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11640
      TabIndex        =   20
      Top             =   4200
      Width           =   525
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   19
      Top             =   9720
      Width           =   420
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   18
      Top             =   10200
      Width           =   420
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   17
      Top             =   8280
      Width           =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   16
      Top             =   8760
      Width           =   420
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   15
      Top             =   9240
      Width           =   420
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�۲���������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   14280
      TabIndex        =   13
      Top             =   2160
      Width           =   1470
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ���ƣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10560
      TabIndex        =   12
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����λ��(��)��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   12720
      TabIndex        =   8
      Top             =   1440
      Width           =   1470
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00;00"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   11760
      TabIndex        =   7
      Top             =   1440
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ƶ�ܳ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10560
      TabIndex        =   6
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ƶ��ַ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   10560
      TabIndex        =   3
      Top             =   960
      Width           =   1050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ϊ�۲�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6600
      TabIndex        =   1
      Top             =   240
      Width           =   1200
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   9975
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
      _cx             =   17595
      _cy             =   11668
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------��ʱ����Delay
Sub Delay(Seconds&)
t& = Timer
Delay: DoEvents
If Timer < t + Seconds Then GoTo Delay
End Sub

'------��Ƶ����
Private Sub Command1_Click()
WindowsMediaPlayer1.url = url
'WindowsMediaPlayer1.Controls.currentPosition = 22

'��ʱ
Delay (1.5)
'�����Ƶ�ܳ�
Label4.Caption = WindowsMediaPlayer1.currentMedia.durationString
key_con.SetFocus

End Sub



'������ת
Private Sub Command2_Click()
WindowsMediaPlayer1.Controls.currentPosition = Text2.Text
key_con.SetFocus

End Sub



Private Sub Command3_Click()
Load Text1(1)

Text1(1).Left = 0

Text1(1).Visible = True
End Sub

Private Sub Command4_Click()
Unload Text1(1)
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = vbKeyReturn Then
    WindowsMediaPlayer1.Controls.currentPosition = Text2.Text
    key_con.SetFocus
  End If
End Sub

'----- key_con ���̿���ʵ��
Private Sub key_con_KeyDown(KeyCode As Integer, Shift As Integer)

'Dim text$
'text = Text3.text
'ElseIf KeyCode = vbKeyUp Then
'Text3.text = "" & Text3.text & "�������ϼ�"
  
  If KeyCode = vbKeyC And Shift = 1 Then
     MsgBox "�㰴�µ���Shift��+��ĸC����ϣ��������д��ĸC"
     
  ElseIf KeyCode = vbKeyUp Then
  

  
ElseIf KeyCode = vbKeyDown Then
'ֹͣ
    WindowsMediaPlayer1.Controls.stop
ElseIf KeyCode = vbKeySpace Then
'��ͣ
    WindowsMediaPlayer1.Controls.pause
ElseIf KeyCode = vbKeyReturn Then
'����
     WindowsMediaPlayer1.Controls.play

ElseIf KeyCode = vbKeyLeft Then
'���ٺ���"
WindowsMediaPlayer1.Controls.fastReverse
ElseIf KeyCode = vbKeyRight Then
'���ٲ���
WindowsMediaPlayer1.Controls.fastForward
ElseIf KeyCode = vbKeyA Then
MsgBox "������A��!"

ElseIf KeyCode = vbKeyAdd Then
'�򿪼�¼ҳ
    Form2.Visible = True
   
    
    


  End If
  
  key_con.Text = ""
  
End Sub




