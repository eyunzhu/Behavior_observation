VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   22800
   LinkTopic       =   "Form2"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   -120
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   1920
      Width           =   10695
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1215
      Left            =   12360
      Top             =   5520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2143
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":009C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from vb_01"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form2.frx":0138
      Height          =   1455
      Left            =   12600
      TabIndex        =   10
      Top             =   2520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox key_con 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "个体编号："
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
      Left            =   480
      TabIndex        =   13
      Top             =   1440
      Width           =   1050
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
      Left            =   1560
      TabIndex        =   12
      Top             =   5040
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4 ："
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
      TabIndex        =   11
      Top             =   5040
      Width           =   420
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
      TabIndex        =   8
      Top             =   4080
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
      Left            =   1560
      TabIndex        =   7
      Top             =   2640
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
      TabIndex        =   6
      Top             =   3360
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
      Left            =   1440
      TabIndex        =   5
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0 ："
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
      Top             =   2040
      Width           =   420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 ："
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
      Top             =   2640
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3 ："
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
      TabIndex        =   2
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2 ："
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
      Top             =   3240
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
  

  
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus
    
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus
ElseIf KeyCode = vbKey0 Or KeyCode = vbKeyNumpad0 Then
'0
    Form2.Text6.SetFocus

  

  End If
  
  key_con.Text = ""
  'Form1.key_con.SetFocus
  
  
End Sub

'----- Text6_ 键盘控制实现
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)

'Dim text$
'text = Text3.text
'ElseIf KeyCode = vbKeyUp Then
'Text3.text = "" & Text3.text & "您按了上键"
  
  If KeyCode = vbKeyC And Shift = 1 Then
     MsgBox "你按下的是Shift键+字母C键组合，即输入大写字母C"
     
  ElseIf KeyCode = vbKeyUp Then
  



  
ElseIf KeyCode = vbKeyF1 Then
'F1
    Form2.Label6.Caption = "455"
    Form2.Text1.Text = "" & Form2.Text1.Text & "" & Form1.Text3.Text & ":" & Form1.WindowsMediaPlayer1.Controls.currentPositionString & "  "
    
    'Text3.Text = "" & Text3.Text & "您按了上键"

  End If
  
  key_con.Text = ""
  'Form1.key_con.SetFocus
  
  
End Sub
