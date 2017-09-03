VERSION 5.00
Begin VB.MDIForm index_MDI 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   12075
   ClientLeft      =   225
   ClientTop       =   765
   ClientWidth     =   22800
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu project 
      Caption         =   "项目"
   End
   Begin VB.Menu gc 
      Caption         =   "观测"
   End
   Begin VB.Menu jl 
      Caption         =   "记录"
   End
   Begin VB.Menu fd 
      Caption         =   "12"
   End
   Begin VB.Menu lo 
      Caption         =   "34"
   End
End
Attribute VB_Name = "index_MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fd_Click()
Form1.Visible = False
Form1.Visible = True
End Sub

Private Sub jl_Click()
'记录按钮
record.Visible = False
record.Visible = True
End Sub




Private Sub gc_Click()
'观测按钮
view.Visible = False
view.Visible = True
view.jump.SetFocus
End Sub

Private Sub lo_Click()
Form2.Visible = False
Form2.Visible = True
End Sub

