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
      Caption         =   "��Ŀ"
   End
   Begin VB.Menu gc 
      Caption         =   "�۲�"
   End
   Begin VB.Menu jl 
      Caption         =   "��¼"
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
'��¼��ť
record.Visible = False
record.Visible = True
End Sub




Private Sub gc_Click()
'�۲ⰴť
view.Visible = False
view.Visible = True
view.jump.SetFocus
End Sub

Private Sub lo_Click()
Form2.Visible = False
Form2.Visible = True
End Sub

