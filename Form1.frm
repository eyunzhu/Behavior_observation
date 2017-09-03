VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'view窗体大小
Me.Height = 5400
Me.Width = 7100


End Sub

Private Sub Form_resize()

If Me.WindowState = 0 Then
    Me.WindowState = 0
    'view窗体位置
    Me.Left = 7100
    Me.Top = 0

End If

End Sub
