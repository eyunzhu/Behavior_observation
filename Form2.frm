VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'view窗体大小
Me.Height = 5400
Me.Width = 8490


End Sub

Private Sub Form_resize()

If Me.WindowState = 0 Then
    Me.WindowState = 0
    'view窗体位置
    Me.Left = 14200
    Me.Top = 0

End If

End Sub

