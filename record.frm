VERSION 5.00
Begin VB.Form record 
   Caption         =   "��¼"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   12585
End
Attribute VB_Name = "record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

'view�����С
Me.Height = 8400
Me.Width = 22660


End Sub

Private Sub Form_resize()

If Me.WindowState = 0 Then
    Me.WindowState = 0
    'view����λ��
    Me.Left = 0
    Me.Top = 5400

End If

End Sub
