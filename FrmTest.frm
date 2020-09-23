VERSION 5.00
Begin VB.Form FrmTest 
   BorderStyle     =   0  'Kein
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OldX
Private OldY As Integer
Private MoveIt As Boolean


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'If the button is the pressed keep the mouse coordinates to move the form.
    If Button = vbLeftButton Then
        OldX = X
        OldY = Y
        MoveIt = True
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MoveIt Then
'Set new window position
        Me.Left = Me.Left + X - OldX
        Me.Top = Me.Top + Y - OldY
    End If

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveIt = False
End Sub
