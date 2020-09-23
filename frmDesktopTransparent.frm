VERSION 5.00
Begin VB.Form frmDesktopTransparent
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Desktop Transparent"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmDesktopTransparent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

on error resume next

    If Button = 1 Then
        MsgBox "Easy-to-create Desktop Transparent form. This use only use 1 API call, that is PaintDesktop. You must also create a timer to get the form frequently refreshed." & vbCrLf & vbCrLf & "Look at the code for more information on creating a desktop transparent form."
    End If

End Sub
Private Sub Timer1_Timer()

on error resume next

    PaintDesktop frmDesktopTransparent.hdc

End Sub
