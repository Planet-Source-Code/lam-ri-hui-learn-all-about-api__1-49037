VERSION 5.00
Begin VB.Form frmKeyboard
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Keyboard"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5625
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4
      Caption         =   "Disable"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   5415
      Begin VB.CommandButton Command4
         Caption         =   "Ctrl - Alt - Delete"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label5
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Here For More Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   14
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame5
      Caption         =   "Enable"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   5415
      Begin VB.CommandButton Command5
         Caption         =   "Ctrl - Alt - Delete"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label6
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Here For More Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   15
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1
      Caption         =   "Toggle"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5415
      Begin VB.Frame Frame2
         Height          =   135
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   5175
      End
      Begin VB.Frame Frame2
         Height          =   135
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   5175
      End
      Begin VB.CommandButton Command3
         Caption         =   "Num Lock"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Picture         =   "frmKeyboard.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton Command2
         Caption         =   "Scroll Lock"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Picture         =   "frmKeyboard.frx":0942
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton Command1
         Caption         =   "Caps Lock"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   120
         Picture         =   "frmKeyboard.frx":1284
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label4
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Here For More Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label3
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Here For More Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Label Label2
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Click Here For More Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   5
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   660
      Left            =   960
      Picture         =   "frmKeyboard.frx":1BC6
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

on error resume next

Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
    Numlockstate = keys(VK_NUMLOCK)

    If Numlockstate <> True Then
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If

End Sub
Private Sub Command2_Click()

on error resume next

Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
    Numlockstate = keys(VK_NUMLOCK)

    If Numlockstate <> True Then
        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If

End Sub
Private Sub Command3_Click()

on error resume next

Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
    Numlockstate = keys(VK_NUMLOCK)

    If Numlockstate <> True Then
        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If

End Sub
Private Sub Command4_Click()

on error resume next

Dim ret As Integer
Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)


End Sub
Private Sub Command5_Click()

on error resume next

Dim ret As Integer
Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)


End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "Use 'keybd_event' API call to toggle Caps Lock." & vbCrLf & vbCrLf & "Look at the code for more information on toggling Caps Lock."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "You can use 'keybd_event' API call to toggle Scroll Lock." & vbCrLf & vbCrLf & "Look at the code for more information on toggling Scroll Lock."

End Sub
Private Sub Label4_Click()

on error resume next

    MsgBox "Use 'keybd_event' API call also to toggle Num Lock." & vbCrLf & vbCrLf & "Look at the code for more information on toggling Num Lock."

End Sub
Private Sub Label5_Click()

on error resume next

    MsgBox "Use 'SystemParametersInfo' API to disable Ctrl - Alt - Delete key." & vbCrLf & vbCrLf & "Look at the code for more information on disabling the Ctrl - Alt - Delete key."

End Sub
Private Sub Label6_Click()

on error resume next

    MsgBox "Use 'SystemParametersInfo' API to enable Ctrl - Alt - Delete key." & vbCrLf & vbCrLf & "Look at the code for more information on enabling the Ctrl - Alt - Delete key."

End Sub
Private Sub Timer1_Timer()

on error resume next

'Declare the variables for the value of Red, Green and Blue
Dim R
Dim G
Dim B
'Initialize the random engine
    Randomize
'Generates the value for Red, Green and Blue
    R = Rnd * 255
    G = Rnd * 255
    B = Rnd * 255
'Change the colour for Label1
    Label1.ForeColor = RGB(R, G, B)

End Sub
