VERSION 5.00
Begin VB.Form frmClock
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Windows Clock"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   6330
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1
      Caption         =   "Hide"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6135
      Begin VB.CommandButton Command1
         Caption         =   "Windows Clock"
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
         Picture         =   "frmClock.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   2775
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
         Left            =   3000
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   6135
      Begin VB.CommandButton Command2
         Caption         =   "Windows Clock"
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
         Picture         =   "frmClock.frx":0353
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2775
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
         Left            =   3000
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   480
      Left            =   840
      Picture         =   "frmClock.frx":07CB
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Clock"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TaskBar As Long
Private NotificationArea As Long
Private Clock As Long
Private Sub Command1_Click()

on error resume next

'call the ShowWindow API
    ShowWindow Clock, 0

End Sub
Private Sub Command2_Click()

on error resume next

'call the ShowWindow API
    ShowWindow Clock, 4

End Sub
Private Sub Form_Load()

on error resume next

'assign the value for each variable
    TaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    NotificationArea = FindWindowEx(TaskBar, 0, "TrayNotifyWnd", vbNullString)
    Clock = FindWindowEx(NotificationArea, 0, "TrayClockWClass", vbNullString)

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "To hide the Windows Clock, use the ShowWindow, FindWindowEX and FindWindow API call." & vbCrLf & vbCrLf & "Look at the code for more information on hiding the Windows Clock."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "To show the Windows Clock, use the ShowWindow, FindWindowEX and FindWindow API call." & vbCrLf & vbCrLf & "Look at the code for more information on showing the Windows Clock."

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
