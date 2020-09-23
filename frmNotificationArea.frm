VERSION 5.00
Begin VB.Form frmNotificationArea
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Windows Nofication Area"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   10665
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5400
      TabIndex        =   5
      Top             =   840
      Width           =   5175
      Begin VB.CommandButton Command2
         Caption         =   "Windows Notification Area"
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
         Picture         =   "frmNotificationArea.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2055
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
         Top             =   480
         Width           =   2775
      End
   End
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
      TabIndex        =   3
      Top             =   840
      Width           =   5175
      Begin VB.CommandButton Command1
         Caption         =   "Windows Notification Area"
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
         Picture         =   "frmNotificationArea.frx":0799
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   2055
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
         TabIndex        =   4
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   450
      Left            =   480
      Picture         =   "frmNotificationArea.frx":0AEC
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Notification Area"
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
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   7770
   End
End
Attribute VB_Name = "frmNotificationArea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TaskBar
Private NotificationArea As Long
Private Sub Command1_Click()

on error resume next

    ShowWindow NotificationArea, 0

End Sub
Private Sub Command2_Click()

on error resume next

    ShowWindow NotificationArea, 4

End Sub
Private Sub Form_Load()

on error resume next

    TaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    NotificationArea = FindWindowEx(TaskBar, 0, "TrayNotifyWnd", vbNullString)

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

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
