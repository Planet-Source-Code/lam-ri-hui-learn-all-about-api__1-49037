VERSION 5.00
Begin VB.Form frmDesktop
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Desktop"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5655
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3
      Caption         =   "Click here to see the Deskop Transparent form."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   5415
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
      TabIndex        =   6
      Top             =   840
      Width           =   5415
      Begin VB.CommandButton Command1
         Caption         =   "Desktop"
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
         Picture         =   "frmDesktop.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   2175
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
         Left            =   2400
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   480
         Width           =   2895
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
      TabIndex        =   4
      Top             =   2640
      Width           =   5415
      Begin VB.CommandButton Command2
         Caption         =   "Desktop"
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
         Picture         =   "frmDesktop.frx":02B0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   2175
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
         Left            =   2400
         MousePointer    =   2  'Cross
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   720
      Left            =   1080
      Picture         =   "frmDesktop.frx":0837
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desktop"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   0
      Width           =   2340
   End
End
Attribute VB_Name = "frmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare the variable
Private Icons As Long
Private Sub Command1_Click()

on error resume next

    ShowWindow Icons, 0

End Sub
Private Sub Command2_Click()

on error resume next

    ShowWindow Icons, 4

End Sub
Private Sub Command3_Click()

on error resume next

    frmDesktopTransparent.Show 1

End Sub
Private Sub Form_Load()

on error resume next

'Assign value for the variable
    Icons = FindWindowEx(0&, 0&, "Progman", vbNullString)

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "It is extremely easy to hide the desktop! You need only an API call, that is FindWindowEx." & vbCrLf & vbCrLf & "Look at the code for more information on hiding the desktop."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "It is extremely easy to show the desktop too! You need only an API call, that is FindWindowEx." & vbCrLf & vbCrLf & "Look at the code for more information on showing the desktop."

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
