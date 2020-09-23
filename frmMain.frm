VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API  by Lam Ri Hui (rihui@email.com)"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   13065
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   13065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "Miscellaneous"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   6375
   End
   Begin VB.CommandButton Command13 
      Caption         =   "API to Shutdown, Log Off and Restart the system"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   8760
      Picture         =   "frmMain.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command12 
      Caption         =   "API related with the Keyboard"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10920
      Picture         =   "frmMain.frx":1C5D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command11 
      Caption         =   "API related with Internet"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6600
      Picture         =   "frmMain.frx":257B
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command10 
      Caption         =   "API related with My Computer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      Picture         =   "frmMain.frx":2B2D
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "API related with Recycle Bin"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6600
      Picture         =   "frmMain.frx":2FB6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "API related with the Windows Notification Area"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      Picture         =   "frmMain.frx":3417
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
      Caption         =   "API related with the Windows clock"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2280
      Picture         =   "frmMain.frx":3BB0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "API related with Desktop"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2280
      Picture         =   "frmMain.frx":4028
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "API related with the mouse"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10920
      Picture         =   "frmMain.frx":45AF
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "API related with the cursor"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   10920
      Picture         =   "frmMain.frx":4A7D
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "API related with the Windows Taskbar"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6600
      Picture         =   "frmMain.frx":4D87
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "API related with the CD-ROM drive"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Picture         =   "frmMain.frx":73FD
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "API related with the Windows 'Start' button"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Picture         =   "frmMain.frx":7993
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "What do you wish to learn?"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   0
      TabIndex        =   15
      Top             =   960
      Width           =   5970
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Learn All About API! Master API in 1 day!"
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
      Left            =   600
      TabIndex        =   14
      Top             =   0
      Width           =   12090
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

On Error Resume Next

    frmStartButton.Show
    frmMain.Hide

End Sub
Private Sub Command10_Click()

On Error Resume Next

    frmMyComputer.Show
    frmMain.Hide

End Sub
Private Sub Command11_Click()

On Error Resume Next

    frmInternet.Show
    frmMain.Hide

End Sub
Private Sub Command12_Click()

On Error Resume Next

    frmKeyboard.Show
    frmMain.Hide

End Sub
Private Sub Command13_Click()

On Error Resume Next

    frmShutLogRes.Show
    frmMain.Hide

End Sub
Private Sub Command14_Click()

On Error Resume Next

    frmMisc.Show
    Me.Hide

End Sub
Private Sub Command2_Click()

On Error Resume Next

    frmCDROM.Show
    frmMain.Hide

End Sub
Private Sub Command3_Click()

On Error Resume Next

    frmTaskbar.Show
    frmMain.Hide

End Sub
Private Sub Command4_Click()

On Error Resume Next

    frmCursor.Show
    frmMain.Hide

End Sub
Private Sub Command5_Click()

On Error Resume Next

    frmMouse.Show
    frmMain.Hide

End Sub
Private Sub Command6_Click()

On Error Resume Next

    frmDesktop.Show
    frmMain.Hide

End Sub
Private Sub Command7_Click()

On Error Resume Next

    frmClock.Show
    frmMain.Hide

End Sub
Private Sub Command8_Click()

On Error Resume Next

    frmNotificationArea.Show
    frmMain.Hide

End Sub
Private Sub Command9_Click()

On Error Resume Next

    frmRecycleBin.Show
    frmMain.Hide

End Sub
Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

Dim res
    res = MsgBox("Do you want to vote for this program?", vbYesNo, "Vote")

    If res = vbYes Then
        Call RunBrowser("http://www.planetsourcecode.com/vb/scripts/showcode.asp?lngWId=1&txtCodeId=49037", 10, 1): MsgBox "Thanks for spending time to vote my program.", , "Thanks"
    End If

End Sub
Private Sub Timer1_Timer()

On Error Resume Next

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
'Generates the value for Red, Green and Blue
    R = Rnd * 255
    G = Rnd * 255
    B = Rnd * 255
'Change the colour for Label1
    Label2.ForeColor = RGB(R, G, B)

End Sub
