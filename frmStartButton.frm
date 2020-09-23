VERSION 5.00
Begin VB.Form frmStartButton
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Windows 'Start' Button"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8325
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4200
      TabIndex        =   17
      Top             =   3840
      Width           =   3975
      Begin VB.TextBox Text5
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Text            =   "0"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text4
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox Text3
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Text            =   "100"
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox Text2
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Text            =   "32"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton Command4
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10
         AutoSize        =   -1  'True
         Caption         =   "Width :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label9
         AutoSize        =   -1  'True
         Caption         =   "Left :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   21
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label8
         AutoSize        =   -1  'True
         Caption         =   "Top :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   20
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label7
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
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   19
         Top             =   1200
         Width           =   3735
      End
      Begin VB.Label Label6
         AutoSize        =   -1  'True
         Caption         =   "Height :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   825
      End
   End
   Begin VB.Frame Frame3
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   3975
      Begin VB.CommandButton Command3
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text1
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Text            =   "Start"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   705
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
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   15
         Top             =   1200
         Width           =   3735
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
      Height          =   2895
      Left            =   4200
      TabIndex        =   12
      Top             =   840
      Width           =   3975
      Begin VB.CommandButton Command2
         Caption         =   "Windows 'Start' Button"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Picture         =   "frmStartButton.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   3735
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
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   13
         Top             =   1920
         Width           =   3735
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
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   3975
      Begin VB.CommandButton Command1
         Caption         =   "Windows 'Start' Button"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Picture         =   "frmStartButton.frx":060D
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   3735
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
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   11
         Top             =   1920
         Width           =   3735
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   435
      Left            =   120
      Picture         =   "frmStartButton.frx":0960
      Top             =   240
      Width           =   1485
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows 'Start' Button"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Width           =   6600
   End
End
Attribute VB_Name = "frmStartButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TaskBar
Private StartButton As Long
Private Sub Command1_Click()

on error resume next

    ShowWindow StartButton, 0

End Sub
Private Sub Command2_Click()

on error resume next

    ShowWindow StartButton, 4

End Sub
Private Sub Command3_Click()

on error resume next

Dim Button As Long
Dim ShellTrayWnd As Long
    ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)

    Button = FindWindowEx(ShellTrayWnd, 0, "Button", vbNullString)
    Call SendMessageByString(Button, WM_SETTEXT, 0&, Text1.Text)

End Sub
Private Sub Command4_Click()

on error resume next

    Call MoveWindow(StartButton, Text4.Text, Text5.Text, Text3.Text, Text2.Text, 1)

End Sub
Private Sub Form_Load()

on error resume next

    TaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    StartButton = FindWindowEx(TaskBar, 0, "button", vbNullString)

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "To hide the Windows 'Start' button, use the ShowWindow API call." & vbCrLf & vbCrLf & "Look at the code for more information on hiding the Windows 'Start' button."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "To show the Windows 'Start' button, use the ShowWindow API call." & vbCrLf & vbCrLf & "Look at the code for more information on showing the Windows 'Start' button."

End Sub
Private Sub Label4_Click()

on error resume next

    MsgBox "The SendMessageByString API call is used to rename the Windows 'Start' button." & vbCrLf & vbCrLf & "Look at the code for more information on renaming the Windows 'Start' button."

End Sub
Private Sub Label7_Click()

on error resume next

    MsgBox "You need only to used 1 API call to resize and reposition the Windows 'Start' button. The API call is MoveWindows." & vbCrLf & vbCrLf & "Look at the code for more information on resizing and repositioning the Windows 'Start' button."

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
