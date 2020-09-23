VERSION 5.00
Begin VB.Form frmTaskbar
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Windows Taskbar"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   6810
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6810
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
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6615
      Begin VB.CommandButton Command1
         Caption         =   "Windows Taskbar"
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
         Picture         =   "frmTaskbar.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   360
         Width           =   3495
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
         Left            =   3720
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   600
         Width           =   2775
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
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   6615
      Begin VB.CommandButton Command2
         Caption         =   "Windows Taskbar"
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
         Picture         =   "frmTaskbar.frx":0353
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   4215
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
         Height          =   1335
         Left            =   4440
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Windows Taskbar"
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
      TabIndex        =   2
      Top             =   0
      Width           =   5100
   End
End
Attribute VB_Name = "frmTaskbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TaskBar As Long
Private Sub Command1_Click()

on error resume next

    ShowWindow TaskBar, 0

End Sub
Private Sub Command2_Click()

on error resume next

    ShowWindow TaskBar, 4

End Sub
Private Sub Form_Load()

on error resume next

    TaskBar = FindWindow("Shell_TrayWnd", vbNullString)

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "It is simple to make the Windows Taskbar invisible. Just use the ShowWindow API call." & vbCrLf & vbCrLf & "Look at the code for more information on hiding the taskbar."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "It is also simple to make the Windows Taskbar visible. Just use the ShowWindow API call." & vbCrLf & vbCrLf & "Look at the code for more information on showing the taskbar."

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
