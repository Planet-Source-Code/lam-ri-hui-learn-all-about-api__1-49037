VERSION 5.00
Begin VB.Form frmShutLogRes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Shutdon, Log Off and Restart"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   10530
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Turn Off Computer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   10335
      Begin VB.CommandButton Command3 
         Caption         =   "Restart"
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
         Left            =   7080
         Picture         =   "frmShutLogRes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Log Off"
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
         Picture         =   "frmShutLogRes.frx":0510
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Shutdown"
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
         Left            =   3600
         Picture         =   "frmShutLogRes.frx":0945
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1095
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
         Height          =   1215
         Left            =   1320
         MousePointer    =   2  'Cross
         TabIndex        =   7
         Top             =   600
         Width           =   1935
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
         Height          =   1215
         Left            =   8280
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   600
         Width           =   1935
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
         Height          =   1215
         Left            =   4800
         MousePointer    =   2  'Cross
         TabIndex        =   5
         Top             =   600
         Width           =   1935
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
      Caption         =   "Shutdown, Log Off and Restart"
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
      Left            =   840
      TabIndex        =   3
      Top             =   0
      Width           =   8685
   End
End
Attribute VB_Name = "frmShutLogRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const REG_SZ As Long = 1
Const EWX_SHUTDOWN As Long = 1
Private Const EWX_LOGOFF As Long = 0
Private Const EWX_REBOOT As Long = 2

Private Sub Command1_Click()

On Error Resume Next

Dim Dummy
    Dummy = ExitWindowsEx(EWX_SHUTDOWN, 0&)


End Sub
Private Sub Command2_Click()

On Error Resume Next

Dim Dummy
    Dummy = ExitWindowsEx(EWX_LOGOFF, 0&)


End Sub
Private Sub Command3_Click()

On Error Resume Next

Dim Dummy
    Dummy = ExitWindowsEx(EWX_REBOOT, 0&)


End Sub
Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next

    frmMain.Show

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

End Sub
