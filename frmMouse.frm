VERSION 5.00
Begin VB.Form frmMouse
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API -  Mouse"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5385
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1
      Caption         =   "Button"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5175
      Begin VB.CommandButton Command2
         Caption         =   "Right Mouse Button"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         Picture         =   "frmMouse.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton Command1
         Caption         =   "Left Mouse Button"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         Picture         =   "frmMouse.frx":0A4B
         Style           =   1  'Graphical
         TabIndex        =   0
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
         TabIndex        =   5
         Top             =   2760
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
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   870
      Left            =   1200
      Picture         =   "frmMouse.frx":1496
      Top             =   0
      Width           =   795
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mouse"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   0
      Width           =   1875
   End
End
Attribute VB_Name = "frmMouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

on error resume next

    SwapMouseButton (False)

End Sub
Private Sub Command2_Click()

on error resume next

    SwapMouseButton (True)

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "Set false for SwapMouseButton API call will set the primary function to left button." & vbCrLf & vbCrLf & "Look at the code for more information on setting the primary functions."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "Set true for SwapMouseButton API call will set the primary function to right button." & vbCrLf & vbCrLf & "Look at the code for more information on setting the primary functions."

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
