VERSION 5.00
Begin VB.Form frmCursor
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Cursor"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5370
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4
      Caption         =   "Get"
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
      TabIndex        =   15
      Top             =   6360
      Width           =   5175
      Begin VB.CommandButton Command4
         Caption         =   "Get Now"
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
         Left            =   480
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label11
         AutoSize        =   -1  'True
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
         Left            =   480
         TabIndex        =   20
         Top             =   720
         Width           =   60
      End
      Begin VB.Label Label10
         AutoSize        =   -1  'True
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
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   60
      End
      Begin VB.Label Label9
         AutoSize        =   -1  'True
         Caption         =   "X :"
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
         Width           =   285
      End
      Begin VB.Label Label8
         AutoSize        =   -1  'True
         Caption         =   "Y :"
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
         TabIndex        =   17
         Top             =   720
         Width           =   270
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
         Left            =   2280
         MousePointer    =   2  'Cross
         TabIndex        =   16
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame3
      Caption         =   "Set Cursor's Position"
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
      TabIndex        =   11
      Top             =   4560
      Width           =   5175
      Begin VB.CommandButton Command3
         Caption         =   "Set Now"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox Text2
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text1
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   2
         Top             =   375
         Width           =   1695
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
         TabIndex        =   14
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label5
         AutoSize        =   -1  'True
         Caption         =   "Y :"
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
         TabIndex        =   13
         Top             =   720
         Width           =   270
      End
      Begin VB.Label Label4
         AutoSize        =   -1  'True
         Caption         =   "X :"
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
         TabIndex        =   12
         Top             =   360
         Width           =   285
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
      TabIndex        =   9
      Top             =   2760
      Width           =   5175
      Begin VB.CommandButton Command2
         Caption         =   "Cursor"
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
         Picture         =   "frmCursor.frx":0000
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
         TabIndex        =   10
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
      TabIndex        =   7
      Top             =   960
      Width           =   5175
      Begin VB.CommandButton Command1
         Caption         =   "Cursor"
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
         Picture         =   "frmCursor.frx":0442
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
         TabIndex        =   8
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
      Height          =   480
      Left            =   1320
      Picture         =   "frmCursor.frx":06F2
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cursor"
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
      TabIndex        =   6
      Top             =   0
      Width           =   1845
   End
End
Attribute VB_Name = "frmCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

on error resume next

    ShowCursor (False)

End Sub
Private Sub Command2_Click()

on error resume next

    ShowCursor (True)

End Sub
Private Sub Command3_Click()

on error resume next

    On Error Resume Next
    SetCursorPos Text1.Text, Text2.Text

End Sub
Private Sub Command4_Click()

on error resume next

Dim Result As Long
Dim Pos As PointAPI
    Result = GetCursorPos(Pos)

    If Result <> 0 Then
        Label10.Caption = Pos.X
        Label11.Caption = Pos.Y
        Else
        Exit Sub
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "The ShowCursor API call is used to hide cursor." & vbCrLf & vbCrLf & "Look at the code for more information on hiding cursor."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "The ShowCursor API call is also used to show cursor." & vbCrLf & vbCrLf & "Look at the code for more information on showing cursor."

End Sub
Private Sub Label6_Click()

on error resume next

    MsgBox "To set the cursor's position, you need to use the SetCursorPos API call." & vbCrLf & vbCrLf & "Look at the code for more information on setting the cursor's position."

End Sub
Private Sub Label7_Click()

on error resume next

    MsgBox "To get the cursor's position, you need to use the GetCursorPos API call." & vbCrLf & vbCrLf & "Look at the code for more information on getting the cursor's position."

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
