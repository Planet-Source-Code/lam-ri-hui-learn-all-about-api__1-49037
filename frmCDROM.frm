VERSION 5.00
Begin VB.Form frmCDROM
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - CD-ROM"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5610
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5610
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2
      Caption         =   "Close"
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
      Top             =   2760
      Width           =   5415
      Begin VB.CommandButton Command2
         Caption         =   "CD-ROM Drive"
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
         TabIndex        =   6
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1
      Caption         =   "Open"
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
      Top             =   960
      Width           =   5415
      Begin VB.CommandButton Command1
         Caption         =   "CD-ROM Drive"
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
      Height          =   765
      Left            =   480
      Picture         =   "frmCDROM.frx":0000
      Top             =   0
      Width           =   765
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CD - ROM"
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
      Width           =   3240
   End
End
Attribute VB_Name = "frmCDROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

on error resume next

    MciSendString "set CDAudio door open", vbNullString, 0, 0

End Sub
Private Sub Command2_Click()

on error resume next

    MciSendString "set CDAudio door closed", vbNullString, 0, 0

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "To open the CD-ROM drive, you can use the 'MciSendString' API call." & vbCrLf & vbCrLf & "Look at the code for more information on how to call the API."

End Sub
Private Sub Label3_Click()

on error resume next

    MsgBox "To close the CD-ROM drive, you can also use the 'MciSendString' API call." & vbCrLf & vbCrLf & "Look at the code for more information on how to call the API."

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
