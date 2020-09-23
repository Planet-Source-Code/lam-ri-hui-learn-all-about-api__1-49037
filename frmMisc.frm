VERSION 5.00
Begin VB.Form frmMisc
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Miscellaneous "
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   9480
   Icon            =   "frmMisc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4
      Caption         =   "Message Box Creator"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4800
      TabIndex        =   29
      Top             =   3480
      Width           =   4335
      Begin VB.CommandButton Command4
         Caption         =   "Create Now"
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
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text6
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   7
         Text            =   "Learn All About API"
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Text5
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Text            =   "frmMisc.frx":0442
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label Label14
         AutoSize        =   -1  'True
         Caption         =   "Caption :"
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
         TabIndex        =   31
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label13
         AutoSize        =   -1  'True
         Caption         =   "Message :"
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
         TabIndex        =   30
         Top             =   840
         Width           =   1005
      End
   End
   Begin VB.Frame Frame7
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   27
      Top             =   7440
      Width           =   9255
      Begin VB.CommandButton Command7
         Caption         =   "Windows Clipboard"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label12
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
         Height          =   495
         Left            =   3720
         MousePointer    =   2  'Cross
         TabIndex        =   28
         Top             =   360
         Width           =   5415
      End
   End
   Begin VB.Frame Frame6
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   5880
      Width           =   9255
      Begin VB.CommandButton Command6
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
         Height          =   855
         Left            =   7200
         TabIndex        =   10
         Top             =   360
         Width           =   1815
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
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label10
         AutoSize        =   -1  'True
         Caption         =   "Click 'Get Now'"
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
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   1650
      End
      Begin VB.Label Label9
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
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   60
      End
   End
   Begin VB.Frame Frame3
      Caption         =   "Computer Name"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   4560
      TabIndex        =   20
      Top             =   840
      Width           =   4815
      Begin VB.CommandButton Command3
         Caption         =   "Change"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   4
         Top             =   480
         Width           =   1095
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
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label8
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
         TabIndex        =   22
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label Label7
         AutoSize        =   -1  'True
         Caption         =   "New Name :"
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
         TabIndex        =   21
         Top             =   600
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   4575
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
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton Command2
         Caption         =   "Email now"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2400
         Picture         =   "frmMisc.frx":051E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6
         AutoSize        =   -1  'True
         Caption         =   "Email Address :"
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
         TabIndex        =   19
         Top             =   480
         Width           =   1590
      End
      Begin VB.Label Label5
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
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   18
         Top             =   1080
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1
      Caption         =   "Beep"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   4335
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
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Text            =   "500"
         Top             =   840
         Width           =   1095
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
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Text            =   "500"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1
         Caption         =   "Beep Now"
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
         Left            =   2760
         Picture         =   "frmMisc.frx":0AFF
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4
         AutoSize        =   -1  'True
         Caption         =   "Duration :"
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
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label3
         AutoSize        =   -1  'True
         Caption         =   "Frequency :"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1200
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
         TabIndex        =   13
         Top             =   1560
         Width           =   4095
      End
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Miscellaneous"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmMisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

on error resume next

    Beep Text1.Text, Text2.Text

End Sub
Private Sub Command2_Click()

on error resume next

    ShellExecute 0&, "Open", "mailto:" & Text3.Text, "", vbNullString, 1

End Sub
Private Sub Command3_Click()

on error resume next

    SetComputerName Text4.Text

End Sub
Private Sub Command4_Click()

on error resume next

    MessageBox Me.hwnd, Text5.Text, Text6.Text, vbOKOnly

End Sub
Private Sub Command5_Click()

on error resume next


End Sub
Private Sub Command6_Click()

on error resume next

    On Error Resume Next
Dim UserNameText As String
    UserNameText = String(200, Chr$(0))

    GetUserName UserNameText, 200
    UserNameText = Left$(UserNameText, InStr(UserNameText, Chr$(0)) - 1)
    Label9.Caption = "The current user is: " & UserNameText
Dim ComputerNameText As String
    ComputerNameText = String(200, Chr$(0))

    GetComputerName ComputerNameText, 200
    ComputerNameText = Left$(ComputerNameText, InStr(ComputerNameText, Chr$(0)) - 1)
    Label10.Caption = "The computer name is: " & ComputerNameText
    Label11.Caption = "Windows has been running for: " & Format(GetTickCount / 60000, "0") & " minutes."

End Sub
Private Sub Command7_Click()

on error resume next

    EmptyClipboard

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label12_Click()

on error resume next

    MsgBox "Except from using EmptyClipboard API call, you can also use 'Clipboard.Clear' to clear all the data on the clipboard." & vbCrLf & vbCrLf & "Look at the code for more information on emptying the clipboard."

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "To beep with different frequency and duration, 1 API call is needed to do this. 'Beep' API call is the only API that can beep with user defined frequency and duration." & vbCrLf & vbCrLf & "Look at the code for more information on beeping with different frquency and duration."

End Sub
Private Sub Label5_Click()

on error resume next

    MsgBox "You can use ShellExecute API call to email someone using your default email software." & vbCrLf & vbCrLf & "Look at the code for more information on emailing."

End Sub
Private Sub Label8_Click()

on error resume next

    MsgBox "Change the computer name using SetComputerName API call." & vbCrLf & vbCrLf & "Look at the code for more information on chaging the computer name."

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
