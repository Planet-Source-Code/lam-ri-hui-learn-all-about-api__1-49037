VERSION 5.00
Begin VB.Form frmMyComputer
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API -  My Computer"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   6270
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1
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
      Height          =   3495
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6015
      Begin VB.Frame Frame2
         Height          =   135
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   5775
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
         TabIndex        =   2
         Text            =   "My Computer"
         Top             =   2160
         Width           =   3375
      End
      Begin VB.CommandButton Command2
         Caption         =   "Change"
         Default         =   -1  'True
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
         Left            =   4440
         TabIndex        =   3
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command1
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
         Left            =   4440
         TabIndex        =   1
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
         TabIndex        =   0
         Text            =   "My Computer"
         Top             =   480
         Width           =   3375
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
         Height          =   495
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   9
         Top             =   2880
         Width           =   5295
      End
      Begin VB.Label Label3
         AutoSize        =   -1  'True
         Caption         =   "Tip :"
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
         TabIndex        =   8
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label4
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
         TabIndex        =   7
         Top             =   480
         Width           =   705
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
         Height          =   495
         Left            =   360
         MousePointer    =   2  'Cross
         TabIndex        =   6
         Top             =   1200
         Width           =   5295
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   690
      Left            =   840
      Picture         =   "frmMyComputer.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "My Computer"
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
      TabIndex        =   4
      Top             =   0
      Width           =   3885
   End
End
Attribute VB_Name = "frmMyComputer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

on error resume next

    CreateRegString HKEY_CLASSES_ROOT, "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "", Text1.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"

End Sub
Private Sub Command2_Click()

on error resume next

    CreateRegString HKEY_CLASSES_ROOT, "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "InfoTip", Text2.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "To change My Computer's name, you need to use 3 API call to complete. They are RegCreateKey to create the subkey entry, RegSetValueEx to write value into the key  and RegCloseKey to close the opened registry key." & vbCrLf & vbCrLf & "Look at the code for more information on changing the name of My Computer."

End Sub
Private Sub Label5_Click()

on error resume next

    MsgBox "To change My Computer's tip, you need to use 3 API call to complete. They are RegCreateKey to create the subkey entry, RegSetValueEx to write value into the key  and RegCloseKey to close the opened registry key." & vbCrLf & vbCrLf & "Look at the code for more information on changing the name of My Computer."

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
