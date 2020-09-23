VERSION 5.00
Begin VB.Form frmRecycleBin
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn All About API - Recycle Bin"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   6240
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3
      Caption         =   "Empty"
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
      TabIndex        =   12
      Top             =   4440
      Width           =   6015
      Begin VB.CommandButton Command3
         Caption         =   "Recycle Bin"
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
         Picture         =   "frmRecycleBin.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   2055
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
         Left            =   2520
         MousePointer    =   2  'Cross
         TabIndex        =   13
         Top             =   480
         Width           =   3375
      End
   End
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
      TabIndex        =   6
      Top             =   840
      Width           =   6015
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
         Text            =   "Recycle Bin"
         Top             =   480
         Width           =   3375
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
         Text            =   "Contains the files and folders that you have deleted."
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Frame Frame2
         Height          =   135
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   5775
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
         TabIndex        =   11
         Top             =   1200
         Width           =   5295
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
         TabIndex        =   10
         Top             =   480
         Width           =   705
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
         TabIndex        =   9
         Top             =   2160
         Width           =   450
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
         TabIndex        =   8
         Top             =   2880
         Width           =   5295
      End
   End
   Begin VB.Timer Timer1
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1
      Height          =   645
      Left            =   840
      Picture         =   "frmRecycleBin.frx":0433
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Label Label1
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Recycle Bin"
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
      Left            =   1560
      TabIndex        =   5
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmRecycleBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

on error resume next

    CreateRegString HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", Text1.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"

End Sub
Private Sub Command2_Click()

on error resume next

    CreateRegString HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", Text2.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"

End Sub
Private Sub Command3_Click()

on error resume next

    Call MakeRecycleBinEmpty("C:\", False, False, False)

End Sub
Private Sub Form_Unload(Cancel As Integer)

on error resume next

    frmMain.Show

End Sub
Private Sub Label2_Click()

on error resume next

    MsgBox "It is very easy to change the Recycle Bin's name, you need to use 3 API call to complete. They are RegCreateKey to create the subkey entry, RegSetValueEx to write value into the key  and RegCloseKey to close the opened registry key." & vbCrLf & vbCrLf & "Look at the code for more information on changing the name of Recycle Bin."

End Sub
Private Sub Label5_Click()

on error resume next

    MsgBox "It is also very easy to change the Recycle Bin's tip, you only need to use 3 API call to complete. They are RegCreateKey to create the subkey entry, RegSetValueEx to write value into the key  and RegCloseKey to close the opened registry key." & vbCrLf & vbCrLf & "Look at the code for more information on changing the tip of Recycle Bin."

End Sub
Private Sub Label6_Click()

on error resume next

    MsgBox "The SHEmptyRecycleBin API call is used to empty the Recycle Bin." & vbCrLf & vbCrLf & "Look at the code for more information on emptying the Recycle Bin."

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
Private Sub MakeRecycleBinEmpty(Optional ByVal Drive As String, Optional NoConfirmation As Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)

on error resume next

Dim hwnd
Dim Flags As Long
    On Error Resume Next

    hwnd = Screen.ActiveForm.hwnd
    If Len(Drive) > 0 Then Drive = Left$(Drive, 1) & ":\"
    Flags = (NoConfirmation And &H1) Or (NoProgress And &H2) Or (NoSound And &H4)
    SHEmptyRecycleBin hwnd, Drive, Flags

End Sub
