Attribute VB_Name = "modAPI"
'*******************************************************
'Desktop Transparent
Public Declare Function PaintDesktop Lib "user32" (ByVal hdc As Long) As Long
'*******************************************************
'*******************************************************
'Empty all data on the clipboard
Public Declare Function EmptyClipboard Lib "user32" () As Long
'*******************************************************
'*******************************************************
'Display a message box
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'*******************************************************
'*******************************************************
'Get the computer Name
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'*******************************************************
'*******************************************************
'Get windows tick count
Public Declare Function GetTickCount Lib "kernel32" () As Long
'*******************************************************
'*******************************************************
'Get the current User Name
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'*******************************************************
'*******************************************************
'System parameters
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97
'*******************************************************
'*******************************************************
'The beep command
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'*******************************************************
'*******************************************************
'Set the computer Name
Public Declare Function SetComputerName Lib "kernel32" Alias "SetComputerNameA" (ByVal lpComputerName As String) As Long
'*******************************************************
'*******************************************************
'Execute a program
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'*******************************************************
'*******************************************************
'Play A Sound, Open/Close CD-ROM
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'*******************************************************
'*******************************************************
'Window Functions
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'*******************************************************
'*******************************************************
'Show and Hide Cursor
Public Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long)
'*******************************************************
'*******************************************************
'Set the cursor position
Public Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long)
'*******************************************************
'*******************************************************
'Get the cursor position
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Type PointAPI
    X As Long
    Y As Long
End Type

'*******************************************************
'*******************************************************
'Dial the internet
Public Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Public Const Internet_Autodial_Force_Unattended As Long = 2
'*******************************************************
'*******************************************************
'Disconnect the internet
Public Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
'*******************************************************
'*******************************************************
'Keyboard events. Similar to SendKeys
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
'Assign value
Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91
Public Const VK_CAPITAL = &H14
Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2
'*******************************************************
'*******************************************************
'Swap Mouse Buttons
Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long
'*******************************************************
'*******************************************************
'Close an open registry key
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'*******************************************************
'*******************************************************
'Registry root directory constants
Public Enum RegistryHives
    HKEY_CLASSES_ROOT = &H80000000

    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
    End Enum
'*******************************************************
'*******************************************************
'Set the information of an existing value. Note that if
'you declare the lpData parameter as String, you must
'pass it By Value.
Private Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'*******************************************************
'*******************************************************
'Create a new registry key
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
'*******************************************************
'*******************************************************
'The close windows functions
Public Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
'*******************************************************
'*******************************************************
'Empty Recycle BIN
Public Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
'*******************************************************
'*******************************************************
'The Send Message functions
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'*******************************************************
Public Const WM_SETTEXT = &HC
Public Sub CreateRegString(ByVal enmHive As RegistryHives, ByVal strSubKey As String, ByVal strEntryLabel As String, ByVal strText As String)

on error resume next

'This will put some text into the specified key and entry label. This
'data can be retrieved with the ReadRegString function
Dim lngResult As Long
Dim hKey As Long
Dim strTotalSubKey As String
'Create a complete sub key and entry path to send to the api call
    strTotalSubKey = strSubKey & Chr(0)
'Now create the sub key entry if it does not exist
    lngResult = RegCreateKey(enmHive, strTotalSubKey, hKey)
'If no handle was returned, then exit
    If hKey = 0 Then
        Exit Sub
    End If
'Write the text into the key with the specified entry name
    lngResult = RegSetValueEx(hKey, strEntryLabel, 0&, REG_SZ, ByVal strText, Len(strText))
'Close the opened key and exit
    lngResult = RegCloseKey(hKey)

End Sub
Public Sub RunBrowser(strURL As String, iWindowStyle As Integer, fH As Long)

on error resume next

Dim lSuccess As Long
'-- Shell to default browser
    lSuccess = ShellExecute(fH, "Open", strURL, 0&, 0&, iWindowStyle)

End Sub
