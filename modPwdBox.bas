Attribute VB_Name = "modPwdBox"
'This Code was written by Dave Andrews
'Feel free to use or modify this module freely
'Special thanks to Joseph Huntley for the skeleton of API forms.

Option Explicit
Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function defWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long

Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type


Private Type POINTAPI
    x As Long
    y As Long
End Type


Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
'Editbox Constants
Const ES_PASSWORD = &H20&
Const ES_CENTER = &H1&
Const EM_SETPASSWORDCHAR = &HCC
Const EM_GETLINE = &HC4
Const EM_LINELENGTH = &HC1

'------Button Constants
Const BS_USERBUTTON = &H8&
Const BS_CENTER = 768
Const BS_PUSHBUTTON = &H0&
Const BS_AUTORADIOBUTTON = &H9&
Const BS_PUSHLIKE = &H1000&
Const BS_LEFTTEXT = &H20&
Const BM_SETSTATE = &HF3
Const BM_GETSTATE = &HF2
Const BM_SETCHECK = &HF1
Const BM_GETCHECK = &HF0
'-----------Window Style Constants
Const WS_BORDER = &H800000
Const WS_CHILD = &H40000000
Const WS_OVERLAPPED = &H0&
Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
Const WS_SYSMENU = &H80000
Const WS_THICKFRAME = &H40000
Const WS_MINIMIZEBOX = &H20000
Const WS_MAXIMIZEBOX = &H10000
Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Const WS_VISIBLE = &H10000000
Const WS_POPUP = &H80000000
Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Const WS_VSCROLL = &H200000
Const WS_EX_TOOLWINDOW = &H80
Const WS_EX_TOPMOST = &H8&
Const WS_EX_CLIENTEDGE = &H200&
Const WS_EX_WINDOWEDGE = &H100&
Const WS_SIZEBOX = &H40000
Public Const WS_EX_DLGMODALFRAME = &H1&
'-----------Window Messaging Constants
Const WM_DESTROY = &H2
Const WM_CLOSE = &H10
Const WM_LBUTTONDOWN = &H201
Const WM_LBUTTONUP = &H202
Const WM_CTLCOLOREDIT = &H133
Const WM_COMMAND = &H111
Const WM_GETTEXT = &HD
Const WM_ENABLE = &HA
Const WM_KEYDOWN = &H100
Const WM_KEYUP = &H101
Const WM_SETTEXT = &HC
Const WM_VSCROLL = &H115
Const WM_MOVE = &H3
Const WM_SIZE = &H5
Const WM_CHAR = &H102
Const WM_ACTIVATE = &H6
Const WM_SETFOCUS = &H7
Const WM_ACTIVATEAPP = &H1C

'--------Window Heiarchy Constants
Const GWL_WNDPROC = (-4)
Const GW_CHILD = 5
Const GW_OWNER = 4
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const SW_SHOWNORMAL = 1
'----------Misc Constants
Const CS_VREDRAW = &H1
Const CS_HREDRAW = &H2
Const CW_USEDEFAULT = &H80000000
Const COLOR_WINDOW = 5
Const SET_BACKGROUND_COLOR = 4103
Const IDC_ARROW = 32512&
Const IDI_APPLICATION = 32512&
Const MB_OK = &H0&
Const MB_ICONEXCLAMATION = &H30&

Dim MyMousePos As POINTAPI 'for getting the mouse positioning

Const gClassName = "Listbox API"

Dim gAppTitle As String


Dim gHwnd As Long

Dim gOKHwnd As Long
Dim gOKOldProc As Long
Dim gTextHwnd As Long
Dim gTextOldProc As Long
Dim gCancelHwnd As Long
Dim gCancelOldProc As Long

Dim PwdChar As String * 1
Dim PwdText As String
Dim TextStyle As String
Dim wTop As Long
Dim wLeft As Long
Dim wHeight As Long
Dim wWidth As Long
Dim Created As Boolean


 Function CreateWindows() As Boolean
    Dim i As Integer
    Dim tStr As String
    Dim ButtonStyle As Long
    ButtonStyle = WS_CHILD Or WS_VISIBLE Or WS_BORDER
    TextStyle = TextStyle Or WS_CHILD Or WS_VISIBLE Or WS_BORDER Or ES_CENTER
    'Create form window.
    gHwnd& = CreateWindowEx(WS_EX_TOOLWINDOW Or WS_EX_TOPMOST, gClassName$, gAppTitle$, WS_POPUPWINDOW Or WS_CAPTION Or WS_VISIBLE Or WS_BORDER, wLeft, wTop, wWidth, wHeight, 0&, 0&, App.hInstance, ByVal 0&)
    'Create Edit Box
    gTextHwnd& = CreateWindowEx(0&, "EDIT", "", TextStyle, 1, 1, wWidth - 9, wHeight - 47, gHwnd&, 0&, App.hInstance, 0&)
     'Create OK and Cancel Buttons
    gOKHwnd = CreateWindowEx(0&, "BUTTON", "OK", ButtonStyle, 1, wHeight - 44, (wWidth - 9) / 2, 20, gHwnd&, 0&, App.hInstance, 0&)
    gCancelHwnd = CreateWindowEx(0&, "BUTTON", "CANCEL", ButtonStyle, (wWidth - 4) / 2, wHeight - 44, (wWidth - 9) / 2, 20, gHwnd&, 0&, App.hInstance, 0&)
        
    '-------Hook OK CANCEL and TEXT-----------
    gOKOldProc& = GetWindowLong(gOKHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gOKHwnd&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
    gCancelOldProc& = GetWindowLong(gCancelHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gCancelHwnd&, GWL_WNDPROC, GetAddress(AddressOf CancelWndProc))
    gTextOldProc& = GetWindowLong(gTextHwnd&, GWL_WNDPROC)
    Call SetWindowLong(gTextHwnd&, GWL_WNDPROC, GetAddress(AddressOf TextWndProc))
    
    
    Call SendMessage(gHwnd, WM_SIZE, 0&, 0&)
    Call SendMessage(gTextHwnd, EM_SETPASSWORDCHAR, Asc(PwdChar), 0&)
    Call SetFocus(gTextHwnd)
    
    
    CreateWindows = (gHwnd& <> 0)
    Created = True
End Function
Function CancelWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN
            Call SendMessage(gHwnd, WM_CLOSE, 0&, 0&)
    End Select
    
  CancelWndProc = CallWindowProc(gCancelOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function
Sub Main()
Dim MyPwd As String
MyPwd = PwdBox("This is my password box.", True, "¤")
MsgBox MyPwd

End Sub

Sub MakeSelection()
Dim tLen As Long
Dim tItem As String
tLen = SendMessage(gTextHwnd&, EM_LINELENGTH, 0&, 0&)
tItem = Space(tLen)
Call SendMessage(gTextHwnd&, EM_GETLINE, 0&, ByVal tItem)
PwdText = tItem
End Sub

Function OKWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg&
        Case WM_LBUTTONDOWN
            MakeSelection
            Call SendMessage(gHwnd, WM_CLOSE, 0&, 0&)
    End Select
    
  OKWndProc = CallWindowProc(gOKOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function

Function TextWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg&
        Case WM_CHAR
            If wParam = 13 Then
                MakeSelection
                Call SendMessage(gHwnd, WM_CLOSE, 0&, 0&)
            End If
    End Select
    
  TextWndProc = CallWindowProc(gTextOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function

Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ''This our default window procedure for the window. It will handle all
    ''of our incoming window messages and we will write code based on the
    ''window message what the program should do.
    Dim i As Integer
      Select Case uMsg&
         Case WM_DESTROY:
            ''Since DefWindowProc doesn't automatically call
            ''PostQuitMessage (WM_QUIT). We need to do it ourselves.
            ''You can use DestroyWindow to get rid of the window manually.
            'SetDate
            Call PostQuitMessage(0&)
      End Select
    ''Let windows call the default window procedure since we're done.
    WndProc = defWindowProc(hwnd&, uMsg&, wParam&, lParam&)

End Function
 
Function RegisterWindowClass() As Boolean

    Dim wc As WNDCLASS
    
    ''Registers our new window with windows so we can use our classname.
    
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnwndproc = GetAddress(AddressOf WndProc) ''Address in memory of default window procedure.
    wc.hInstance = App.hInstance
    wc.hIcon = LoadIcon(0&, IDI_APPLICATION) ''Default application icon
    wc.hCursor = LoadCursor(0&, IDC_ARROW) ''Default arrow
    wc.hbrBackground = COLOR_WINDOW ''Default a color for window.
    wc.lpszClassName = gClassName$

    RegisterWindowClass = RegisterClass(wc) <> 0
    
End Function
 Function GetAddress(ByVal lngAddr As Long) As Long
    ''Used with AddressOf to return the address in memory of a procedure.

    GetAddress = lngAddr&
    
End Function

Function PwdBox(Optional Title As String = "Enter Password", Optional Mask As Boolean = True, Optional MaskChar As String = "*") As String
    Call GetCursorPos(MyMousePos)
    wLeft = MyMousePos.x
    wTop = MyMousePos.y
    wWidth = 300
    wHeight = 70
    If Title <> "" Then gAppTitle$ = Title Else gAppTitle$ = "Enter Password"
    Dim wMsg As Msg
    Dim tSec As String
    If Mask Then
        TextStyle = ES_PASSWORD
        PwdChar = MaskChar
    Else
        PwdChar = ""
    End If
    ''Call procedure to register window classname. If false, then exit.
    If RegisterWindowClass = False Then Exit Function
    
      ''Create window
      If CreateWindows() Then
         ''Loop will exit when WM_QUIT is sent to the window.
         Do While GetMessage(wMsg, 0&, 0&, 0&)
            ''TranslateMessage takes keyboard messages and converts
            ''them to WM_CHAR for easier processing.
            Call TranslateMessage(wMsg)
            ''Dispatchmessage calls the default window procedure
            ''to process the window message. (WndProc)
            Call DispatchMessage(wMsg)
            DoEvents
         Loop
      End If
    
    Call UnregisterClass(gClassName$, App.hInstance)
    PwdBox = PwdText
End Function
