Attribute VB_Name = "Module1"
'For 32 bit & 64 bit
Option Explicit

'WinMain
Declare PtrSafe Function GetModuleHandleW Lib "Kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
Declare PtrSafe Function LoadIconW Lib "User32" (ByVal hInstance As LongPtr, ByVal lpIconName As LongPtr) As LongPtr
Declare PtrSafe Function LoadCursorW Lib "User32" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Declare PtrSafe Function GetStockObject Lib "Gdi32" (ByVal i As LongPtr) As LongPtr
Declare PtrSafe Function RegisterClassExW Lib "User32" (ByVal lpWndClass As LongPtr) As LongPtr
Declare PtrSafe Function CreateWindowExW Lib "User32" (ByVal dwExStyle As LongPtr, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal nWidth As LongPtr, ByVal nHeight As LongPtr, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
Declare PtrSafe Function ShowWindow Lib "User32" (ByVal hWnd As LongPtr, ByVal nCmdShow As LongPtr) As LongPtr
Declare PtrSafe Function UpdateWindow Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function PeekMessageW Lib "User32" (ByVal lpMsg As LongPtr, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As LongPtr, ByVal wMsgFilterMax As LongPtr, ByVal wRemoveMsg As LongPtr) As LongPtr
Declare PtrSafe Function TranslateMessage Lib "User32" (ByVal lpMsg As LongPtr) As LongPtr
Declare PtrSafe Function DispatchMessageW Lib "User32" (ByVal lpMsg As LongPtr) As LongPtr
Declare PtrSafe Function SendMessageW Lib "User32" (ByVal hWnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Declare PtrSafe Function GetLastError Lib "Kernel32" () As LongPtr
'WndProc
Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As LongPtr) As LongPtr
Declare PtrSafe Function GetCursorPos Lib "User32" (ByVal lpPoint As LongPtr) As LongPtr
Declare PtrSafe Function SetCursorPos Lib "User32" (ByVal X As LongPtr, ByVal Y As LongPtr) As LongPtr
Declare PtrSafe Function DefWindowProcW Lib "User32" (ByVal hWnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Declare PtrSafe Function DestroyWindow Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function PostQuitMessage Lib "User32" (ByVal nExitCode As LongPtr) As LongPtr
'StatusBar
Declare PtrSafe Function InitCommonControlsEx Lib "Comctl32" (ByVal picce As LongPtr) As LongPtr
'PopupMenu
Declare PtrSafe Function CreateMenu Lib "User32" () As LongPtr
Declare PtrSafe Function CreatePopupMenu Lib "User32" () As LongPtr
Declare PtrSafe Function SetMenu Lib "User32" (ByVal hWnd As LongPtr, ByVal hMenu As LongPtr) As LongPtr
Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function AppendMenuW Lib "User32" (ByVal hMenu As LongPtr, ByVal uFlags As LongPtr, ByVal uIDNewItem As LongPtr, ByVal lpNewItem As LongPtr) As LongPtr
Declare PtrSafe Function CheckMenuItem Lib "User32" (ByVal hMenu As LongPtr, ByVal uIDCheckItem As LongPtr, ByVal uCheck As LongPtr) As LongPtr
'Declare PtrSafe Function DestroyMenu Lib "User32" (ByVal hMenu As LongPtr) As LongPtr
'GDI & Paint
Declare PtrSafe Function GetDC Lib "User32" (ByVal hdc As LongPtr) As LongPtr
Declare PtrSafe Function ReleaseDC Lib "User32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As LongPtr
Declare PtrSafe Function DeleteDC Lib "Gdi32" (ByVal hdc As LongPtr) As LongPtr
Declare PtrSafe Function GetClientRect Lib "User32" (ByVal hWnd As LongPtr, ByVal lpRect As LongPtr) As LongPtr
Declare PtrSafe Function InvalidateRect Lib "User32" (ByVal hWnd As LongPtr, ByVal lpRect As LongPtr, ByVal bErase As LongPtr) As LongPtr
Declare PtrSafe Function ChoosePixelFormat Lib "Gdi32" (ByVal hdc As LongPtr, ByVal ppfd As LongPtr) As LongPtr
Declare PtrSafe Function SetPixelFormat Lib "Gdi32" (ByVal hdc As LongPtr, ByVal format As LongPtr, ByVal ppfd As LongPtr) As LongPtr
Declare PtrSafe Function SwapBuffers Lib "Gdi32" (ByVal hdc As LongPtr) As LongPtr
'OpenGL
Declare PtrSafe Function wglCreateContext Lib "Opengl32" (ByVal hdc As LongPtr) As LongPtr
Declare PtrSafe Function wglMakeCurrent Lib "Opengl32" (ByVal hdc As LongPtr, ByVal hGLRC As LongPtr) As LongPtr
Declare PtrSafe Function wglDeleteContext Lib "Opengl32" (ByVal hGLRC As LongPtr) As LongPtr
Declare PtrSafe Function glViewport Lib "Opengl32" (ByVal X As LongPtr, ByVal Y As LongPtr, ByVal Width As LongPtr, ByVal height As LongPtr) As LongPtr
Declare PtrSafe Function glMatrixMode Lib "Opengl32" (ByVal mode As LongPtr) As LongPtr
Declare PtrSafe Function glLoadIdentity Lib "Opengl32" () As LongPtr
Declare PtrSafe Function glEnable Lib "Opengl32" (ByVal cap As LongPtr) As LongPtr
Declare PtrSafe Function glHint Lib "Opengl32" (ByVal target As LongPtr, ByVal mode As LongPtr) As LongPtr
Declare PtrSafe Function glClearColor Lib "Opengl32" (ByVal red As Single, ByVal green As Single, ByVal blue As Single, ByVal alpha As Single) As LongPtr
Declare PtrSafe Function glClear Lib "Opengl32" (ByVal mask As LongPtr) As LongPtr
Declare PtrSafe Function glBegin Lib "Opengl32" (ByVal mode As LongPtr) As LongPtr
Declare PtrSafe Function glEnd Lib "Opengl32" () As LongPtr
Declare PtrSafe Function glPushMatrix Lib "Opengl32" () As LongPtr
Declare PtrSafe Function glPopMatrix Lib "Opengl32" () As LongPtr
Declare PtrSafe Function glTranslatef Lib "Opengl32" (ByVal X As Single, ByVal Y As Single, ByVal z As Single) As LongPtr
Declare PtrSafe Function glRotatef Lib "Opengl32" (ByVal Angle As Single, ByVal X As Single, ByVal Y As Single, ByVal z As Single) As LongPtr
Declare PtrSafe Function glScalef Lib "Opengl32" (ByVal X As Single, ByVal Y As Single, ByVal z As Single) As LongPtr
Declare PtrSafe Function glColor3f Lib "Opengl32" (ByVal red As Single, ByVal green As Single, ByVal blue As Single) As LongPtr
Declare PtrSafe Function glVertex3f Lib "Opengl32" (ByVal X As Single, ByVal Y As Single, ByVal z As Single) As LongPtr
Declare PtrSafe Function gluPerspective Lib "Glu32" (ByVal fovy As Double, ByVal aspect As Double, ByVal zNear As Double, ByVal zFar As Double) As LongPtr

'Structures
'WNDCLASSEX wcx;
Type WNDCLASSEX
    cbSize As Long
    style As Long
    lpfnWndProc As LongPtr
    cbClsExtra As Long
    cbWndExtra As Long
    hInstance As LongPtr
    hIcon As LongPtr
    hCursor As LongPtr
    hbrBackground As LongPtr
    lpszMenuName As LongPtr
    lpszClassName As LongPtr
    hIconSm As LongPtr
End Type
Public wcx As WNDCLASSEX
Public lpWndClass As LongPtr 'Pointer
'POINT pt
Type POINT2D
    X As Long
    Y As Long
End Type
Public pt As POINT2D
Public lpMsgPt As LongPtr
'MSG msg;
Type Msg
    hWnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    pt As POINT2D
    lPrivate As LongPtr
End Type
Public wmsg As Msg
Public lpMsg As LongPtr 'Pointer
'RECT rct
Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type
Public RectMain As RECT
Public lpRectMain As LongPtr 'Pointer
Public RectWidth As Long
Public RectHeight As Long
Public RectAspect As Long
'PIXELFORMATDESCRIPTOR pfd = { 0 };
Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlphaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type
Public pfd As PIXELFORMATDESCRIPTOR
Public lpPFD As LongPtr 'Pointer
'INITCOMMONCONTROLSEX icce;
Type ICCESTRUCT
    dwSize As Long
    dwICC As Long
End Type
Public icce As ICCESTRUCT
Public lpIcce As LongPtr 'Pointer
'Const
Public Const PM_REMOVE = 1
'Window Styles
Public Const CS_VREDRAW = 1
Public Const CS_HREDRAW = 2
Public Const CS_DBLCLKS = 8
Public Const IDI_APPLICATION = 32512
Public Const IDC_ARROW = 32512
Public Const CW_USEDEFAULT = &H80000000
Public Const WS_OVERLAPPED = 0
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const SW_SHOW = 5
'Messages
Public Const WM_CREATE = 1
Public Const WM_DESTROY = 2
Public Const WM_SIZE = 5
Public Const WM_CLOSE = &H10
Public Const WM_QUIT = &H12
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_COMMAND = &H111
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_USER = &H400
'Virtual Keys
Public Const MB_OK = 0
Public Const MB_YESNO = 4
Public Const MB_ICONQUESTION = &H20
Public Const MB_ICONINFORMATION = &H40
Public Const IDYES = 6
Public Const IDNO = 7
Public Const VK_RETURN = &HD
Public Const VK_SHIFT = &H10
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
'Brushes
Public Const GRAY_BRUSH = 2
'Status Bar
Public Const ICC_BAR_CLASSES = 4
Public Const SBARS_SIZEGRIP = &H100
Public Const SBT_NOBORDERS = &H101 '0x0100
Public Const SB_SETPARTS = WM_USER + 4 '0x0404
Public Const SB_SETTEXTW = WM_USER + 11 '0x040B
'Menu
Public Const TPM_LEFTALIGN = 0
Public Const TPM_TOPALIGN = 0
Public Const TPM_RIGHTBUTTON = 2
'https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-insertmenua
Public Const MF_STRING = 0
Public Const MF_SEPARATOR = &H800
Public Const MF_ENABLED = 0
Public Const MF_GRAYED = 1
Public Const MF_DISABLED = 2
Public Const MF_UNCHECKED = 0
Public Const MF_CHECKED = 8
Public Const MF_POPUP = &H10
'OpenGL
Public Const PFD_SUPPORT_OPENGL = &H20
Public Const PFD_DOUBLEBUFFER = 1
Public Const PFD_DRAW_TO_WINDOW = 4
Public Const PFD_TYPE_RGBA = 0
Public Const PFD_MAIN_PLANE = 0
Public Const GL_MODELVIEW = &H1700
Public Const GL_PROJECTION = &H1701
Public Const GL_DEPTH_TEST = &HB71
Public Const GL_COLOR_BUFFER_BIT = &H4000
Public Const GL_DEPTH_BUFFER_BIT = &H100
Public Const GL_NICEST = &H1102
Public Const GL_PERSPECTIVE_CORRECTION_HINT = &HC50
Public Const GL_LINES = 1
Public Const GL_TRIANGLES = 4
Public Const GL_TRIANGLE_STRIP = 5
Public Const GL_QUADS = 7
Public Const GL_QUAD_STRIP = 8

'Custom Values
'Numeric
Public Const PI_OVER_180 = 1.74532925199433E-02
'Screen
Public Const DEFAULT_SCREEN_WIDTH = 1024
Public Const DEFAULT_SCREEN_HEIGHT = 768
'Menu ID's
Public Const IDM_APP_EXIT = 1009
Public Const IDM_HELP_ABOUT = 9001

'Global Handles
Public ghInst As LongPtr
Public ghWnd As LongPtr
'Status Bar
Public hwndStatusBar As LongPtr
Public idStatusBar As LongPtr
Public szStatusClassName As String
Public lpszStatusClassName As LongPtr
Public xStatusParts(7) As Long 'Divide Status Bar by 8 parts
Public lpStatusParts As LongPtr
'Menu Handles
Public hMenu As LongPtr
Public hMenuFile As LongPtr
Public hMenuHelp As LongPtr
'Menu Text Strings
Public szMenuFile As String
Public szMenuFileExit As String
Public szMenuHelp As String
Public szMenuHelpAbout As String
'Menu Text Pointers
Public lpszMenuFile As LongPtr
Public lpszMenuFileExit As LongPtr
Public lpszMenuHelp As LongPtr
Public lpszMenuHelpAbout As LongPtr
'OpenGL
Public ghDC As LongPtr
Public ghRC As LongPtr
Public iPixelFormat As LongPtr
'Model Scale
Public GlobalScale As Single
'Model Angles
Public aXY_Model As Single 'Model Turn
'Camera Position
Public xCam, yCam, zCam As Double
'Camera Angles
Public aYZ_Cam As Single '1 - Side - Camera Tilt
Public aXY_Cam As Single '2 - Plan - Camera Turn
Public aXZ_Cam As Single '3 - Front - Camera Roll
'Motion
Public LinearSpeed As Single
Public LinearBoost As Single
Public dStep As Single
Public AngularSpeed As Single
Public AngularBoost As Single
Public dAngle As Single
Public aRad, cosA, sinA As Single
Public dxCam0, dyCam0, dzCam0  As Single
Public dxCam1, dyCam1, dzCam1  As Single
Public dxCam2, dyCam2, dzCam2  As Single
Public dxCam3, dyCam3, dzCam3  As Single

'Keyboard Buffer
Public key(127) As Byte
Public nKeyCode As Integer

'Name Strings
Public szCaption As String
Public szMsgText As String
Public szMsgTitle As String
'Name Pointers
Public lpszCaption As LongPtr
Public lpszMsgText As LongPtr
Public lpszMsgTitle As LongPtr
'Status Strings
Public sz_xCam As String
Public sz_yCam As String
Public sz_zCam As String
Public sz_aXY_Model As String
Public sz_aYZ_Cam As String
Public sz_aXY_Cam As String
Public sz_aXZ_Cam As String
'Status Pointers
Public lpsz_xCam As LongPtr
Public lpsz_yCam As LongPtr
Public lpsz_zCam As LongPtr
Public lpsz_aXY_Model As LongPtr
Public lpsz_aYZ_Cam As LongPtr
Public lpsz_aXY_Cam As LongPtr
Public lpsz_aXZ_Cam As LongPtr

'Debug
Public nLastError As LongPtr

Sub Start()
    Call WinMain(0, 0, 0, 0)
End Sub

Function WinMain(ByVal hInstance As LongPtr, ByVal hPrevInstance As LongPtr, ByVal lpCmdLine As LongPtr, ByVal nCmdShow As LongPtr)
'Main Cycle
    Dim nPeek As LongPtr
'Debug
    Dim nWndClass As LongPtr 'For Debug Purpose

'Pointers
    lpWndClass = VarPtr(wcx.cbSize)
    lpMsg = VarPtr(wmsg.hWnd)

    wcx.cbSize = Len(wcx) '30h bytes
    wcx.style = CS_HREDRAW Or CS_VREDRAW Or CS_DBLCLKS
    wcx.lpfnWndProc = GetAddr(AddressOf WndProc)
    wcx.cbClsExtra = 0
    wcx.cbWndExtra = 0
    wcx.hInstance = GetModuleHandleW(0)
    If wcx.hInstance = 0 Then
        Call MsgBox("GetModuleHandle Error: " & GetLastError())
        Exit Function
    Else
        ghInst = wcx.hInstance
    End If
    wcx.hIcon = LoadIconW(0, IDI_APPLICATION)
    wcx.hCursor = LoadCursorW(0, IDC_ARROW)
    wcx.hbrBackground = GetStockObject(GRAY_BRUSH)
    wcx.lpszMenuName = 0
    wcx.lpszClassName = StrPtr("MainWindowClassName")
    wcx.hIconSm = wcx.hIcon

    nWndClass = RegisterClassExW(lpWndClass)
    If nWndClass = 0 Then
        nLastError = GetLastError()
        If nLastError = 1410 Then
            'Call MsgBox("RegisterClass Error 1410 (0x582): ERROR_CLASS_ALREADY_EXISTS")
        Else
            Call MsgBox("RegisterClass Error: " & nLastError)
            Exit Function
        End If
    End If

    szCaption = "VBA WinAPI OpenGL Demo"
    lpszCaption = StrPtr(szCaption)
    ghWnd = CreateWindowExW(0, wcx.lpszClassName, lpszCaption, WS_OVERLAPPEDWINDOW Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, CW_USEDEFAULT, CW_USEDEFAULT, DEFAULT_SCREEN_WIDTH, DEFAULT_SCREEN_HEIGHT, 0, 0, wcx.hInstance, 0)
    If ghWnd = 0 Then
        Call MsgBox("CreateWindow Error: " & GetLastError())
        Exit Function
    End If

    Call InitializeGL

    Call ShowWindow(ghWnd, SW_SHOW)
    Call UpdateWindow(ghWnd)

WinMainLoop:
    nPeek = PeekMessageW(lpMsg, 0, 0, 0, PM_REMOVE)
    If nPeek = 0 Then
        Call DrawGLScene 'We'll draw the OpenGL Scene outside the WndProc
    Else
        If wmsg.message = WM_QUIT Then
            WinMain = wmsg.wParam
            Exit Function 'Don't use Call ExitProcess(0)
        Else
            Call TranslateMessage(lpMsg)
            Call DispatchMessageW(lpMsg)
        End If
    End If
    GoTo WinMainLoop
End Function

Function WndProc(ByVal hWnd As LongPtr, ByVal message As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'Pointers
    lpRectMain = VarPtr(RectMain.left)

Select Case message
    Case WM_KEYDOWN
        GoTo lbl_wmKeyDown
    Case WM_KEYUP
        GoTo lbl_wmKeyUp
    Case WM_LBUTTONDOWN
        GoTo lbl_wmLButtonDown
    Case WM_COMMAND
        GoTo lbl_wmCommand
    Case WM_SIZE
        GoTo lbl_wmSize
    Case WM_CLOSE
        GoTo lbl_wmClose
    Case WM_DESTROY
        GoTo lbl_wmDestroy
    Case WM_CREATE
        GoTo lbl_wmCreate
End Select

lbl_DefWndProc:
    WndProc = DefWindowProcW(hWnd, message, wParam, lParam)
    Exit Function
    
lbl_wmCreate:
    Call InitCommonControlsEx(lpIcce)
    Call DoCreateMenu(hWnd)
    Call DoCreateStatusBar(hWnd)
    GoTo lbl_WndProc_Return0

lbl_wmSize:
    Call GLResize(hWnd)
    GoTo lbl_WndProc_Return0

lbl_wmKeyDown:
    nKeyCode = CInt(wParam And 32767) 'Low Word, Signed Integer
    key(nKeyCode) = 1
    Select Case nKeyCode 'For single (not continuous) KeyStrokes
        Case 9 'Tab
            aXY_Model = aXY_Model + 30 'Object Turn Counter-Clockwise 30 degrees
        Case &HD 'Enter
            Call AboutProc(hWnd)
        Case &H20 'SpaceBar
            Call ResetScene
        Case &H1B 'Esc
            Call CloseWndProc(hWnd)
        End Select
    GoTo lbl_WndProc_Return0

lbl_wmKeyUp:
    nKeyCode = CInt(wParam And 32767) 'Low Word, Signed Integer
    key(nKeyCode) = 0
    GoTo lbl_WndProc_Return0

lbl_wmLButtonDown:
    Call AboutProc(hWnd)
    GoTo lbl_WndProc_Return0

lbl_wmCommand:
    Select Case wParam
        Case IDM_APP_EXIT
            Call CloseWndProc(hWnd)
        Case IDM_HELP_ABOUT
            Call AboutProc(hWnd)
    End Select
    GoTo lbl_WndProc_Return0

lbl_wmClose:
    Call CloseWndProc(hWnd)
    GoTo lbl_WndProc_Return0

lbl_wmDestroy:
    Call PostQuitMessage(0)
lbl_WndProc_Return0:
    WndProc = 0
End Function

Function GetAddr(ByVal lpProc As LongPtr) As LongPtr
    'This Function has been created to fit the 'AddressOf' syntax
    GetAddr = lpProc
End Function

Function DoCreateMenu(ByVal hWnd As LongPtr)
'Menu Text  Strings
    szMenuFile = "&File"
    szMenuFileExit = "E&xit" + vbTab + "Ctrl+W"
    szMenuHelp = "&Help"
    szMenuHelpAbout = "&About..."
'Menu Text Pointers
    lpszMenuFile = StrPtr(szMenuFile)
    lpszMenuFileExit = StrPtr(szMenuFileExit)
    lpszMenuHelp = StrPtr(szMenuHelp)
    lpszMenuHelpAbout = StrPtr(szMenuHelpAbout)

'Main Menu
    hMenu = CreateMenu()
    hMenuFile = CreatePopupMenu()
    Call AppendMenuW(hMenu, MF_POPUP, hMenuFile, lpszMenuFile)
        Call AppendMenuW(hMenuFile, MF_STRING, IDM_APP_EXIT, lpszMenuFileExit)
    hMenuHelp = CreatePopupMenu()
    Call AppendMenuW(hMenu, MF_POPUP, hMenuHelp, lpszMenuHelp)
        Call AppendMenuW(hMenuHelp, MF_STRING, IDM_HELP_ABOUT, lpszMenuHelpAbout)
    Call SetMenu(hWnd, hMenu)
    Call DrawMenuBar(hWnd)
End Function

Function DoCreateStatusBar(ByVal hWnd As LongPtr)
    'https://learn.microsoft.com/en-us/windows/win32/winauto/status-bar-control
    szStatusClassName = "msctls_statusbar32" '"STATUSCLASSNAMEW"
    lpszStatusClassName = StrPtr(szStatusClassName)
    idStatusBar = 1 'Child window identifier for Status Bar
    hwndStatusBar = CreateWindowExW(0, lpszStatusClassName, 0, SBARS_SIZEGRIP Or WS_CHILD Or WS_VISIBLE, 0, 0, 0, 0, hWnd, idStatusBar, ghInst, 0)
    If hwndStatusBar = 0 Then
        Call MsgBox("Status Bar Error: " & GetLastError())
        Exit Function
    End If
    xStatusParts(0) = 100
    xStatusParts(1) = 200
    xStatusParts(2) = 300
    xStatusParts(3) = 450
    xStatusParts(4) = 600
    xStatusParts(5) = 750
    xStatusParts(6) = 900
    xStatusParts(7) = -1
    lpStatusParts = VarPtr(xStatusParts(0))
    Call SendMessageW(hwndStatusBar, SB_SETPARTS, 8, lpStatusParts)
End Function

Function CloseWndProc(ByVal hWnd As LongPtr)
    szMsgText = "Close?"
    lpszMsgText = StrPtr(szMsgText)
    szMsgTitle = "Such A Good Application"
    lpszMsgTitle = StrPtr(szMsgTitle)
    If MessageBoxW(hWnd, lpszMsgText, lpszMsgTitle, MB_YESNO Or MB_ICONQUESTION) = IDYES Then
        Call wglMakeCurrent(0, 0)
        Call wglDeleteContext(ghRC)
        Call ReleaseDC(hWnd, ghDC)
        Call DestroyWindow(hWnd)
    End If
End Function

Function AboutProc(ByVal hWnd As LongPtr)
    szMsgText = _
    "Camera Motion:" & Chr(13) & _
    "Arrow Up - Move Forward" & Chr(13) & _
    "Arrow Down - Move Backward" & Chr(13) & _
    "Arrow Left - Move Left" & Chr(13) & _
    "Arrow Right - Move Right" & Chr(13) & _
    "Page Up - Move Up" & Chr(13) & _
    "Page Down - Move Down" & Chr(13) & Chr(13) & _
    "Camera Rotation:" & Chr(13) & _
    "W - Look Down" & Chr(13) & _
    "S - Look Up" & Chr(13) & _
    "A - Look Left" & Chr(13) & _
    "D - Look Right" & Chr(13) & Chr(13) & _
    "Q - Roll the Camera Clockwise" & Chr(13) & _
    "E - Roll the Camera Counter-Clockwise" & Chr(13) & _
    "Object Rotation:" & Chr(13) & _
    "Z - Turn the Object Clockwise" & Chr(13) & _
    "C - Turn the Object Counter-Clockwise" & Chr(13) & _
    "Tab - Turn the Object Clockwise Quick" & Chr(13) & Chr(13) & _
    "Shift - Boost"
    lpszMsgText = StrPtr(szMsgText)
    szMsgTitle = "Manual"
    lpszMsgTitle = StrPtr(szMsgTitle)
    Call MessageBoxW(hWnd, lpszMsgText, lpszMsgTitle, MB_OK Or MB_ICONINFORMATION)
End Function

Function InitializeGL()
    Dim i As Integer
    For i = 0 To 127
        key(i) = 0
    Next i
'PIXELFORMATDESCRIPTOR pfd
    pfd.nSize = Len(pfd) 'sizeof( PIXELFORMATDESCRIPTOR )
    pfd.nVersion = 1 'always 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER Or PFD_DRAW_TO_WINDOW
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24
    pfd.cRedBits = 0
    pfd.cRedShift = 0
    pfd.cGreenBits = 0
    pfd.cGreenShift = 0
    pfd.cBlueBits = 0
    pfd.cBlueShift = 0
    pfd.cAlphaBits = 0
    pfd.cAlphaShift = 0
    pfd.cAccumBits = 0
    pfd.cAccumRedBits = 0
    pfd.cAccumGreenBits = 0
    pfd.cAccumBlueBits = 0
    pfd.cAccumAlphaBits = 0
    pfd.cDepthBits = 32
    pfd.cStencilBits = 0
    pfd.cAuxBuffers = 0
    pfd.iLayerType = PFD_MAIN_PLANE
    pfd.bReserved = 0
    pfd.dwLayerMask = 0
    pfd.dwVisibleMask = 0
    pfd.dwDamageMask = 0
    lpPFD = VarPtr(pfd.nSize)

    ghDC = GetDC(ghWnd)

    iPixelFormat = ChoosePixelFormat(ghDC, lpPFD)
    If iPixelFormat = 0 Then
        MsgBox ("ChoosePixelFormat Error: " & GetLastError())
    End If

    Call SetPixelFormat(ghDC, iPixelFormat, lpPFD)

    ghRC = wglCreateContext(ghDC)
    If ghRC = 0 Then
        nLastError = GetLastError()
        If nLastError = 2000 Then
            MsgBox ("wglCreateContext Error 2000 (0x7D0): ERROR_INVALID_PIXEL_FORMAT")
        Else
            MsgBox ("wglCreateContext Error: " & nLastError)
        End If
    End If

    Call wglMakeCurrent(ghDC, ghRC)
    
    Call glEnable(GL_DEPTH_TEST)
    Call glHint(GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST)

    Call ResetScene
End Function

Function GLResize(ByVal hWnd As LongPtr)
    Call GetClientRect(hWnd, lpRectMain)
    RectWidth = RectMain.right - RectMain.left
    RectHeight = RectMain.bottom - RectMain.top
    If RectHeight > 0 Then
        RectAspect = RectWidth / RectHeight
    End If
    'Main Viewport
    Call glViewport(0, 0, RectWidth, RectHeight)
    'Status Bar
    xStatusParts(6) = RectWidth - 50
    Call SendMessageW(hwndStatusBar, WM_SIZE, 0, 0)
End Function

Function DrawGLScene()
    Call glClear(GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT)
        Call CheckKeys
        Call RefreshStatus
        Call SetScene
        Call DrawAxes
        Call DrawObject
    Call SwapBuffers(ghDC)
End Function

Function CheckKeys() 'For Continuous KeyStrokes
'Return to Normal Speed
    dStep = LinearSpeed
    dAngle = AngularSpeed
'Boost
    If key(&H10) <> 0 Then  'Shift
        dStep = dStep * LinearBoost
        dAngle = dAngle * AngularBoost
    End If
'Object Rotation
    If key(&H5A) <> 0 Then  'Z - Object Turn Counter-Clockwise
        aXY_Model = aXY_Model + dAngle
        aXY_Model = CheckAngle(aXY_Model)
    End If
    If key(&H43) <> 0 Then  'C - Object Turn Clockwise
        aXY_Model = aXY_Model - dAngle
        aXY_Model = CheckAngle(aXY_Model)
    End If
'CameraRotation
    If key(&H57) <> 0 Then  'W - Camera Tilt Down
        aYZ_Cam = aYZ_Cam - dAngle
        aYZ_Cam = CheckAngle(aYZ_Cam)
    End If
    If key(&H53) <> 0 Then  'S - Camera Tilt Up
        aYZ_Cam = aYZ_Cam + dAngle
        aYZ_Cam = CheckAngle(aYZ_Cam)
    End If
    If key(&H41) <> 0 Then  'A - Camera Turn Clockwise
        aXY_Cam = aXY_Cam - dAngle
        aXY_Cam = CheckAngle(aXY_Cam)
    End If
    If key(&H44) <> 0 Then  'D - Camera Turn Counter-Clockwise
        aXY_Cam = aXY_Cam + dAngle
        aXY_Cam = CheckAngle(aXY_Cam)
    End If
    If key(&H51) <> 0 Then  'Q - Camera Roll Clockwise
        aXZ_Cam = aXZ_Cam + dAngle
        aXZ_Cam = CheckAngle(aXZ_Cam)
    End If
    If key(&H45) <> 0 Then  'E - Camera Roll Counter-Clockwise
        aXZ_Cam = aXZ_Cam - dAngle
        aXZ_Cam = CheckAngle(aXZ_Cam)
    End If
'Camera Move Forward and Backward
    If key(&H26) <> 0 Then 'Up Arrow
        dxCam0 = 0
        dyCam0 = 0
        dzCam0 = dStep
        Call CameraMove
    End If
    If key(&H28) <> 0 Then 'Down Arrow
        dxCam0 = 0
        dyCam0 = 0
        dzCam0 = -dStep
        Call CameraMove
    End If
'Camera Move Left and Right
    If key(&H25) <> 0 Then 'Left Arrow
        dxCam0 = dStep
        dyCam0 = 0
        dzCam0 = 0
        Call CameraMove
    End If
    If key(&H27) <> 0 Then 'Right Arrow
        dxCam0 = -dStep
        dyCam0 = 0
        dzCam0 = 0
        Call CameraMove
    End If
'Camera Move Up and Down
    If key(&H21) <> 0 Then 'Page Up
        dxCam0 = 0
        dyCam0 = -dStep
        dzCam0 = 0
        Call CameraMove
    End If
    If key(&H22) <> 0 Then 'Page Down
        dxCam0 = 0
        dyCam0 = dStep
        dzCam0 = 0
        Call CameraMove
    End If
End Function

Function CheckAngle(ByVal Angle As Single) As Single
    If Angle > 360 Then
        CheckAngle = Angle - 360
    ElseIf Angle < 0 Then
        CheckAngle = Angle + 360
    Else
        CheckAngle = Angle
    End If
End Function

Function CheckDistance(ByVal Distance As Single) As Single
    If Distance > 20 Then
        CheckDistance = 20
    ElseIf Distance < -20 Then
        CheckDistance = -20
    Else
        CheckDistance = Distance
    End If
End Function

Function CameraMove()
    'RotX - UnTilt
    aRad = aYZ_Cam * PI_OVER_180
    cosA = Cos(aRad)
    sinA = Sin(aRad)
    dxCam1 = dxCam0
    dyCam1 = dyCam0 * cosA + dzCam0 * sinA
    dzCam1 = dzCam0 * cosA - dyCam0 * sinA
    'RotZ - UnTurn
    aRad = aXY_Cam * PI_OVER_180
    cosA = Cos(aRad)
    sinA = Sin(aRad)
    dxCam2 = dxCam1 * cosA + dyCam1 * sinA
    dyCam2 = dyCam1 * cosA - dxCam1 * sinA
    dzCam2 = dzCam1
    'RotY - UnRoll
    aRad = aXZ_Cam * PI_OVER_180
    cosA = Cos(aRad)
    sinA = Sin(aRad)
    dxCam3 = dxCam2 * cosA - dzCam2 * sinA
    dyCam3 = dyCam2
    dzCam3 = dxCam2 * sinA + dzCam2 * cosA
    'Move
    xCam = xCam + dxCam3
    yCam = yCam + dyCam3
    zCam = zCam + dzCam3
    'Check
    xCam = CheckDistance(xCam)
    yCam = CheckDistance(yCam)
    zCam = CheckDistance(zCam)
End Function

Function RefreshStatus()
    sz_xCam = "xCam = " & format(xCam, "0.000")
    sz_yCam = "yCam = " & format(yCam, "0.000")
    sz_zCam = "zCam = " & format(zCam, "0.000")
    'sz_xCam = "dxCam3 = " & format(dxCam3 * 1000, "0.000")
    'sz_yCam = "dyCam3 = " & format(dyCam3 * 1000, "0.000")
    'sz_zCam = "dzCam3 = " & format(dzCam3 * 1000, "0.000")
    sz_aXY_Model = "aXY_Model = " & format(aXY_Model, "0.000")
    sz_aYZ_Cam = "aYZ_Cam = " & format(aYZ_Cam, "0.000")
    sz_aXY_Cam = "aXY_Cam = " & format(aXY_Cam, "0.000")
    sz_aXZ_Cam = "aXZ_Cam = " & format(aXZ_Cam, "0.000")
    
    lpsz_xCam = StrPtr(sz_xCam)
    lpsz_yCam = StrPtr(sz_yCam)
    lpsz_zCam = StrPtr(sz_zCam)
    lpsz_aXY_Model = StrPtr(sz_aXY_Model)
    lpsz_aYZ_Cam = StrPtr(sz_aYZ_Cam)
    lpsz_aXY_Cam = StrPtr(sz_aXY_Cam)
    lpsz_aXZ_Cam = StrPtr(sz_aXZ_Cam)
    
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 0, lpsz_xCam)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 1, lpsz_yCam)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 2, lpsz_zCam)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 3, lpsz_aXY_Model)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 4, lpsz_aYZ_Cam)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 5, lpsz_aXY_Cam)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 6, lpsz_aXZ_Cam)
End Function

Function SetScene()
    'Set Object
    Call glMatrixMode(GL_MODELVIEW)
    Call glLoadIdentity
    Call glScalef(GlobalScale, GlobalScale, GlobalScale)
    Call glRotatef(aXY_Model, 0, 0, 1)
    'Set Camera
    Call glMatrixMode(GL_PROJECTION)
    Call glLoadIdentity
    Call gluPerspective(90, RectAspect, 0.1, 1000)
    Call glRotatef(aYZ_Cam, 1, 0, 0) '1 - Camera Tilt
    Call glRotatef(aXY_Cam, 0, 0, 1) '2 - Camera Turn
    Call glRotatef(aXZ_Cam, 0, 1, 0) '3 - Camera Roll
    Call glTranslatef(xCam, yCam, zCam)
End Function

Function ResetScene()
    GlobalScale = 0.001
    If Len(ghInst) = 4 Then
        LinearSpeed = 0.005 '32-bit Systems are Faster
        AngularSpeed = 0.1
    Else
        LinearSpeed = 0.05 '64-bit Systems are Slower
        AngularSpeed = 1
    End If
    LinearBoost = 10
    AngularBoost = 10
    dStep = LinearSpeed
    dAngle = AngularSpeed
    aXY_Model = 20
    aYZ_Cam = 300 'Cameta Tilt
    aXY_Cam = 0  'Camera Turn
    aXZ_Cam = 0 'Camera Roll
    xCam = 0 'Negative World Coordinates
    yCam = 9 'Negative World Coordinates
    zCam = -4.5 'Negative World Coordinates
    dxCam3 = 0
    dyCam3 = 0
    dzCam3 = 0
End Function

Function DrawAxes()
    Call glColor3f(1, 0, 0) 'Red
    Call glBegin(GL_LINES)
        Call glVertex3f(-7000, 0, 0)
        Call glVertex3f(7000, 0, 0)
        Call glVertex3f(6900, -50, 0) 'Arrow
        Call glVertex3f(7000, 0, 0)
        Call glVertex3f(7000, 0, 0)
        Call glVertex3f(6900, 50, 0)
    Call glEnd

    Call glColor3f(0, 1, 0) 'Green
    Call glBegin(GL_LINES)
        Call glVertex3f(0, -7000, 0)
        Call glVertex3f(0, 7000, 0)
        Call glVertex3f(-50, 6900, 0) 'Arrow
        Call glVertex3f(0, 7000, 0)
        Call glVertex3f(0, 7000, 0)
        Call glVertex3f(50, 6900, 0)
    Call glEnd

    Call glColor3f(0, 0, 1) 'Blue
    Call glBegin(GL_LINES)
        Call glVertex3f(0, 0, 0)
        Call glVertex3f(0, 0, 1000)
        Call glVertex3f(-50, 0, 900) 'Arrow
        Call glVertex3f(0, 0, 1000)
        Call glVertex3f(0, 0, 1000)
        Call glVertex3f(50, 0, 900)
    Call glEnd
End Function

Function DrawObject()
Dim X, Y As Integer
    Call glMatrixMode(GL_MODELVIEW)
        Call glPushMatrix
        Call glTranslatef(-5875, -5875, 0)
        For Y = 51 To 4 Step -1
            Call glPushMatrix
            For X = 4 To 51
                Select Case ActiveSheet.Cells(Y, X).Value
                    Case 1
                        Call glColor3f(0.75, 0.75, 0.75) 'Gray
                        Call DrawCap
                    Case 2
                        Call glColor3f(0.15, 0.75, 0.15) 'Green
                        Call DrawCap
                    Case 3
                        Call glColor3f(0.85, 0.85, 0.15) 'Yellow
                        Call DrawCap
                    Case 4
                        Call glColor3f(0.85, 0.15, 0.15) 'Red
                        Call DrawCap
                    Case 5
                        Call glColor3f(0.25, 0.4, 0.85) 'Blue
                        Call DrawCap
                End Select
                Call glTranslatef(250, 0, 0)
            Next X
            Call glPopMatrix
            Call glTranslatef(0, 250, 0)
        Next Y
        Call glPopMatrix
End Function

Function DrawCap()
    'Call glPushMatrix
    'Call glRotatef(CSng(5 * Rnd - 2.5), 0, 0, 1)
    'Call glTranslatef(0, 0, CSng(10 * Rnd))
        Call glBegin(GL_QUADS)
            Call glVertex3f(-100, -100, 0)
            Call glVertex3f(100, -100, 0)
            Call glVertex3f(100, 100, 0)
            Call glVertex3f(-100, 100, 0)
        Call glEnd
        Call glBegin(GL_QUAD_STRIP)
            Call glVertex3f(-120, -120, -20)
            Call glVertex3f(-100, -100, 0)
            Call glVertex3f(120, -120, -20)
            Call glVertex3f(100, -100, 0)
            Call glVertex3f(120, 120, -20)
            Call glVertex3f(100, 100, 0)
            Call glVertex3f(-120, 120, -20)
            Call glVertex3f(-100, 100, 0)
            Call glVertex3f(-120, -120, -20)
            Call glVertex3f(-100, -100, 0)
        Call glEnd
    'Call glPopMatrix
End Function


