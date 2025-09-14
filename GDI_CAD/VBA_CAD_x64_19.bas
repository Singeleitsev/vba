Attribute VB_Name = "VBA_CAD"
'For 32 bit & 64 bit

Option Explicit

'WinMain
Declare PtrSafe Function GetModuleHandleW Lib "Kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
Declare PtrSafe Function LoadIconW Lib "User32" (ByVal hInstance As LongPtr, ByVal lpIconName As LongPtr) As LongPtr
Declare PtrSafe Function LoadCursorW Lib "User32" (ByVal hInstance As LongPtr, ByVal lpCursorName As LongPtr) As LongPtr
Declare PtrSafe Function GetStockObject Lib "Gdi32" (ByVal i As LongPtr) As LongPtr
Declare PtrSafe Function RegisterClassExW Lib "User32" (ByVal lpWndClass As LongPtr) As LongPtr
Declare PtrSafe Function CreateWindowExW Lib "User32" ( _
    ByVal dwExStyle As LongPtr, _
    ByVal lpClassName As LongPtr, _
    ByVal lpWindowName As LongPtr, _
    ByVal dwStyle As LongPtr, _
    ByVal X As LongPtr, _
    ByVal Y As LongPtr, _
    ByVal nWidth As LongPtr, _
    ByVal nHeight As LongPtr, _
    ByVal hWndParent As LongPtr, _
    ByVal hMenu As LongPtr, _
    ByVal hInstance As LongPtr, _
    ByVal lpParam As LongPtr _
    ) As LongPtr
Declare PtrSafe Function ShowWindow Lib "User32" (ByVal hWnd As LongPtr, ByVal nCmdShow As LongPtr) As LongPtr
Declare PtrSafe Function UpdateWindow Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function SetForegroundWindow Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function PeekMessageW Lib "User32" (ByVal lpMsg As LongPtr, ByVal hWnd As LongPtr, ByVal wMsgFilterMin As LongPtr, ByVal wMsgFilterMax As LongPtr, ByVal wRemoveMsg As LongPtr) As LongPtr
Declare PtrSafe Function TranslateMessage Lib "User32" (ByVal lpMsg As LongPtr) As LongPtr
Declare PtrSafe Function DispatchMessageW Lib "User32" (ByVal lpMsg As LongPtr) As LongPtr
Declare PtrSafe Function SendMessageW Lib "User32" (ByVal hWnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Declare PtrSafe Function GetLastError Lib "Kernel32" () As LongPtr
'StatusBar
Declare PtrSafe Function InitCommonControlsEx Lib "Comctl32" (ByVal picce As LongPtr) As LongPtr
'WndProc
Declare PtrSafe Function ClientToScreen Lib "User32" (ByVal hWnd As LongPtr, ByVal lpPoint As LongPtr) As LongPtr
Declare PtrSafe Function GetKeyState Lib "User32" (ByVal nVirtKey As LongPtr) As LongPtr
Declare PtrSafe Function MessageBoxW Lib "User32" (ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As LongPtr) As LongPtr
Declare PtrSafe Function DefWindowProcW Lib "User32" (ByVal hWnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Declare PtrSafe Function DestroyWindow Lib "User32" (ByVal hWnd As LongPtr) As LongPtr
'GDI & Paint
Declare PtrSafe Function GetClientRect Lib "User32" (ByVal hWnd As LongPtr, ByVal lpRect As LongPtr) As LongPtr
Declare PtrSafe Function DeleteDC Lib "Gdi32" (ByVal hdc As LongPtr) As LongPtr
Declare PtrSafe Function BeginPaint Lib "User32" (ByVal hWnd As LongPtr, ByVal lpPaint As LongPtr) As LongPtr
Declare PtrSafe Function EndPaint Lib "User32" (ByVal hWnd As LongPtr, ByVal lpPaint As LongPtr) As LongPtr
Declare PtrSafe Function InvalidateRect Lib "User32" (ByVal hWnd As LongPtr, ByVal lpRect As LongPtr, ByVal bErase As LongPtr) As LongPtr
Declare PtrSafe Function SelectObject Lib "Gdi32" (ByVal hdc As LongPtr, ByVal h As LongPtr) As LongPtr
'Declare PtrSafe Function BeginPath Lib "Gdi32" (ByVal hDC As LongPtr) As LongPtr
'Declare PtrSafe Function EndPath Lib "Gdi32" (ByVal hDC As LongPtr) As LongPtr
Declare PtrSafe Function MoveToEx Lib "Gdi32" (ByVal hdc As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal lpPt As LongPtr) As LongPtr
Declare PtrSafe Function LineTo Lib "Gdi32" (ByVal hdc As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr) As LongPtr
Declare PtrSafe Function Ellipse Lib "Gdi32" (ByVal hdc As LongPtr, ByVal left As LongPtr, ByVal top As LongPtr, ByVal right As LongPtr, ByVal bottom As LongPtr) As LongPtr
'Text
'Declare PtrSafe Function SetTextAlign Lib "Gdi32" (ByVal hDC As LongPtr, ByVal align As LongPtr) As LongPtr
'Declare PtrSafe Function SetTextColor Lib "Gdi32" (ByVal hDC As LongPtr, ByVal color As LongPtr) As LongPtr
'Declare PtrSafe Function DrawTextW Lib "User32" (ByVal hDC As LongPtr, ByVal lpchText As LongPtr, ByVal cchText As LongPtr, ByVal lprc As LongPtr, ByVal format As LongPtr) As LongPtr
Declare PtrSafe Function TextOutW Lib "Gdi32" (ByVal hdc As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal lpString As LongPtr, ByVal c As LongPtr) As LongPtr
'PopupMenu
Declare PtrSafe Function CreateMenu Lib "User32" () As LongPtr '<-
Declare PtrSafe Function CreatePopupMenu Lib "User32" () As LongPtr
Declare PtrSafe Function SetMenu Lib "User32" (ByVal hWnd As LongPtr, ByVal hMenu As LongPtr) As LongPtr
Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal hWnd As LongPtr) As LongPtr '<-
Declare PtrSafe Function TrackPopupMenu Lib "User32" ( _
    ByVal hMenu As LongPtr, _
    ByVal uFlags As LongPtr, _
    ByVal X As LongPtr, _
    ByVal Y As LongPtr, _
    ByVal nReserved As LongPtr, _
    ByVal hWnd As LongPtr, _
    ByVal prcRect_Ignored As LongPtr) _
    As LongPtr
Declare PtrSafe Function AppendMenuW Lib "User32" (ByVal hMenu As LongPtr, ByVal uFlags As LongPtr, ByVal uIDNewItem As LongPtr, ByVal lpNewItem As LongPtr) As LongPtr
Declare PtrSafe Function CheckMenuItem Lib "User32" (ByVal hMenu As LongPtr, ByVal uIDCheckItem As LongPtr, ByVal uCheck As LongPtr) As LongPtr
Declare PtrSafe Function DestroyMenu Lib "User32" (ByVal hMenu As LongPtr) As LongPtr

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
Public Pt As POINT2D
Public lpPt As LongPtr
Public PtMenuRMB As POINT2D
Public lpPtMenuRMB As LongPtr
'MSG msg;
Type Msg
    hWnd As LongPtr
    message As Long
    wParam As LongPtr
    lParam As LongPtr
    time As Long
    Pt As POINT2D
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
'INITCOMMONCONTROLSEX icce;
Type ICCESTRUCT
    dwSize As Long
    dwICC As Long
End Type
Public icce As ICCESTRUCT
Public lpIcce As LongPtr 'Pointer
'PAINTSTRUCT ps
Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint(3) As Long
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(31) As Byte
End Type
Public ps As PAINTSTRUCT
Public lpPaint As LongPtr

'Const
Public Const PM_REMOVE = 1
Public Const CS_VREDRAW = 1
Public Const CS_HREDRAW = 2
Public Const CS_DBLCLKS = 8
Public Const IDI_APPLICATION = 32512
Public Const IDC_ARROW = 32512
Public Const COLOR_WINDOW = 5
Public Const CW_USEDEFAULT = &H80000000
'Window Styles
Public Const WS_OVERLAPPED = 0
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_OVERLAPPEDWINDOW = WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX
Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const SW_SHOW = 5
Public Const SBARS_SIZEGRIP = &H100
Public Const SBT_NOBORDERS = &H101 '0x0100
'Messages
Public Const WM_CREATE = 1
Public Const WM_DESTROY = 2
Public Const WM_SIZE = 5
Public Const WM_PAINT = 15
Public Const WM_CLOSE = &H10
Public Const WM_QUIT = &H12
Public Const WM_NCHITTEST = &H84
'Public Const WM_KEYDOWN = &H100 'Does not work
Public Const WM_KEYUP = &H101
Public Const WM_COMMAND = &H111
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSEWHEEL = &H20A
Public Const WM_USER = &H400
Public Const SB_SETPARTS = WM_USER + 4 '0x0404
Public Const SB_SETTEXTW = WM_USER + 11 '0x040B
'Virtual Keys
Public Const VK_SHIFT = &H10
Public Const VK_ESCAPE = &H1B
Public Const VK_SPACE = &H20
Public Const MB_OK = 0
Public Const MB_YESNO = 4
Public Const MB_ICONQUESTION = &H20
Public Const IDYES = 6
Public Const IDNO = 7
Public Const DT_TOP = 0
Public Const DT_LEFT = 0
Public Const DT_CENTER = 1
Public Const DT_RIGHT = 2
Public Const DT_BOTTOM = 8
'Brushes
Public Const WHITE_BRUSH = 0
Public Const GRAY_BRUSH = 2
Public Const BLACK_BRUSH = 4
Public Const NULL_BRUSH = 5
'https://learn.microsoft.com/en-us/dotnet/api/system.drawing.copypixeloperation
Public Const PATCOPY = 15728673
Public Const SRCCOPY = 13369376
Public Const WHITENESS = 16711778
'https://learn.microsoft.com/en-us/windows/win32/api/commctrl/ns-commctrl-initcommoncontrolsex
Public Const ICC_BAR_CLASSES = 4
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
'Menu ID's - Custom Values
Public Const IDM_FILE_NEW = 1001
Public Const IDM_FILE_OPEN = 1002
Public Const IDM_FILE_SAVE = 1003
Public Const IDM_FILE_SAVE_AS = 1004
Public Const IDM_APP_EXIT = 1005
Public Const IDM_EDIT_UNDO = 2001
Public Const IDM_EDIT_REDO = 2002
Public Const IDM_EDIT_CUT = 2003
Public Const IDM_EDIT_COPY = 2004
Public Const IDM_EDIT_PASTE = 2005
Public Const IDM_EDIT_CLEAR = 2006
Public Const IDM_VIEW_MATRICES = 3001
Public Const IDM_DRAW_LINE = 4001
Public Const IDM_DRAW_CIRCLE = 4002

'Global Handles
Public ghInst As LongPtr
Public ghWnd As LongPtr
'Main Window
Public nWndClass As LongPtr 'For Debug Purpose
Public lpWndProc As LongPtr
Public ghDC As LongPtr
'Public ghBrush As LongPtr
'Public ghBit As LongPtr

'Name Strings
Public szAppName As String
Public szClassName As String
Public szCaption As String
'Name Pointers
Public lpszAppName As LongPtr
Public lpszCaption As LongPtr

'Status Bar
Public hwndStatusBar As LongPtr
Public lpStatusProc As LongPtr
Public xStatusParts(3) As Long 'Right edge coordinates for 4 parts
Public lpStatusParts As LongPtr
'Status Strings
Public szStatusMessage As String
'Status Pointers
Public lpszStatusMessage As LongPtr

'Menu Handles
Public hMenu As LongPtr
Public hMenuFile As LongPtr
Public hMenuEdit As LongPtr
Public hMenuView As LongPtr
Public hMenuDraw As LongPtr
Public hMenuRMB As LongPtr
'Menu Text Strings
Public szMenuFile As String
Public szMenuFileNew As String
Public szMenuFileOpen As String
Public szMenuFileSave As String
Public szMenuFileSaveAs As String
Public szMenuFileExit As String
Public szMenuEdit As String
Public szMenuEditUndo As String
Public szMenuEditRedo As String
Public szMenuEditCut As String
Public szMenuEditCopy As String
Public szMenuEditPaste As String
Public szMenuEditDelete As String
Public szMenuView As String
Public szMenuViewMatrices As String
Public szMenuDraw As String
Public szMenuDrawLine As String
Public szMenuDrawCircle As String
'Menu Text Pointers
Public lpszMenuFile As LongPtr
Public lpszMenuFileNew As LongPtr
Public lpszMenuFileOpen As LongPtr
Public lpszMenuFileSave As LongPtr
Public lpszMenuFileSaveAs As LongPtr
Public lpszMenuFileExit As LongPtr
Public lpszMenuEdit As LongPtr
Public lpszMenuEditUndo As LongPtr
Public lpszMenuEditRedo As LongPtr
Public lpszMenuEditCut As LongPtr
Public lpszMenuEditCopy As LongPtr
Public lpszMenuEditPaste As LongPtr
Public lpszMenuEditDelete As LongPtr
Public lpszMenuView As String
Public lpszMenuViewMatrices As String
Public lpszMenuDraw As LongPtr
Public lpszMenuDrawLine As LongPtr
Public lpszMenuDrawCircle As LongPtr

'Cursor
Public xCursorScreen As Integer
Public yCursorScreen As Integer
Public xCursorWorld As Double
Public yCursorWorld As Double
Public szCursorX As String
Public szCursorY As String
Public lpszCursorX As LongPtr
Public lpszCursorY As LongPtr

'Draw
Public nDrawMode As Long
Public nActivePrimitive As Long
'Lines - World Coordinates Array
Public Const PRIMITIVE_LINE = 1
Public nLine As Long ' 255 Lines Possible, nLine = 0 Must Be Skipped
Public xLineStart(255) As Double
Public yLineStart(255) As Double
Public zLineStart(255) As Double
Public xLineEnd(255) As Double
Public yLineEnd(255) As Double
Public zLineEnd(255) As Double
'Circles - World Coordinates Array
Public Const PRIMITIVE_CIRCLE = 2
Public nCircle As Long ' 255 Circles Possible, nCircle = 0 Must Be Skipped
Public xCircleCenter(255) As Double
Public yCircleCenter(255) As Double
Public zCircleCenter(255) As Double
Public rCircleRadius(255) As Double
Public xCircleBoundaryLeft As Integer
Public yCircleBoundaryTop As Integer
Public xCircleBoundaryRight As Integer
Public yCircleBoundaryBottom As Integer
'Vertices Buffer
Public xVertexScreen As Integer
Public yVertexScreen As Integer
'Pan
Public dxPan As Integer
Public dyPan As Integer
'Zoom
Public fDelta As Double
Public fScale As Double
'Matrices
Public nMatrixShow As String
Public szMatrix As String
Public lpszMatrix As LongPtr
'View Matrix
Public mtxView(2, 2) As Double
'Inverse Matrix
Public mtxInv(2, 2) As Double
'Transformation Matrix
Public mtxTrans(2, 2) As Double


'Messages
Public szMsgText As String
Public lpszMsgText As LongPtr
Public szMsgTitle As String
Public lpszMsgTitle As LongPtr

'Debug
Public nLastError As LongPtr
'Public nWndProcPass As LongPtr

Sub SetButton()
    Dim wshWorkSheet As Worksheet
    Dim btnStart As Button
    Dim r As Range
 
    Set wshWorkSheet = ActiveSheet
    wshWorkSheet.Buttons.Delete
    
    Set r = wshWorkSheet.Cells(2, 2)
    Set btnStart = wshWorkSheet.Buttons.Add(r.left, r.top, r.Width, r.Height)
    With btnStart
        .OnAction = "Start"
        .Caption = "Start"
        .Name = "Start"
    End With
End Sub

Sub Start()
    Dim echo As LongPtr
    
    ActiveWindow.WindowState = xlMinimized
    ThisWorkbook.Saved = True
    
    szAppName = "AutoCAD Home Edition"
    lpszAppName = StrPtr(szAppName)
    
    echo = WinMain(0, 0, 0, 0)
End Sub

Function WinMain(ByVal hInstance As LongPtr, ByVal hPrevInstance As LongPtr, ByVal lpCmdLine As LongPtr, ByVal nCmdShow As LongPtr) As LongPtr
'Pointers
    lpWndClass = VarPtr(wcx)
    lpMsg = VarPtr(wmsg)
'WndProc
    lpRectMain = VarPtr(RectMain)
    lpPaint = VarPtr(ps)
    lpIcce = VarPtr(icce)
    lpPtMenuRMB = VarPtr(PtMenuRMB)
'Main Cycle
    Dim nPeek As LongPtr

'WNDCLASSEX wcx;
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
    wcx.hbrBackground = GetStockObject(WHITE_BRUSH) 'COLOR_WINDOW
    wcx.lpszMenuName = 0
    szClassName = "MainWindowClassName"
    wcx.lpszClassName = StrPtr(szClassName)
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

    ghWnd = CreateWindowExW(0, wcx.lpszClassName, lpszCaption, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, 0, 0, wcx.hInstance, 0)
    If ghWnd = 0 Then
        Call MsgBox("CreateWindow Error: " & GetLastError())
        Exit Function
    End If

'INITCOMMONCONTROLSEX
    icce.dwSize = Len(icce)
    icce.dwICC = ICC_BAR_CLASSES 'Load toolbar, status bar, trackbar, and tooltip control classes

    Call ShowWindow(ghWnd, SW_SHOW)
    Call UpdateWindow(ghWnd)
    Call SetForegroundWindow(ghWnd)

WinMainLoop:
    If wmsg.message = WM_QUIT Then
        WinMain = wmsg.wParam
        Exit Function
    End If
    nPeek = PeekMessageW(lpMsg, ghWnd, 0, 0, PM_REMOVE)
    If nPeek = 0 Or nPeek = -1 Then
        WinMain = 0
        Exit Function
    Else
        Call TranslateMessage(lpMsg)
        Call DispatchMessageW(lpMsg)
    End If
    GoTo WinMainLoop
End Function

Function GetAddr(ByVal lpProc As LongPtr) As LongPtr
    'This Function has been created to fit the 'AddressOf' syntax
    GetAddr = lpProc
End Function

Function WndProc(ByVal hWnd As LongPtr, ByVal message As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
'Argument
    Dim wParamLowWord As Integer
'Counter
    Dim i As Long

If message = WM_PAINT Then
    ghDC = BeginPaint(hWnd, lpPaint)

    Call ShowMatrices

    'Coordinates in the Status Bar
    szCursorX = "X = " & Format(xCursorScreen, "0000")
    lpszCursorX = StrPtr(szCursorX)
    szCursorY = "Y = " & Format(yCursorScreen, "0000")
    lpszCursorY = StrPtr(szCursorY)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 0, lpszCursorX)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 1, lpszCursorY)
    Call SendMessageW(hwndStatusBar, SB_SETTEXTW, 2 Or SBT_NOBORDERS, lpszStatusMessage)

    'Draw All Lines
    For i = 1 To nLine 'nLine = 0 Must Be Skipped
        xVertexScreen = CInt(xLineStart(i) * mtxView(0, 0) + yLineStart(i) * mtxView(0, 1) + mtxView(0, 2))
        yVertexScreen = CInt(xLineStart(i) * mtxView(1, 0) + yLineStart(i) * mtxView(1, 1) + mtxView(1, 2))
        Call MoveToEx(ghDC, xVertexScreen, yVertexScreen, lpPt)
        xVertexScreen = CInt(xLineEnd(i) * mtxView(0, 0) + yLineEnd(i) * mtxView(0, 1) + mtxView(0, 2))
        yVertexScreen = CInt(xLineEnd(i) * mtxView(1, 0) + yLineEnd(i) * mtxView(1, 1) + mtxView(1, 2))
        Call LineTo(ghDC, xVertexScreen, yVertexScreen)
    Next i
    'Draw All Circles
    For i = 1 To nCircle 'nCircle = 0 Must Be Skipped
        Call SelectObject(ghDC, GetStockObject(NULL_BRUSH))
        xCircleBoundaryLeft = CInt((xCircleCenter(i) - rCircleRadius(i)) * mtxView(0, 0) + (yCircleCenter(i) - rCircleRadius(i)) * mtxView(0, 1) + mtxView(0, 2))
        yCircleBoundaryTop = CInt((yCircleCenter(i) - rCircleRadius(i)) * mtxView(1, 0) + (yCircleCenter(i) - rCircleRadius(i)) * mtxView(1, 1) + mtxView(1, 2))
        xCircleBoundaryRight = CInt((xCircleCenter(i) + rCircleRadius(i)) * mtxView(0, 0) + (yCircleCenter(i) + rCircleRadius(i)) * mtxView(0, 1) + mtxView(0, 2))
        yCircleBoundaryBottom = CInt((yCircleCenter(i) + rCircleRadius(i)) * mtxView(1, 0) + (yCircleCenter(i) + rCircleRadius(i)) * mtxView(1, 1) + mtxView(1, 2))
        Call Ellipse(ghDC, xCircleBoundaryLeft, yCircleBoundaryTop, xCircleBoundaryRight, yCircleBoundaryBottom)
    Next i

    Call EndPaint(hWnd, lpPaint)
    GoTo lbl_WndProc_Return0
End If

'Mouse Handlers
If message = WM_MOUSEMOVE Then
    xCursorScreen = CLng(lParam And 32767) 'Low Word, Signed Integer
    yCursorScreen = CLng(lParam / 65536) 'High Word
    'Get World Coordinates from the Screen
    xCursorWorld = CDbl(xCursorScreen) * mtxInv(0, 0) + CDbl(yCursorScreen) * mtxInv(0, 1) + mtxInv(0, 2)
    yCursorWorld = CDbl(xCursorScreen) * mtxInv(1, 0) + CDbl(yCursorScreen) * mtxInv(1, 1) + mtxInv(1, 2)
    If nDrawMode = 1 Then
        If nActivePrimitive = PRIMITIVE_LINE Then
            'If GetKeyState(VK_SHIFT) < 0 Then
            If GetKeyState(VK_SHIFT) > 32767 Then 'Signed Integer
                If Abs(xCursorWorld - xLineStart(nLine)) > Abs(yCursorWorld - yLineStart(nLine)) Then
                    xLineEnd(nLine) = xCursorWorld
                    yLineEnd(nLine) = yLineStart(nLine)
                    GoTo wm_200_Redraw
                Else
                    xLineEnd(nLine) = xLineStart(nLine)
                    yLineEnd(nLine) = yCursorWorld
                    GoTo wm_200_Redraw
                End If
            Else
                xLineEnd(nLine) = xCursorWorld
                yLineEnd(nLine) = yCursorWorld
                GoTo wm_200_Redraw
            End If
        End If
        If nActivePrimitive = PRIMITIVE_CIRCLE Then
            rCircleRadius(nCircle) = Sqr((xCursorWorld - xCircleCenter(nCircle)) ^ 2 + (yCursorWorld - yCircleCenter(nCircle)) ^ 2)
            GoTo wm_200_Redraw
        End If
    End If
wm_200_Redraw:
    Call InvalidateRect(hWnd, 0, 1)
    GoTo lbl_WndProc_Return0
End If
If message = WM_MOUSEWHEEL Then
    fDelta = CDbl(wParam / 65536) 'High Word =120 or =-120
    fScale = 1 + (fDelta / 1000) '=1+(120/1000)=1.12 or =1+(-120/1000)=0.88
    Call ClearTransformationMatrix
        mtxTrans(0, 2) = -CDbl(xCursorScreen)
        mtxTrans(1, 2) = -CDbl(yCursorScreen)
    Call MatrixMultiply
    Call ClearTransformationMatrix
        mtxTrans(0, 0) = fScale
        mtxTrans(1, 1) = fScale
    Call MatrixMultiply
    Call ClearTransformationMatrix
        mtxTrans(0, 2) = CDbl(xCursorScreen)
        mtxTrans(1, 2) = CDbl(yCursorScreen)
    Call MatrixMultiply
    Call InvalidateRect(hWnd, 0, 1)
    GoTo lbl_WndProc_Return0
End If
If message = WM_LBUTTONDOWN Then
    If nDrawMode = 0 Then
        If nActivePrimitive = PRIMITIVE_LINE Then
            nLine = nLine + 1 'nLine = 0 Must Be Skipped
            xLineStart(nLine) = xCursorWorld
            yLineStart(nLine) = yCursorWorld
            nDrawMode = 1
            GoTo lbl_WndProc_Return0
        End If
        If nActivePrimitive = PRIMITIVE_CIRCLE Then
            nCircle = nCircle + 1 'nCircle = 0 Must Be Skipped
            xCircleCenter(nCircle) = xCursorWorld
            yCircleCenter(nCircle) = yCursorWorld
            nDrawMode = 1
            GoTo lbl_WndProc_Return0
        End If
    Else
        nDrawMode = 0
        GoTo lbl_WndProc_Return0
    End If
    GoTo lbl_WndProc_Return0
End If
If message = WM_RBUTTONUP Then
    If nDrawMode = 1 Then
        If nActivePrimitive = PRIMITIVE_LINE Then
            nLine = nLine - 1
            nDrawMode = 0
            GoTo lbl_WndProc_Return0
        End If
        If nActivePrimitive = PRIMITIVE_CIRCLE Then
            nCircle = nCircle - 1
            nDrawMode = 0
            GoTo lbl_WndProc_Return0
        End If
    Else
        PtMenuRMB.X = CLng(lParam And 32767) 'Low Word, Signed Integer
        PtMenuRMB.Y = CLng(lParam / 65536) 'High Word
        Call ClientToScreen(hWnd, lpPtMenuRMB)
        Call TrackPopupMenu(hMenuRMB, TPM_LEFTALIGN Or TPM_RIGHTBUTTON, PtMenuRMB.X, PtMenuRMB.Y, 0, hWnd, 0)
        GoTo lbl_WndProc_Return0
    End If
    GoTo lbl_WndProc_Return0
End If
If message = WM_MBUTTONDOWN Then
    dxPan = CLng(lParam And 32767) 'Low Word, Signed Integer
    dyPan = CLng(lParam / 65536) 'High Word
    Call InvalidateRect(hWnd, 0, 1)
    GoTo lbl_WndProc_Return0
End If
If message = WM_MBUTTONUP Then
    dxPan = CLng(lParam And 32767) - dxPan
    dyPan = CLng(lParam / 65536) - dyPan
    Call ClearTransformationMatrix
        mtxTrans(0, 2) = CDbl(dxPan)
        mtxTrans(1, 2) = CDbl(dyPan)
    Call MatrixMultiply
    Call InvalidateRect(hWnd, 0, 1)
    GoTo lbl_WndProc_Return0
End If
If message = WM_MBUTTONDBLCLK Then
    If nLine + nCircle = 0 Then
        GoTo lbl_WndProc_Return0
    End If
    Call GetClientRect(hWnd, lpRectMain)
    Call ZoomExtents(RectMain.right - RectMain.left, RectMain.bottom - RectMain.top)
    'Correct the Error Caused by WM_MBUTTONUP
    Call ClearTransformationMatrix
        mtxTrans(0, 2) = -CLng(lParam And 32767)
        mtxTrans(1, 2) = -CLng(lParam / 65536)
    Call MatrixMultiply
    Call InvalidateRect(hWnd, 0, 1)
    GoTo lbl_WndProc_Return0
End If

'KeyStroke Handlers
If message = WM_KEYUP Then
    wParamLowWord = CLng(wParam And 32767) 'Low Word, Signed Integer
    If wParamLowWord = VK_ESCAPE Then
        If nDrawMode = 1 Then
            If nActivePrimitive = PRIMITIVE_LINE Then
                nLine = nLine - 1
                nDrawMode = 0
                GoTo lbl_WndProc_Return0
            End If
            If nActivePrimitive = PRIMITIVE_CIRCLE Then
                nCircle = nCircle - 1
                nDrawMode = 0
                GoTo lbl_WndProc_Return0
            End If
        Else
            Call LoadIdentity
            GoTo lbl_WndProc_Return0
        End If
    End If
    If wParamLowWord = VK_SPACE Then
        If nDrawMode = 0 Then
            Call GetClientRect(hWnd, lpRectMain)
            Call ZoomExtents(RectMain.right - RectMain.left, RectMain.bottom - RectMain.top)
            Call InvalidateRect(hWnd, 0, 1)
        End If
    End If
    If wParamLowWord = &H31 Then 'The "1" Button
        If nDrawMode = 0 Then
            nActivePrimitive = PRIMITIVE_LINE
            Call CheckMenuItem(hMenu, IDM_DRAW_LINE, MF_CHECKED)
            Call CheckMenuItem(hMenu, IDM_DRAW_CIRCLE, MF_UNCHECKED)
            Call CheckMenuItem(hMenuRMB, IDM_DRAW_LINE, MF_CHECKED)
            Call CheckMenuItem(hMenuRMB, IDM_DRAW_CIRCLE, MF_UNCHECKED)
        End If
    End If
    If wParamLowWord = &H32 Then 'The "2" Button
        If nDrawMode = 0 Then
            nActivePrimitive = PRIMITIVE_CIRCLE
            Call CheckMenuItem(hMenu, IDM_DRAW_LINE, MF_UNCHECKED)
            Call CheckMenuItem(hMenu, IDM_DRAW_CIRCLE, MF_CHECKED)
            Call CheckMenuItem(hMenuRMB, IDM_DRAW_LINE, MF_UNCHECKED)
            Call CheckMenuItem(hMenuRMB, IDM_DRAW_CIRCLE, MF_CHECKED)
        End If
    End If
    GoTo lbl_WndProc_Return0
End If

'Menu Handlers
If message = WM_COMMAND Then
    If wParam = IDM_APP_EXIT Then
        szMsgText = "Close?"
        lpszMsgText = StrPtr(szMsgText)
        szMsgTitle = "Such A Good Application"
        lpszMsgTitle = StrPtr(szMsgTitle)
        If MessageBoxW(hWnd, lpszMsgText, lpszMsgTitle, MB_YESNO Or MB_ICONQUESTION) = IDYES Then
            Call DestroyWindow(hWnd)
        End If
        GoTo lbl_WndProc_Return0
    End If
    If wParam = IDM_DRAW_LINE Then
        nActivePrimitive = PRIMITIVE_LINE
        Call CheckMenuItem(hMenu, IDM_DRAW_LINE, MF_CHECKED)
        Call CheckMenuItem(hMenu, IDM_DRAW_CIRCLE, MF_UNCHECKED)
        Call CheckMenuItem(hMenuRMB, IDM_DRAW_LINE, MF_CHECKED)
        Call CheckMenuItem(hMenuRMB, IDM_DRAW_CIRCLE, MF_UNCHECKED)
        szStatusMessage = "Draw a Line"
        lpszStatusMessage = StrPtr(szStatusMessage)
        GoTo lbl_WndProc_Return0
    End If
    If wParam = IDM_DRAW_CIRCLE Then
        nActivePrimitive = PRIMITIVE_CIRCLE
        Call CheckMenuItem(hMenu, IDM_DRAW_LINE, MF_UNCHECKED)
        Call CheckMenuItem(hMenu, IDM_DRAW_CIRCLE, MF_CHECKED)
        Call CheckMenuItem(hMenuRMB, IDM_DRAW_LINE, MF_UNCHECKED)
        Call CheckMenuItem(hMenuRMB, IDM_DRAW_CIRCLE, MF_CHECKED)
        szStatusMessage = "Draw a Circle"
        lpszStatusMessage = StrPtr(szStatusMessage)
        GoTo lbl_WndProc_Return0
    End If
    If wParam = IDM_VIEW_MATRICES Then
        If nMatrixShow = 1 Then
            nMatrixShow = 0
            Call CheckMenuItem(hMenu, IDM_VIEW_MATRICES, MF_UNCHECKED)
            GoTo lbl_WndProc_Return0
        Else
            nMatrixShow = 1
            Call CheckMenuItem(hMenu, IDM_VIEW_MATRICES, MF_CHECKED)
            GoTo lbl_WndProc_Return0
        End If
    End If
    GoTo lbl_DefWndProc 'lbl_WndProc_Return0
End If

'Customer's Message Handlers
'If Message = WM_NCHITTEST Then
    'GoTo lbl_WndProc_Return0 'Skip WM_NCHITTEST
'End If

'System Message Handlers
If message = WM_SIZE Then
    Call GetClientRect(hWnd, lpRectMain)
    xStatusParts(3) = RectMain.right - RectMain.left - 50
    Call SendMessageW(hwndStatusBar, WM_SIZE, 0, 0)
    GoTo lbl_WndProc_Return0
End If
If message = WM_CLOSE Then
    szMsgText = "Close?"
    lpszMsgText = StrPtr(szMsgText)
    szMsgTitle = "Such A Good Application"
    lpszMsgTitle = StrPtr(szMsgTitle)
    If MessageBoxW(hWnd, lpszMsgText, lpszMsgTitle, MB_YESNO Or MB_ICONQUESTION) = IDYES Then
        Call DestroyWindow(hWnd)
    End If
    GoTo lbl_WndProc_Return0
End If
If message = WM_DESTROY Then
    Call DeleteDC(ghDC)
    Call DestroyMenu(hMenu)
    'Don't Call PostQuitMessage(0)
    ActiveWindow.WindowState = xlMaximized
    GoTo lbl_WndProc_Return0
End If

'Initialization
If message = WM_CREATE Then
    Call InitCommonControlsEx(lpIcce)
    Call DoCreateMenu(hWnd)
    Call DoCreateStatusBar(hWnd, RectMain.right - RectMain.left)
    'Call ClearAll
    Call LoadIdentity
    'Initialize Counters and Flags
    nDrawMode = 0
    nActivePrimitive = PRIMITIVE_LINE
    nLine = 0
    nCircle = 0
    nMatrixShow = 1
    szStatusMessage = "Draw a Line"
    lpszStatusMessage = StrPtr(szStatusMessage)
    GoTo lbl_WndProc_Return0
End If

lbl_DefWndProc:
    WndProc = DefWindowProcW(hWnd, message, wParam, lParam)
    Exit Function
lbl_WndProc_Return0:
    WndProc = 0
End Function

Sub DoCreateMenu(ByVal hWnd As LongPtr)
'Menu Text  Strings
    szMenuFile = "&File"
    szMenuFileNew = "&New" + vbTab + "Ctrl+N"
    szMenuFileOpen = "&Open..." + vbTab + "Ctrl+O"
    szMenuFileSave = "&Save" + vbTab + "Ctrl+S"
    szMenuFileSaveAs = "Save &As..." + vbTab + "Ctrl+Shift+S"
    szMenuFileExit = "E&xit" + vbTab + "Ctrl+W"
    szMenuEdit = "&Edit"
    szMenuEditUndo = "&Undo" + vbTab + "Ctrl+Z"
    szMenuEditRedo = "Redo" + vbTab + "Ctrl+Y"
    szMenuEditCut = "Cu&t" + vbTab + "Ctrl+X"
    szMenuEditCopy = "&Copy" + vbTab + "Ctrl+C"
    szMenuEditPaste = "&Paste" + vbTab + "Ctrl+V"
    szMenuEditDelete = "De&lete" + vbTab + "Del"
    szMenuView = "View"
    szMenuViewMatrices = "Matrices"
    szMenuDraw = "Draw"
    szMenuDrawLine = "Line" + vbTab + "1"
    szMenuDrawCircle = "Circle" + vbTab + "2"
'Menu Text Pointers
    lpszMenuFile = StrPtr(szMenuFile)
    lpszMenuFileNew = StrPtr(szMenuFileNew)
    lpszMenuFileOpen = StrPtr(szMenuFileOpen)
    lpszMenuFileSave = StrPtr(szMenuFileSave)
    lpszMenuFileSaveAs = StrPtr(szMenuFileSaveAs)
    lpszMenuFileExit = StrPtr(szMenuFileExit)
    lpszMenuEdit = StrPtr(szMenuEdit)
    lpszMenuEditUndo = StrPtr(szMenuEditUndo)
    lpszMenuEditRedo = StrPtr(szMenuEditRedo)
    lpszMenuEditCut = StrPtr(szMenuEditCut)
    lpszMenuEditCopy = StrPtr(szMenuEditCopy)
    lpszMenuEditPaste = StrPtr(szMenuEditPaste)
    lpszMenuEditDelete = StrPtr(szMenuEditDelete)
    lpszMenuView = StrPtr(szMenuView)
    lpszMenuViewMatrices = StrPtr(szMenuViewMatrices)
    lpszMenuDraw = StrPtr(szMenuDraw)
    lpszMenuDrawLine = StrPtr(szMenuDrawLine)
    lpszMenuDrawCircle = StrPtr(szMenuDrawCircle)

'Main Menu
    hMenu = CreateMenu()
    hMenuFile = CreatePopupMenu()
    Call AppendMenuW(hMenu, MF_POPUP, hMenuFile, lpszMenuFile)
        Call AppendMenuW(hMenuFile, MF_STRING Or MF_GRAYED, IDM_FILE_NEW, lpszMenuFileNew)
        Call AppendMenuW(hMenuFile, MF_STRING Or MF_GRAYED, IDM_FILE_OPEN, lpszMenuFileOpen)
        Call AppendMenuW(hMenuFile, MF_STRING Or MF_GRAYED, IDM_FILE_SAVE, lpszMenuFileSave)
        Call AppendMenuW(hMenuFile, MF_STRING Or MF_GRAYED, IDM_FILE_SAVE_AS, lpszMenuFileSaveAs)
        Call AppendMenuW(hMenuFile, MF_SEPARATOR, 0, 0)
        Call AppendMenuW(hMenuFile, MF_STRING, IDM_APP_EXIT, lpszMenuFileExit)
    hMenuEdit = CreatePopupMenu()
    Call AppendMenuW(hMenu, MF_POPUP, hMenuEdit, lpszMenuEdit)
        Call AppendMenuW(hMenuEdit, MF_STRING Or MF_GRAYED, IDM_EDIT_UNDO, lpszMenuEditUndo)
        Call AppendMenuW(hMenuEdit, MF_STRING Or MF_GRAYED, IDM_EDIT_REDO, lpszMenuEditRedo)
        Call AppendMenuW(hMenuEdit, MF_SEPARATOR, 0, 0)
        Call AppendMenuW(hMenuEdit, MF_STRING Or MF_GRAYED, IDM_EDIT_CUT, lpszMenuEditCut)
        Call AppendMenuW(hMenuEdit, MF_STRING Or MF_GRAYED, IDM_EDIT_COPY, lpszMenuEditCopy)
        Call AppendMenuW(hMenuEdit, MF_STRING Or MF_GRAYED, IDM_EDIT_PASTE, lpszMenuEditPaste)
        Call AppendMenuW(hMenuEdit, MF_STRING Or MF_GRAYED, IDM_EDIT_CLEAR, lpszMenuEditDelete)
    hMenuView = CreatePopupMenu()
    Call AppendMenuW(hMenu, MF_POPUP, hMenuView, lpszMenuView)
        Call AppendMenuW(hMenuView, MF_STRING Or MF_CHECKED, IDM_VIEW_MATRICES, lpszMenuViewMatrices)
    hMenuDraw = CreatePopupMenu()
    Call AppendMenuW(hMenu, MF_POPUP, hMenuDraw, lpszMenuDraw)
        Call AppendMenuW(hMenuDraw, MF_STRING Or MF_CHECKED, IDM_DRAW_LINE, lpszMenuDrawLine)
        Call AppendMenuW(hMenuDraw, MF_STRING, IDM_DRAW_CIRCLE, lpszMenuDrawCircle)
    Call SetMenu(hWnd, hMenu)
    Call DrawMenuBar(hWnd)

'Right Mouse Button Menu
    hMenuRMB = CreatePopupMenu()
        Call AppendMenuW(hMenuRMB, MF_STRING Or MF_CHECKED, IDM_DRAW_LINE, lpszMenuDrawLine)
        Call AppendMenuW(hMenuRMB, MF_STRING, IDM_DRAW_CIRCLE, lpszMenuDrawCircle)
        Call AppendMenuW(hMenuRMB, MF_SEPARATOR, 0, 0)
        Call AppendMenuW(hMenuRMB, MF_STRING Or MF_GRAYED, IDM_EDIT_CUT, lpszMenuEditCut)
        Call AppendMenuW(hMenuRMB, MF_STRING Or MF_GRAYED, IDM_EDIT_COPY, lpszMenuEditCopy)
        Call AppendMenuW(hMenuRMB, MF_STRING Or MF_GRAYED, IDM_EDIT_PASTE, lpszMenuEditPaste)
End Sub

Sub DoCreateStatusBar(ByVal hWnd As LongPtr, ByVal Width As LongPtr)
    Dim idStatusBar As String
    Dim szStatusClassName As String
    Dim lpszStatusClassName As LongPtr
    'https://learn.microsoft.com/en-us/windows/win32/winauto/status-bar-control
    szStatusClassName = "msctls_statusbar32" '"STATUSCLASSNAMEW"
    lpszStatusClassName = StrPtr(szStatusClassName)
    idStatusBar = 1 'Child window identifier for Status Bar
    hwndStatusBar = CreateWindowExW(0, lpszStatusClassName, 0, SBARS_SIZEGRIP Or WS_CHILD Or WS_VISIBLE, 0, 0, 0, 0, hWnd, idStatusBar, ghInst, 0)
    If hwndStatusBar = 0 Then
        Call MsgBox("Status Bar Error: " & GetLastError())
        Exit Sub
    'Else
        'Debug.Print "hWnd =  " & hWnd
    End If
    xStatusParts(0) = 50
    xStatusParts(1) = 100
    xStatusParts(2) = CLng(Width - 50)
    xStatusParts(3) = -1
    lpStatusParts = VarPtr(xStatusParts(0))
    Call SendMessageW(hwndStatusBar, SB_SETPARTS, 4, lpStatusParts)
End Sub

Sub ClearAll()
    'Counter
    Dim i As Long
    nLine = 0
    For i = 0 To 255
        xLineStart(i) = 0
        yLineStart(i) = 0
        zLineStart(i) = 0
        xLineEnd(i) = 0
        yLineEnd(i) = 0
        zLineEnd(i) = 0
    Next i
    nCircle = 0
    For i = 0 To 255
        xCircleCenter(i) = 0
        yCircleCenter(i) = 0
        zCircleCenter(i) = 0
        rCircleRadius(i) = 0
    Next i
End Sub

Sub LoadIdentity()
    'https://learn.microsoft.com/ru-ru/windows/win32/gdiplus/-gdiplus-matrix-representation-of-transformations-about
    'View Matrix
    mtxView(0, 0) = 1
    mtxView(1, 0) = 0
    mtxView(2, 0) = 0
    mtxView(0, 1) = 0
    mtxView(1, 1) = 1
    mtxView(2, 1) = 0
    mtxView(0, 2) = 0
    mtxView(1, 2) = 0
    mtxView(2, 2) = 1
    'Inverse Matrix
    mtxInv(0, 0) = 1
    mtxInv(1, 0) = 0
    mtxInv(2, 0) = 0
    mtxInv(0, 1) = 0
    mtxInv(1, 1) = 1
    mtxInv(2, 1) = 0
    mtxInv(0, 2) = 0
    mtxInv(1, 2) = 0
    mtxInv(2, 2) = 1
    'Transformation Matrix
    mtxTrans(0, 0) = 1
    mtxTrans(1, 0) = 0
    mtxTrans(2, 0) = 0
    mtxTrans(0, 1) = 0
    mtxTrans(1, 1) = 1
    mtxTrans(2, 1) = 0
    mtxTrans(0, 2) = 0
    mtxTrans(1, 2) = 0
    mtxTrans(2, 2) = 1
End Sub

Sub ClearTransformationMatrix()
    mtxTrans(0, 0) = 1
    mtxTrans(1, 0) = 0
    mtxTrans(2, 0) = 0
    mtxTrans(0, 1) = 0
    mtxTrans(1, 1) = 1
    mtxTrans(2, 1) = 0
    mtxTrans(0, 2) = 0
    mtxTrans(1, 2) = 0
    mtxTrans(2, 2) = 1
End Sub

Sub Inverse()
    Dim detA As Double
    'For 2x2 Matrix:
    'detA = fScale * fScale
    'dxWorld = -dxScreen/fScale
    'dyWorld = -dyScreen/fScale
    
    detA = _
    mtxView(0, 0) * mtxView(1, 1) * mtxView(2, 2) - _
    mtxView(0, 0) * mtxView(1, 2) * mtxView(2, 1) - _
    mtxView(0, 1) * mtxView(1, 0) * mtxView(2, 2) + _
    mtxView(0, 1) * mtxView(1, 2) * mtxView(2, 0) + _
    mtxView(0, 2) * mtxView(1, 0) * mtxView(2, 1) - _
    mtxView(0, 2) * mtxView(1, 1) * mtxView(2, 0)

    If detA = 0 Then
        Exit Sub
    End If
    mtxInv(0, 0) = (mtxView(1, 1) * mtxView(2, 2) - mtxView(2, 1) * mtxView(1, 2)) / detA
    mtxInv(1, 0) = -(mtxView(1, 0) * mtxView(2, 2) - mtxView(2, 0) * mtxView(1, 2)) / detA
    mtxInv(2, 0) = (mtxView(1, 0) * mtxView(2, 1) - mtxView(2, 0) * mtxView(1, 1)) / detA
    mtxInv(0, 1) = -(mtxView(0, 1) * mtxView(2, 2) - mtxView(2, 1) * mtxView(0, 2)) / detA
    mtxInv(1, 1) = (mtxView(0, 0) * mtxView(2, 2) - mtxView(2, 0) * mtxView(0, 2)) / detA
    mtxInv(2, 1) = -(mtxView(0, 0) * mtxView(2, 1) - mtxView(2, 0) * mtxView(0, 1)) / detA
    mtxInv(0, 2) = (mtxView(0, 1) * mtxView(1, 2) - mtxView(1, 1) * mtxView(0, 2)) / detA
    mtxInv(1, 2) = -(mtxView(0, 0) * mtxView(1, 2) - mtxView(1, 0) * mtxView(0, 2)) / detA
    mtxInv(2, 2) = (mtxView(0, 0) * mtxView(1, 1) - mtxView(1, 0) * mtxView(0, 1)) / detA
End Sub

Sub MatrixSimpleMove()
    'mtxView = mtxView+mtxTrans
    mtxView(0, 2) = mtxView(0, 2) + mtxTrans(0, 2)
    mtxView(1, 2) = mtxView(1, 2) + mtxTrans(1, 2)
    mtxView(2, 2) = mtxView(2, 2) + mtxTrans(2, 2)
    Call Inverse
End Sub

Sub MatrixSimpleScale()
    'mtxView = mtxView*mtxTrans
    mtxView(0, 0) = mtxView(0, 0) * mtxTrans(0, 0)
    mtxView(1, 1) = mtxView(1, 1) * mtxTrans(1, 1)
    mtxView(2, 2) = mtxView(2, 2) * mtxTrans(2, 2)
    Call Inverse
End Sub

Sub MatrixMultiply()
    Dim mtxBuffer(2, 2) As Double
'Counters
    Dim i As Long
    Dim j As Long
    For i = 0 To 2
        For j = 0 To 2
            mtxBuffer(i, j) = mtxView(i, j)
        Next j
    Next i
'mtxView = mtxBuffer*mtxTrans
    mtxView(0, 0) = mtxBuffer(0, 0) * mtxTrans(0, 0) + mtxBuffer(1, 0) * mtxTrans(0, 1) + mtxBuffer(2, 0) * mtxTrans(0, 2)
    mtxView(1, 0) = mtxBuffer(0, 0) * mtxTrans(1, 0) + mtxBuffer(1, 0) * mtxTrans(1, 1) + mtxBuffer(2, 0) * mtxTrans(1, 2)
    mtxView(2, 0) = mtxBuffer(0, 0) * mtxTrans(2, 0) + mtxBuffer(1, 0) * mtxTrans(2, 1) + mtxBuffer(2, 0) * mtxTrans(2, 2)
    mtxView(0, 1) = mtxBuffer(0, 1) * mtxTrans(0, 0) + mtxBuffer(1, 1) * mtxTrans(0, 1) + mtxBuffer(2, 1) * mtxTrans(0, 2)
    mtxView(1, 1) = mtxBuffer(0, 1) * mtxTrans(1, 0) + mtxBuffer(1, 1) * mtxTrans(1, 1) + mtxBuffer(2, 1) * mtxTrans(1, 2)
    mtxView(2, 1) = mtxBuffer(0, 1) * mtxTrans(2, 0) + mtxBuffer(1, 1) * mtxTrans(2, 1) + mtxBuffer(2, 1) * mtxTrans(2, 2)
    mtxView(0, 2) = mtxBuffer(0, 2) * mtxTrans(0, 0) + mtxBuffer(1, 2) * mtxTrans(0, 1) + mtxBuffer(2, 2) * mtxTrans(0, 2)
    mtxView(1, 2) = mtxBuffer(0, 2) * mtxTrans(1, 0) + mtxBuffer(1, 2) * mtxTrans(1, 1) + mtxBuffer(2, 2) * mtxTrans(1, 2)
    mtxView(2, 2) = mtxBuffer(0, 2) * mtxTrans(2, 0) + mtxBuffer(1, 2) * mtxTrans(2, 1) + mtxBuffer(2, 2) * mtxTrans(2, 2)
    Call Inverse
End Sub

Sub ShowMatrices()
    If nMatrixShow = 1 Then
    'Show View Matrix
        szMatrix = Format(mtxView(0, 0), "0000.0000") & " | " & Format(mtxView(1, 0), "0000.0000") & " | " & Format(mtxView(2, 0), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 10, lpszMatrix, Len(szMatrix))
        szMatrix = Format(mtxView(0, 1), "0000.0000") & " | " & Format(mtxView(1, 1), "0000.0000") & " | " & Format(mtxView(2, 1), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 30, lpszMatrix, Len(szMatrix))
        szMatrix = Format(mtxView(0, 2), "0000.0000") & " | " & Format(mtxView(1, 2), "0000.0000") & " | " & Format(mtxView(2, 2), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 50, lpszMatrix, Len(szMatrix))
    'Show Inverse Matrix
        szMatrix = Format(mtxInv(0, 0), "0000.0000") & " | " & Format(mtxInv(1, 0), "0000.0000") & " | " & Format(mtxInv(2, 0), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 80, lpszMatrix, Len(szMatrix))
        szMatrix = Format(mtxInv(0, 1), "0000.0000") & " | " & Format(mtxInv(1, 1), "0000.0000") & " | " & Format(mtxInv(2, 1), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 100, lpszMatrix, Len(szMatrix))
        szMatrix = Format(mtxInv(0, 2), "0000.0000") & " | " & Format(mtxInv(1, 2), "0000.0000") & " | " & Format(mtxInv(2, 2), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 120, lpszMatrix, Len(szMatrix))
    'Show Transformation Matrix
        szMatrix = Format(mtxTrans(0, 0), "0000.0000") & " | " & Format(mtxTrans(1, 0), "0000.0000") & " | " & Format(mtxTrans(2, 0), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 150, lpszMatrix, Len(szMatrix))
        szMatrix = Format(mtxTrans(0, 1), "0000.0000") & " | " & Format(mtxTrans(1, 1), "0000.0000") & " | " & Format(mtxTrans(2, 1), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 170, lpszMatrix, Len(szMatrix))
        szMatrix = Format(mtxTrans(0, 2), "0000.0000") & " | " & Format(mtxTrans(1, 2), "0000.0000") & " | " & Format(mtxTrans(2, 2), "0000.0000")
        lpszMatrix = StrPtr(szMatrix)
        Call TextOutW(ghDC, 10, 190, lpszMatrix, Len(szMatrix))
    End If
End Sub

Sub ZoomExtents(ByVal Width As LongPtr, ByVal Height As LongPtr)
    Dim xLeft, xRight, yTop, yBottom As Double
    Dim ExtentsHor, ExtentsVer As Double
    Dim xCenterExtents, yCenterExtents As Double
    Dim Ratio, RatioVer, RatioHor As Double
    Dim xCenterClient, yCenterClient As Double
    Dim i As Integer
    'Horizontal Extents
    If nLine > 0 Then
        xLeft = xLineStart(1)
        xRight = xLineStart(1)
        For i = 1 To nLine
            xLeft = Min(xLineStart(i), xLeft)
            xLeft = Min(xLineEnd(i), xLeft)
            xRight = Max(xLineStart(i), xRight)
            xRight = Max(xLineEnd(i), xRight)
        Next i
    End If
    If nCircle > 0 Then
        If nLine = 0 Then
            xLeft = xCircleCenter(1) - rCircleRadius(1)
            xRight = xCircleCenter(1) + rCircleRadius(1)
        End If
        For i = 1 To nCircle
            xLeft = Min(xCircleCenter(i) - rCircleRadius(i), xLeft)
            xRight = Max(xCircleCenter(i) + rCircleRadius(i), xRight)
        Next i
    End If
    ExtentsHor = Abs(xRight - xLeft)
    'Horizontal Ratio
    If ExtentsHor = 0 Then
        RatioHor = 1
    Else
        RatioHor = Width / ExtentsHor
    End If
    'VerticalExtents
    If nLine > 0 Then
        yTop = yLineStart(1)
        yBottom = yLineStart(1)
        For i = 1 To nLine
            yTop = Min(yLineStart(i), yTop)
            yTop = Min(yLineEnd(i), yTop)
            yBottom = Max(yLineStart(i), yBottom)
            yBottom = Max(yLineEnd(i), yBottom)
        Next i
    End If
    If nCircle > 0 Then
        If nLine = 0 Then
            yTop = yCircleCenter(1) - rCircleRadius(1)
            yBottom = yCircleCenter(1) + rCircleRadius(1)
        End If
        For i = 1 To nCircle
            yTop = Min(yCircleCenter(i) - rCircleRadius(i), yTop)
            yBottom = Max(yCircleCenter(i) + rCircleRadius(i), yBottom)
        Next i
    End If
    ExtentsVer = Abs(yBottom - yTop)
    'Vertical Ratio
    If ExtentsVer = 0 Then
        RatioVer = 1
    Else
        RatioVer = Height / ExtentsVer
    End If
    'Compare Ratios
    If RatioHor = 1 Then
        If RatioVer = 1 Then
            Ratio = 1
        Else
            Ratio = RatioVer
        End If
    Else
        If RatioVer = 1 Then
            Ratio = RatioHor
        Else
            Ratio = Min(RatioHor, RatioVer)
        End If
    End If
    'Move Center of Drawing to Initial Coordinates
    xCenterExtents = xLeft + (xRight - xLeft) / 2
    yCenterExtents = yTop + (yBottom - yTop) / 2
    Call LoadIdentity
        mtxTrans(0, 2) = -xCenterExtents
        mtxTrans(1, 2) = -yCenterExtents
    Call MatrixMultiply
    'Mutiply Transform Matrix by the Less Ratio
    Call ClearTransformationMatrix
        mtxTrans(0, 0) = Ratio
        mtxTrans(1, 1) = Ratio
    Call MatrixMultiply
    'Move Projection to Center of Client Area
    xCenterClient = Width / 2
    yCenterClient = Height / 2
    Call ClearTransformationMatrix
        mtxTrans(0, 2) = xCenterClient
        mtxTrans(1, 2) = yCenterClient
    Call MatrixMultiply
End Sub

Function Max(ByVal n1 As Double, ByVal n2 As Double)
    If n2 > n1 Then
        Max = n2
    Else
        Max = n1
    End If
End Function

Function Min(ByVal n1 As Double, ByVal n2 As Double)
    If n2 < n1 Then
        Min = n2
    Else
        Min = n1
    End If
End Function



