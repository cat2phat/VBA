Option Compare Database
Option Explicit

Private Declare Function GetWindowPlacement Lib "User32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hwnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Type POINTAPI
    X As Long
    y As Long
End Type

Type RECT '****'
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type WINDOWPLACEMENT
    Length As Long
    flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT '****'
End Type

'*****
'MODIFY WINDOW DIMENSION FUNCTION CALLS
'*****
Sub ResizeWindow(wWidth As Variant, hHeight As Variant)
    Call ChangeWindowDimension(, , wWidth, hHeight, 1, 1)
End Sub

Sub MoveWindowToTopY()
    Call ChangeWindowDimension(, 0, , , 1, 1)
End Sub

Sub MoveWindowToCenterY()
    Dim desktopCenterY As Integer
    Dim currWindowPos As RECT
    Dim currWindowHeight As Integer
    Dim centeredY As Integer
    
    desktopCenterY = GetPixelY / 2 ' getpixely found in mdl_OperatingSystemInfo
    
    currWindowPos = GetWindowPos(Access.Application.hWndAccessApp)
    currWindowHeight = currWindowPos.Bottom - currWindowPos.Top
    
    centeredY = desktopCenterY - (currWindowHeight / 2)
    
    Call ChangeWindowDimension(, centeredY, , , 1, 1)
End Sub

Sub MoveWindowToBottomY()
    Dim currWindowPos As RECT
    Dim currWindowHeight As Integer
    Dim bottomY As Integer
    
    currWindowPos = GetWindowPos(Access.Application.hWndAccessApp)
    currWindowHeight = currWindowPos.Bottom - currWindowPos.Top
    
    bottomY = GetPixelY - currWindowHeight ' getpixely found in mdl_OperatingSystemInfo
    
    If bottomY < 0 Then: bottomY = 0
    
    Call ChangeWindowDimension(, bottomY, , , 1, 1)
End Sub

Sub MoveWindowToLeftX()
    Call ChangeWindowDimension(0, , , , 1, 1)
End Sub

Sub MoveWindowToCenterX()
    Dim desktopCenterX As Integer
    Dim currWindowPos As RECT
    Dim currWindowWidth As Integer
    Dim centeredX As Integer
    
    desktopCenterX = GetPixelX / 2 ' getpixelx found in mdl_OperatingSystemInfo
    
    currWindowPos = GetWindowPos(Access.Application.hWndAccessApp)
    currWindowWidth = currWindowPos.Right - currWindowPos.Left
    
    centeredX = desktopCenterX - (currWindowWidth / 2)
    
    Call ChangeWindowDimension(centeredX, , , , 1, 1)
End Sub

Sub MoveWindowToRightX()
    Dim currWindowPos As RECT
    Dim currWindowWidth As Integer
    Dim centeredX As Integer
    
    currWindowPos = GetWindowPos(Access.Application.hWndAccessApp)
    currWindowWidth = currWindowPos.Right - currWindowPos.Left
    
    centeredX = GetPixelX - currWindowWidth ' getpixelx found in mdl_OperatingSystemInfo
    
    Call ChangeWindowDimension(centeredX, , , , 1, 1)
End Sub

Sub MoveWindowToTheCenter()
    MoveWindowToCenterX
    MoveWindowToCenterY
End Sub

'*****
'MODIFY WINDOW DIMENSION FUNCTIONS
'*****

'Changes the dimension and placement of the window in pixels
Public Sub ChangeWindowDimension( _
    Optional xLeft As Variant, _
    Optional yTop As Variant, _
    Optional wWidth As Variant, _
    Optional hHeight As Variant, _
    Optional bConsiderDPI As Boolean = True, _
    Optional bConsiderTaskbar As Boolean = True)
    
    Dim lhwnd As Long
    Dim apiRECT As RECT
    Dim currWindowPos As RECT
    Dim dpiWidth As Integer
    Dim dpiHeight As Integer
    Dim dpi As Integer
    Dim dpiMultiplier As Double
    
    lhwnd = Access.Application.hWndAccessApp
    
    currWindowPos = GetWindowPos(lhwnd)
    
    If IsMissing(xLeft) Then: xLeft = currWindowPos.Left
    If IsMissing(yTop) Then: yTop = currWindowPos.Top
    
    If bConsiderDPI Then
        dpi = GetDPI
        dpiMultiplier = dpi / 96
    End If
    
    If Not IsMissing(wWidth) Then
        dpiWidth = wWidth * dpiMultiplier
    Else
        dpiWidth = currWindowPos.Right - currWindowPos.Left
    End If
    
    If Not IsMissing(hHeight) Then
        dpiHeight = hHeight * dpiMultiplier
    Else
        dpiHeight = currWindowPos.Bottom - currWindowPos.Top
    End If
    
    If bConsiderTaskbar Then
        Call GetDesktopArea(apiRECT) ' found in mdl_OperatingSystemInfo
    Else
        apiRECT.Left = 0
        apiRECT.Top = 0
        apiRECT.Bottom = 0
        apiRECT.Right = 0
    End If
    
    apiRECT.Left = apiRECT.Left + xLeft
    apiRECT.Top = apiRECT.Top + yTop
    If apiRECT.Bottom - (apiRECT.Top + dpiHeight) < 0 Then
        apiRECT.Top = apiRECT.Top - ((apiRECT.Top + dpiHeight) - apiRECT.Bottom)
    End If
    If apiRECT.Right - (apiRECT.Left + dpiWidth) < 0 Then
        apiRECT.Left = apiRECT.Left - ((apiRECT.Left + dpiWidth) - apiRECT.Right)
    End If

    If apiRECT.Top < 0 Then: apiRECT.Top = 0
    If apiRECT.Left < 0 Then: apiRECT.Left = 0

    MoveWindow lhwnd, apiRECT.Left, apiRECT.Top, dpiWidth, dpiHeight, 1
End Sub

'Returns the dimensions and placement of the window
Public Function GetWindowPos(ByVal hwnd As Long) As RECT
    Dim wp As WINDOWPLACEMENT
    
    wp.Length = Len(wp)
    Call GetWindowPlacement(hwnd, wp)
    GetWindowPos = wp.rcNormalPosition
End Function


