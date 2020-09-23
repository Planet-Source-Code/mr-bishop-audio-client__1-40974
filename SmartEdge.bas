Attribute VB_Name = "SmartEdge"
' ----------------------------------------------------
' Module SmartEdge
'
' Version... 1.2
' Date...... 22 April 2001
'
' Copyright (C) 2001 Andrés Pons (andres@vbsmart.com)
' ----------------------------------------------------

Option Explicit

'Different border styles we can apply to any window:
Public Enum sedBorderStyle
    sedNone
    sedSunken
    sedSunkenOuter
    sedRaised
    sedRaisedInner
    sedBump
    sedEtched
End Enum

Private Enum sedBorderWidth
    sbwNone
    sbwSingle
    sbwDouble
End Enum

Private Const SED_OLDPROC = "SED_OLDPROC"
Private Const SED_OLDGWLSTYLE = "SED_OLDGWLSTYLE"
Private Const SED_OLDGWLEXSTYLE = "SED_OLDGWLEXSTYLE"
Private Const SED_BORDERS = "SED_BORDERS"

'API declarations:
Private Const WM_NCPAINT = &H85

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_SHOWWINDOW = &H40

Private Const BDR_INNER = &HC
Private Const BDR_OUTER = &H3
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const GWL_WNDPROC = (-4)
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const WS_THICKFRAME = &H40000
Private Const WS_BORDER = &H800000
Private Const WS_EX_WINDOWEDGE = &H100&
Private Const WS_EX_CLIENTEDGE = &H200&
Private Const WS_EX_STATICEDGE = &H20000

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hWnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hWnd As Long, ByVal lpString As String) As Long

' ------------------------------------------------------------------------
' Sub pWindowProc()
'
'      Once a window has been subclassed, this procedure will
'      receive all messages instead of its original windowproc.
'      We can now filter the WM_NCPAINT message and use our own
'      functions to draw the window edge.
'
Private Function pWindowProc( _
    ByVal hWnd As Long, _
    ByVal uMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long
    
    Select Case uMsg 'Select the message uMsg and only modify WM_NCPAINT...
    
        Case WM_NCPAINT 'Message sent to a window when its non-client area needs to be re-drawn ...
            
            'Call our own drawing function...
            pDrawBorder hWnd, wParam, GetProp(hWnd, SED_BORDERS)
        
        Case Else
            'All other messages should be sent to the original windowproc...
            pWindowProc = CallWindowProc(GetProp(hWnd, SED_OLDPROC), hWnd, uMsg, wParam, lParam)
            
    End Select
    
End Function

' ------------------------------------------------------------------------
' Sub pDrawBorder()
'
'      The purpose of this procedure is to draw the window edge.
'      Called by pWindowProc when receives the message WM_NCPAINT.
'
Private Sub pDrawBorder(ByVal hWnd As Long, ByVal wParam As Long, ByVal lBorderType As sedBorderStyle)

    Dim lRet As Long
    Dim lMode As Long
    Dim hDC As Long
    Dim Rec As RECT
    
    'There's no drawing needed when there's no border assigned...
    If lBorderType = sedNone Then Exit Sub
    
    'Get a device context for this window handle...
    hDC = GetWindowDC(hWnd)
    
    'Get the RECT that contains the window...
    lRet = GetWindowRect(hWnd, Rec)
    
    'Transform from screen coordinates to client coordinates...
    Rec.Right = Rec.Right - Rec.Left
    Rec.Bottom = Rec.Bottom - Rec.Top
    Rec.Left = 0
    Rec.Top = 0

    'Choose the drawing flags based on the selected border style...
    lMode = 0
    Select Case lBorderType
        Case sedRaised
            lMode = BDR_RAISED
        Case sedRaisedInner
            lMode = BDR_RAISEDINNER
        Case sedSunken
            lMode = BDR_SUNKEN
        Case sedSunkenOuter
            lMode = BDR_SUNKENOUTER
        Case sedEtched
            lMode = BDR_SUNKENOUTER Or BDR_RAISEDINNER
        Case sedBump
            lMode = BDR_SUNKENINNER Or BDR_RAISEDOUTER
    End Select
    
    'Draw the window border by using the API DrawEdge...
    lRet = DrawEdge(hDC, Rec, lMode, BF_RECT)
    
    'Release the device context...
    lRet = ReleaseDC(hWnd, hDC)

End Sub

' ------------------------------------------------------------------------
' Function EdgeSubClass()
'
'      This function allows you to modify the border style of any window
'      Example: retVal = EdgeSubClass(Picture1.Hwnd,sedRaisedInner)
'
Public Function EdgeSubClass(ByVal hWnd As Long, ByVal eBorderStyle As sedBorderStyle) As Boolean
    
    Dim lRet As Long
    
    'Check whether the window was already subclassed
    'and get the original windowproc...
    lRet = GetProp(hWnd, SED_OLDPROC)

    If lRet <> 0 Then
        'Unsubclass the window...
        SetWindowLong hWnd, GWL_WNDPROC, lRet
    Else
        'Store the window style (only the first time we subclass the window)...
        SetProp hWnd, SED_OLDGWLSTYLE, GetWindowLong(hWnd, GWL_STYLE)
        SetProp hWnd, SED_OLDGWLEXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE)
    End If
    
    'Change to the window border that best suits our drwaing requirements...
    pSetBorder hWnd, eBorderStyle
    
    'Subclass the window...
    lRet = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf pWindowProc)

    'Store the original windowproc and the new border style...
    SetProp hWnd, SED_OLDPROC, lRet
    SetProp hWnd, SED_BORDERS, CLng(eBorderStyle)

    'Refresh the window (this forces Windows to send a WM_NCPAINT message)...
    SetWindowPos hWnd, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or _
        SWP_NOSIZE Or _
        SWP_NOOWNERZORDER Or _
        SWP_NOZORDER Or _
        SWP_FRAMECHANGED

    EdgeSubClass = (lRet <> 0)

End Function

' ------------------------------------------------------------------------
' Function EdgeUnSubClass()
'
'      This function restores a subclassed window to its original status
'      Example: retVal = EdgeSubClass(Picture1.Hwnd)
'
Public Function EdgeUnSubClass(ByVal hWnd As Long) As Boolean
    
    Dim lRet As Long

    'Get the original windowproc for this window...
    lRet = GetProp(hWnd, SED_OLDPROC)
    
    If lRet <> 0 Then
        'Unsubclass the window by assigning the original windowproc...
        lRet = SetWindowLong(hWnd, GWL_WNDPROC, lRet)
        
        'Restore the original window styles...
        SetWindowLong hWnd, GWL_STYLE, GetProp(hWnd, SED_OLDGWLSTYLE)
        SetWindowLong hWnd, GWL_EXSTYLE, GetProp(hWnd, SED_OLDGWLEXSTYLE)
        
        'Refresh the window (sends message WM_NCPAINT)...
        SetWindowPos hWnd, 0, 0, 0, 0, 0, _
            SWP_NOMOVE Or _
            SWP_NOSIZE Or _
            SWP_NOOWNERZORDER Or _
            SWP_NOZORDER Or _
            SWP_FRAMECHANGED
        
        'Remove all stored information for this window...
        RemoveProp hWnd, SED_OLDPROC
        RemoveProp hWnd, SED_OLDGWLSTYLE
        RemoveProp hWnd, SED_OLDGWLEXSTYLE
        RemoveProp hWnd, SED_BORDERS
        
    End If
    
    EdgeUnSubClass = (lRet <> 0)

End Function

' ------------------------------------------------------------------------
' Sub pSetBorder()
'
'      Different border styles have different widths.
'      To draw the border edge we first need to make sure that
'      the window border width will be correct.
'
Private Sub pSetBorder(ByVal hWnd As Long, ByVal eBorderStyle As sedBorderStyle)

    Dim pWidth As sedBorderWidth
    
    'Depending on the border style we want to draw,
    'we need a width of 0, 1 or 2 pixels...
    Select Case eBorderStyle
        Case sedNone
            pWidth = sbwNone
        Case sedRaised
            pWidth = sbwDouble
        Case sedRaisedInner
            pWidth = sbwSingle
        Case sedSunken
            pWidth = sbwDouble
        Case sedSunkenOuter
            pWidth = sbwSingle
        Case sedEtched
            pWidth = sbwDouble
        Case sedBump
            pWidth = sbwDouble
    End Select
    
    'Change the border style depending on the width...
    Select Case pWidth
        Case sbwNone
            pWinStyleNeg hWnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hWnd, GWL_EXSTYLE, WS_EX_STATICEDGE Or WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
        Case sbwSingle
            pWinStyleNeg hWnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hWnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE Or WS_EX_WINDOWEDGE
            pWinStyleAdd hWnd, GWL_EXSTYLE, WS_EX_STATICEDGE
        Case sbwDouble
            pWinStyleNeg hWnd, GWL_STYLE, WS_BORDER Or WS_THICKFRAME
            pWinStyleNeg hWnd, GWL_EXSTYLE, WS_EX_STATICEDGE Or WS_EX_WINDOWEDGE
            pWinStyleAdd hWnd, GWL_EXSTYLE, WS_EX_CLIENTEDGE
    End Select
    
    'Refresh the window (sends message WM_NCPAINT)...
    SetWindowPos hWnd, 0, 0, 0, 0, 0, _
        SWP_NOMOVE Or _
        SWP_NOSIZE Or _
        SWP_NOOWNERZORDER Or _
        SWP_NOZORDER Or _
        SWP_FRAMECHANGED
        
End Sub

' ------------------------------------------------------------------------
' Sub pWinStyleAdd()
'
Private Sub pWinStyleAdd(ByVal hWnd As Long, ByVal lStyle As Long, ByVal lFlags As Long)
    
    'Add flags to the window style
    SetWindowLong hWnd, lStyle, GetWindowLong(hWnd, lStyle) Or lFlags
    
End Sub

' ------------------------------------------------------------------------
' Sub pWinStyleNeg()
'
Private Sub pWinStyleNeg(ByVal hWnd As Long, ByVal lStyle As Long, ByVal lFlags As Long)
    
    'Remove flags from the window style
    SetWindowLong hWnd, lStyle, GetWindowLong(hWnd, lStyle) And Not lFlags
    
End Sub

