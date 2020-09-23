Attribute VB_Name = "HookClipboard"
Option Explicit

'KPD-Team 1999
'URL: http://www.allapi.net/
'E-Mail: KPDTeam@Allapi.net
'These routines are explained in our subclassing tutorial.
'http://www.allapi.net/vbtutor/subclass.htm

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetClipboardViewer Lib "user32" (ByVal hWnd As Long) As Long

Public Const WM_DRAWCLIPBOARD = &H308
Public Const GWL_WNDPROC = (-4)

Public PrevProc As Long
Public Sub HookForm(F As Form)
    PrevProc = SetWindowLong(F.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHookForm(F As Form)
    SetWindowLong F.hWnd, GWL_WNDPROC, PrevProc
End Sub
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim sTemp As String

WindowProc = CallWindowProc(PrevProc, hWnd, uMsg, wParam, lParam)

If uMsg = WM_DRAWCLIPBOARD Then
    
    sTemp = Clipboard.GetText(vbCFText)
    If sTemp <> "" Then frmClipList.lstLocal.AddItem sTemp, 0

End If

End Function
