Attribute VB_Name = "Module1"
'// does it really get any easier... special thanks to allapi.com
'// and its easy to mod to intercept any WM mouse event
'// just remember, dont exit the form by stopping it in the ide... knuckle head
'// yeah... and meatwad says hi
Option Explicit
Private Const TME_CANCEL = &H80000000
Private Const TME_HOVER = &H1&
Private Const TME_LEAVE = &H2&
Private Const TME_NONCLIENT = &H10&
Private Const TME_QUERY = &H40000000
Private Const WM_MOUSELEAVE = &H2A3&
Private Const WM_LBUTTONDBLCLK As Integer = &H203
Private Const WM_LBUTTONDOWN As Integer = &H201
Private Const WM_LBUTTONUP  As Integer = &H202
Private Const WM_MBUTTONDBLCLK  As Integer = &H209
Private Const WM_MBUTTONDOWN  As Integer = &H207
Private Const WM_MBUTTONUP  As Integer = &H208
Private Const WM_MOUSEACTIVATE  As Integer = &H21
Private Const WM_MOUSEFIRST  As Integer = &H200
Private Const WM_MOUSELAST  As Integer = &H209
Private Const WM_MOUSEMOVE  As Integer = &H200
Private Const WM_RBUTTONDBLCLK  As Integer = &H206
Private Const WM_RBUTTONDOWN  As Integer = &H204
Private Const WM_RBUTTONUP  As Integer = &H205

Private Type TMET
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Declare Function TrackMouseEvent2 Lib "comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TMET) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC As Long = (-4)
Private PrevProc As Long
Private ET As TMET

Public Sub Hook(F As Variant)

    PrevProc = SetWindowLong(F.hwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Private Sub LBD(object As Variant)

    On Error Resume Next
        object.Caption = "Left Down"
        object.BackColor = vbYellow

End Sub

Private Sub ML(object As Variant)

    On Error Resume Next
        object.Caption = "mouseleave"
        object.BackColor = vbRed

End Sub

Private Sub mouseMoveHook(object As Variant)

    On Error Resume Next
        ET.cbSize = Len(ET)
        ET.hwndTrack = object.hwnd
        ET.dwFlags = TME_LEAVE
        TrackMouseEvent2 ET
        object.BackColor = vbWhite
        object.Caption = "mousemove"

End Sub

Public Sub UnHook(F As Variant)

    SetWindowLong F.hwnd, GWL_WNDPROC, PrevProc

End Sub

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Dim control As control

    With Form1
        If uMsg = WM_MOUSEMOVE Then
            For Each control In Form1.Controls
                If TypeOf control Is Shape Then
                  Else 'NOT TYPEOF...

                    If hwnd = control.hwnd Then
                        mouseMoveHook control
                    End If
                End If
            Next control

          ElseIf uMsg = WM_MOUSELEAVE Then 'NOT UMSG...
            For Each control In Form1.Controls
                If TypeOf control Is Shape Then
                  Else 'NOT TYPEOF...

                    If hwnd = control.hwnd Then
                        ML control
                    End If
                End If
            Next control

          ElseIf uMsg = WM_LBUTTONDOWN Then 'NOT UMSG...
            For Each control In Form1.Controls
                If TypeOf control Is Shape Then
                  Else 'NOT TYPEOF...

                    If hwnd = control.hwnd Then
                        LBD control
                    End If
                End If
            Next control
        End If

    End With 'FORM1
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)

End Function

