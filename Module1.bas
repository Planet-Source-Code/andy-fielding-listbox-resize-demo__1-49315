Attribute VB_Name = "Module1"
Option Explicit

' Declare API function to get the user's system-wide scrollbar-width setting
' (set in Display Properties)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXVSCROLL = 2

' Declare API function to detect scrollbar in list box
' (By the way, this can be used to detect scrollbars in any control that can
' have them)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Const GWL_STYLE = (-16)
Const WS_VSCROLL = &H200000
Const WS_HSCROLL = &H100000

Function HasVerticalScrollbar(ctrl As Control) As Boolean

  ' If the control whose name is passed to this function (in this case,
  ' the list box) has a vertical scrollbar, return "True"
  HasVerticalScrollbar = (GetWindowLong(ctrl.hwnd, GWL_STYLE) And WS_VSCROLL)

End Function

Function HasHorizontalScrollbar(ctrl As Control) As Boolean

  ' We don't need to detect a horizontal scrollbar in this demo, but
  ' for your reference, here's the code to do that

  HasHorizontalScrollbar = (GetWindowLong(ctrl.hwnd, GWL_STYLE) And WS_HSCROLL)

End Function

