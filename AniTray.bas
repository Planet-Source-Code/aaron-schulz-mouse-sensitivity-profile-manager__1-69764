Attribute VB_Name = "modTrayManager"
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205

Global Const WM_MOUSEMOVE = &H200
Global Const NIM_ADD = 0&
Global Const NIM_DELETE = 2&
Global Const NIM_MODIFY = 1&
Global Const NIF_ICON = 2&
Global Const NIF_MESSAGE = 1&
Global Const ABM_GETTASKBARPOS = &H5&

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type APPBARDATA
        cbSize As Long
        hwnd As Long
        uCallbackMessage As Long
        uEdge As Long
        rc As RECT
        lParam As Long
End Type

Global Notify As NOTIFYICONDATA
Global BarData As APPBARDATA

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, pData As APPBARDATA) As Long

Sub modIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    Dim Result As Long
    Notify.cbSize = 88&
    Notify.hwnd = Form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_MODIFY, Notify)
End Sub

Sub AddIcon(Form1 As Form, IconID As Long, Icon As Object, ToolTip As String)
    Dim Result As Long
    BarData.cbSize = 36&
    Result = SHAppBarMessage(ABM_GETTASKBARPOS, BarData)
    Notify.cbSize = 88&
    Notify.hwnd = Form1.hwnd
    Notify.uID = IconID
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.hIcon = Icon
    Notify.szTip = ToolTip & Chr$(0)
    Result = Shell_NotifyIcon(NIM_ADD, Notify)
End Sub

Sub delIcon(IconID As Long)
    Dim Result As Long
    Notify.uID = IconID
    Result = Shell_NotifyIcon(NIM_DELETE, Notify)
End Sub




