VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Azza's Mouse Profile Manager"
   ClientHeight    =   1440
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.PictureBox PictureHelp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   4200
      Picture         =   "frmMain.frx":E968
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   960
      Width           =   360
   End
   Begin VB.Timer TimerUpdateMenu 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   960
   End
   Begin VB.HScrollBar HScrollMouseSpeed 
      Height          =   375
      Left            =   1320
      Max             =   20
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.ComboBox ComboUsers 
      Height          =   315
      ItemData        =   "frmMain.frx":F052
      Left            =   1320
      List            =   "frmMain.frx":F059
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Mouse Speed:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "User:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   150
      Width           =   375
   End
   Begin VB.Label LabelStatus 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4575
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu Users 
         Caption         =   "Users"
         Index           =   0
      End
      Begin VB.Menu Break 
         Caption         =   "-"
      End
      Begin VB.Menu Help 
         Caption         =   "Help"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Small program to set mouse speed profiles for different users
'My kids and I want different mouse speeds when we use the computer - thought this might be a decent solution

'uses Message Alerter (http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=4503&lngWId=1)
'to handle tray icon functions.

'mouse set/get speed API declarations
Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
    ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Const SPI_SETMOUSESPEED = 113
Const SPI_GETMOUSESPEED = 112

'listbox/combobox search API declarations
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Const LB_FINDSTRING = &H18F
Const LB_FINDSTRINGEXACT = &H1A2
Const CB_FINDSTRING = &H14C
Const CB_FINDSTRINGEXACT = &H158

'API's for getting taskbar size
'from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=23892&lngWId=1
Private Const SPI_GETWORKAREA = 48
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'trayicon variables
Public IconObject As Object

Private Sub SetMouseSpeed(lngSpeed As Long)
    'mouse speed range 0-20
    If lngSpeed < 0 Then
        lngSpeed = 0
    ElseIf lngSpeed > 20 Then
        lngSpeed = 20
    End If

    SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal lngSpeed, 0
End Sub

Private Function GetMouseSpeed() As Long
    Dim Speed As Long
    ' note that Speed is passed ByRef
    SystemParametersInfo SPI_GETMOUSESPEED, 0, Speed, 0
    GetMouseSpeed = Speed
End Function


Private Sub ComboUsers_LostFocus()
    Dim lngIndex As Long

    'check if we have this user name
    lngIndex = ListBoxFindString(ComboUsers, ComboUsers.Text)
    If lngIndex = -1 Then
        'add user to popup menu
        RemoveStatusFromMenu
        Load Users(Users.Count)
        Users(Users.Count - 1).Caption = ComboUsers.Text
        
        'add new user
        ComboUsers.AddItem ComboUsers.Text

        'add default value for new user
        HScrollMouseSpeed.Value = 10
        UpdateUserSpeed ComboUsers.Text, HScrollMouseSpeed.Value
        

        UpdateStatus ComboUsers.Text
    Else
        'go to selected user
        ComboUsers.ListIndex = lngIndex
    End If
    
    'remember currently selected user
    SaveSetting "Azza's Mouse Settings", "Current", "User", ComboUsers.Text
    
    UpdateStatus ComboUsers.Text
End Sub

Private Sub ComboUsers_Click()
    SelectUser ComboUsers.Text

    'set this user as default
    SaveSetting "Azza's Mouse Settings", "Current", "User", ComboUsers.Text
End Sub

Private Sub Help_Click()
    Dim strText As String
    
    strText = "Azza's Mouse Manager" & vbNewLine & vbNewLine & _
    "A simple program to allow easy selection between mouse speed profiles for different users." & vbNewLine & _
    "Left-Click on the traybar icon to open the main screen." & vbNewLine & _
    "For each user profile, simply type a user name into the selection box, and then select an associated mouse speed." & vbNewLine & _
    "Press <Delete> to remove a user." & vbNewLine & vbNewLine & _
    "For easy profile selection, right-click on the traybar icon."
    
    MsgBox strText
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIndex As Long
    
    If KeyCode = 46 Then
        'indicated to delete user
                           
        'don't allow deletion of default user
        If ComboUsers.Text <> "<default>" Then

            'make sure combobox entry is valid
            lngIndex = ListBoxFindString(ComboUsers, ComboUsers.Text)
            If lngIndex <> -1 Then
                DeleteSetting "Azza's Mouse Settings", "User", ComboUsers.Text
                ComboUsers.RemoveItem (lngIndex)
                
                'select default as new user
                ComboUsers.ListIndex = 0
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strCurrentUser As String
    Dim varSettings As Variant
    
    Me.Left = -100000
    
    'set trayicon
    Set IconObject = frmMain.Icon
    AddIcon frmMain, IconObject.Handle, IconObject, "Azza's Mouse Manager"

    
    'get list of users
    varSettings = GetAllSettings("Azza's Mouse Settings", "User")
    'populate user list to combo box
    If IsEmpty(varSettings) = False Then
    For i = 0 To UBound(varSettings)
        If varSettings(i, 0) <> "<default>" Then
            ComboUsers.AddItem varSettings(i, 0)
        End If

        'add user to menu
        If i > 0 Then
            Load Users(i)
        End If
        Users(i).Caption = varSettings(i, 0)
    Next i
    End If
    
    'get current user
    strCurrentUser = GetSetting("Azza's Mouse Settings", "Current", "User", "<default>")
    If ListBoxFindString(ComboUsers, strCurrentUser) = -1 Then
        'last user not found - user default
        strCurrentUser = "<default>"
    End If
    
    'setup user
    SelectUser strCurrentUser
        
    ComboUsers.ListIndex = ListBoxFindString(ComboUsers, strCurrentUser)

    UpdateStatus strCurrentUser
    
    Me.Hide
End Sub

Private Sub SelectUser(strCurrentUser As String)
    Dim lngSpeed As Long
    
    'get mouse speed for current user
    lngSpeed = Val(GetSetting("Azza's Mouse Settings", "User", strCurrentUser, -1))
    If lngSpeed = -1 Then
        'no setting saved - probably first time program loaded
        
        'get current mouse speed
        lngSpeed = GetMouseSpeed
        
        'save <default> setting with current mouse speed
        UpdateUserSpeed strCurrentUser, lngSpeed

    End If
        
    'set scroll bar to user mouse speed
    HScrollMouseSpeed.Value = lngSpeed
    
    UpdateStatus strCurrentUser
End Sub

Private Sub UpdateUserSpeed(strUserName As String, lngUserSpeed As Long)
    'remember mouse speed setting for this user
    SaveSetting "Azza's Mouse Settings", "User", strUserName, lngUserSpeed
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'clean up trayicon
    delIcon IconObject.Handle
    delIcon frmMain.Icon.Handle
End Sub

Private Sub HScrollMouseSpeed_Change()
    SetMouseSpeed HScrollMouseSpeed.Value
    
    UpdateStatus ComboUsers.Text
End Sub

Private Sub HScrollMouseSpeed_Scroll()
    SetMouseSpeed HScrollMouseSpeed.Value
    
    'remember mouse speed for this user
    UpdateUserSpeed ComboUsers.Text, HScrollMouseSpeed.Value
    
    UpdateStatus ComboUsers.Text
End Sub

Private Sub UpdateStatus(strUser As String)
    LabelStatus.Caption = "User: " & strUser & "   Current Speed: " & (HScrollMouseSpeed.Value * 5) & "%"

    TimerUpdateMenu.Enabled = True
End Sub

Private Sub OutputStatusToMenu()
    'update status in menu
    RemoveStatusFromMenu

    'add status message to menu
    Load Users(Users.Count)
    Users(Users.Count - 1).Caption = "-"
    Load Users(Users.Count)
    Users(Users.Count - 1).Caption = LabelStatus.Caption
End Sub

Private Sub RemoveStatusFromMenu()
    'looks for and removes status from menu
    
    'look for seperator - the current status will be below seperator
    For i = Users.Count - 1 To 0 Step -1
        If Users.Item(i).Caption = "-" Then
            Do Until Users.Count - 1 < i
                'seperator found - delete all from here on
                Unload Users.Item(Users.Count - 1)
            Loop
            Exit For
        End If
    Next i
End Sub

'use SendMessage to search ListBox/ComboBox
'from http://www.devx.com/vb2themax/Tip/19121
Function ListBoxFindString(ctrl As Control, ByVal search As String, _
    Optional startIndex As Long = -1, Optional ExactMatch As Boolean) As Long
    Dim uMsg As Long
    If TypeOf ctrl Is ListBox Then
        uMsg = IIf(ExactMatch, LB_FINDSTRINGEXACT, LB_FINDSTRING)
    ElseIf TypeOf ctrl Is ComboBox Then
        uMsg = IIf(ExactMatch, CB_FINDSTRINGEXACT, CB_FINDSTRING)
    Else
        Exit Function
    End If
    ListBoxFindString = SendMessage(ctrl.hwnd, uMsg, startIndex, ByVal search)
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Message As Long
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
        Case WM_LBUTTONUP:
            'show whole interface - bottom right
            Me.WindowState = vbNormal
            Me.Left = Screen.Width - Me.Width
            Me.Top = Screen.Height - Me.Height - GetTaskbarHeight
            Me.Show
        Case WM_RBUTTONUP:
            'show user list
            Me.Hide
            Me.PopupMenu Popup

    End Select
End Sub

Public Function GetTaskbarHeight() As Integer
    Dim lRes As Long
    Dim rectVal As RECT
    
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, rectVal, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - rectVal.Bottom) * Screen.TwipsPerPixelX
End Function

Private Sub PictureHelp_Click()
    Help_Click
End Sub

Private Sub TimerUpdateMenu_Timer()
    TimerUpdateMenu.Enabled = False
    OutputStatusToMenu
End Sub

Private Sub Users_Click(Index As Integer)
    SelectUser Users.Item(Index).Caption

    'set this user as default
    SaveSetting "Azza's Mouse Settings", "Current", "User", Users.Item(Index).Caption

End Sub
