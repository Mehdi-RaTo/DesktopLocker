VERSION 5.00
Begin VB.Form formMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00333333&
   BorderStyle     =   0  'None
   Caption         =   "formMain"
   ClientHeight    =   5940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   FillColor       =   &H00CCCCCC&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00CCCCCC&
   Icon            =   "formMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer timer 
      Interval        =   1000
      Left            =   360
      Top             =   5280
   End
   Begin VB.Frame frameDeveloper 
      Appearance      =   0  'Flat
      BackColor       =   &H00333333&
      Caption         =   "Developer"
      ForeColor       =   &H00CCCCCC&
      Height          =   855
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   4455
      Begin VB.Label labelDevUrl 
         Alignment       =   2  'Center
         BackColor       =   &H00333333&
         Caption         =   "https://github.com/Mehdi-RaTo"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label labelDevName 
         Alignment       =   2  'Center
         BackColor       =   &H00333333&
         Caption         =   "Mehdi-RaTo"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frameStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00333333&
      Caption         =   "Status"
      ForeColor       =   &H00CCCCCC&
      Height          =   615
      Left            =   360
      TabIndex        =   12
      Top             =   3480
      Width           =   4455
      Begin VB.Label labelStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00333333&
         Caption         =   "Unlocked"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame frameSettings 
      Appearance      =   0  'Flat
      BackColor       =   &H00333333&
      Caption         =   "Settings"
      Enabled         =   0   'False
      ForeColor       =   &H00CCCCCC&
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   4455
      Begin VB.CheckBox checkStartup 
         Appearance      =   0  'Flat
         BackColor       =   &H00333333&
         Caption         =   "Startup"
         Enabled         =   0   'False
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox inputNewPasswordConfirm 
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "•"
         TabIndex        =   3
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox inputNewPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "•"
         TabIndex        =   1
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label buttonUninstall 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Uninstall"
         ForeColor       =   &H008080FF&
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label buttonCancel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   1920
         TabIndex        =   7
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label buttonSave 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   1920
         TabIndex        =   6
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label labelNewPasswordConfirm 
         BackColor       =   &H00333333&
         Caption         =   "Confirm New Password"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label labelNewPassword 
         BackColor       =   &H00333333&
         Caption         =   "New Password"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
   End
   Begin VB.Frame frameEnter 
      Appearance      =   0  'Flat
      BackColor       =   &H00333333&
      Enabled         =   0   'False
      ForeColor       =   &H00CCCCCC&
      Height          =   3135
      Left            =   360
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   4455
      Begin VB.TextBox inputEnterPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "•"
         TabIndex        =   9
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label buttonSettings 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Settings"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   1920
         TabIndex        =   17
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label labelEnterPassword 
         BackColor       =   &H00333333&
         Caption         =   "Password"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label buttonUnlock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00292929&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Unlock"
         ForeColor       =   &H00CCCCCC&
         Height          =   300
         Left            =   1920
         TabIndex        =   10
         Top             =   2280
         Width           =   2295
      End
   End
End
Attribute VB_Name = "formMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Windows API Functions
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function EnumDisplayMonitors Lib "user32" (ByVal hdc As Long, ByVal lprcClip As Long, ByVal lpfnEnum As Long, ByVal dwData As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

' Define Constants
Private Const REG_KEY As String = "Software\Microsoft\Windows\CurrentVersion\Run"
Private Const APP_NAME As String = "DesktopLocker"
Private Const CFG_FILE As String = "DesktopLocker.mrt"

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

Const SM_XVIRTUALSCREEN = 76
Const SM_YVIRTUALSCREEN = 77
Const SM_CXVIRTUALSCREEN = 78
Const SM_CYVIRTUALSCREEN = 79

' Global Variables
Dim currentPassword As String

Private Sub ChangeStatus(msg As String)
    labelStatus.Caption = Trim(msg)
End Sub

Private Sub FrameVisible(frameId As Integer)
    If frameId = 1 Then
        inputNewPassword.Text = ""
        inputNewPasswordConfirm.Text = ""
        
        ' Load Startup
        Dim startupPath As String
        startupPath = GetStartupPath()
        If startupPath <> "" Then
            checkStartup.Value = vbChecked
            buttonUninstall.Visible = True
        Else
            checkStartup.Value = vbUnchecked
            buttonUninstall.Visible = False
        End If
        
        frameSettings.Enabled = True
        frameSettings.Visible = True
        frameEnter.Enabled = False
        frameEnter.Visible = False
    Else
        inputEnterPassword.Text = ""
        
        frameSettings.Enabled = False
        frameSettings.Visible = False
        frameEnter.Enabled = True
        frameEnter.Visible = True
    End If
End Sub

Private Sub SetFullScreenAcrossMonitors()
    On Error Resume Next
    Dim x As Long, y As Long, width As Long, height As Long
    
    x = GetSystemMetrics(SM_XVIRTUALSCREEN)
    y = GetSystemMetrics(SM_YVIRTUALSCREEN)
    width = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    height = GetSystemMetrics(SM_CYVIRTUALSCREEN)
    
    Me.Move x, y, width * Screen.TwipsPerPixelX, height * Screen.TwipsPerPixelY
End Sub

Private Sub AddToStartup()
    On Error Resume Next
    Dim reg As Object
    Dim appPath As String
    appPath = App.path & "\" & App.EXEName & ".exe"
    Set reg = CreateObject("WScript.Shell")
    reg.RegWrite "HKEY_CURRENT_USER\" & REG_KEY & "\" & APP_NAME, appPath, "REG_SZ"
    Set reg = Nothing
End Sub

Private Sub RemoveFromStartup()
    On Error Resume Next
    Dim reg As Object
    Set reg = CreateObject("WScript.Shell")
    reg.RegDelete "HKEY_CURRENT_USER\" & REG_KEY & "\" & APP_NAME
    Set reg = Nothing
End Sub

Private Function GetStartupPath() As String
    On Error Resume Next
    Dim reg As Object
    Dim path As String
    Set reg = CreateObject("WScript.Shell")
    path = reg.RegRead("HKEY_CURRENT_USER\" & REG_KEY & "\" & APP_NAME)
    Set reg = Nothing
    GetStartupPath = path
End Function

Private Sub buttonCancel_Click()
    If currentPassword <> "" Then
        FrameVisible 2
        ChangeStatus "Locked"
        Exit Sub
    End If
    
    Call buttonUninstall_Click
End Sub

Private Sub buttonSave_Click()
    If inputNewPassword.Text = "" Or inputNewPasswordConfirm.Text = "" Or inputNewPassword.Text <> inputNewPasswordConfirm.Text Then
        ChangeStatus "Unlocked - Invalid Password!"
        inputNewPassword.Text = ""
        inputNewPasswordConfirm.Text = ""
        Exit Sub
    End If
    
    currentPassword = inputNewPassword.Text
    inputNewPassword.Text = ""
    inputNewPasswordConfirm.Text = ""
    
    Open App.path & "\" & CFG_FILE For Output As #1
        Print #1, "Desktop Locker Password: " & currentPassword & " Developer: Mehdi-RaTo (https://github.com/Mehdi-RaTo)"
    Close #1
    
    If checkStartup.Value <> vbChecked Then
        AddToStartup
    End If
    
    FrameVisible 2
    ChangeStatus "Locked"
End Sub

Private Sub buttonSettings_Click()
    If inputEnterPassword.Text <> currentPassword Then
        ChangeStatus "Locked - Wrong Password!"
        inputEnterPassword.Text = ""
        Exit Sub
    End If
    
    FrameVisible 1
    ChangeStatus "Unlocked"
End Sub

Private Sub buttonUninstall_Click()
    On Error Resume Next
    
    Kill App.path & "\" & CFG_FILE

    If checkStartup.Value = vbChecked Then
        RemoveFromStartup
    End If
    
    End
End Sub

Private Sub buttonUnlock_Click()
    If inputEnterPassword.Text <> currentPassword Then
        ChangeStatus "Locked - Wrong Password!"
        inputEnterPassword.Text = ""
        Exit Sub
    End If
    
    End
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Call SetFullScreenAcrossMonitors
    
    Dim fileContent As String
    Dim arrFileContent() As String
    Open App.path & "\" & CFG_FILE For Input As #1
        Line Input #1, fileContent
    Close #1
    arrFileContent = Split(fileContent, "Desktop Locker Password: ")
    If UBound(arrFileContent) > 0 Then
        currentPassword = Replace(arrFileContent(1), " Developer: Mehdi-RaTo (https://github.com/Mehdi-RaTo)", "")
        If currentPassword <> "" Then
            FrameVisible 2
            ChangeStatus "Locked"
            Exit Sub
        End If
    End If
    
    FrameVisible 1
    ChangeStatus "Unlocked"
End Sub

Private Sub timer_Timer()
    On Error Resume Next
    Call SetFullScreenAcrossMonitors
    Call SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Me.BorderStyle = 0
    Me.WindowState = 0
    Me.SetFocus
End Sub
