VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Johan Nigrini Login"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Change"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1800
      Top             =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1999 - 2002, Wouter Nigrini"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   5.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   667
      TabIndex        =   6
      Top             =   1200
      Width           =   1920
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
    Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
    Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long


    Dim X As Long, Y As Long
    Dim i  As Long
    
Private Sub cmdChange_Click()
    Me.PopupMenu frmMenu.mnuMenu, 1
End Sub

Private Sub cmdLogin_Click()
    'Check to see if the usernaqme and password was entered correct
    If txtUser.Text = strUser And txtPassword.Text = strPass Then
       'Disable the mouse jumping
       Timer1.Enabled = False
       
       'If the Fail count is not = to 0 then display message
       If iLogFail <> 0 Then
          If iEndTask <> 0 Then
             MsgBox "User(s) tried to access to your computer " & iLogFail & " time(s) but they failed!" & vbNewLine _
                  & iEndTask & " user(s) tried to End Task this utility to gain entry to your computer!", vbInformation + vbOKOnly, "Failed logons"
          Else
             MsgBox "User(s) tried to access to your computer " & iLogFail & " time(s) but they failed!", vbInformation + vbOKOnly, "Failed logons"
          End If
       End If
       
       'Exit the program
       End
    Else
       'If username and password was incorrect then display an error message
       MsgBox "Invalid Username and password, please try again." & vbNewLine _
            & "NOTE: Username and password is case sensitive!", vbCritical + vbOKOnly, "Invalid Login"
        
       iLogFail = iLogFail + 1
    End If

End Sub

Private Sub Form_Activate()
    'Set the focus to the Username box
    txtUser.SetFocus
End Sub

Private Sub Form_Load()
    Dim nBufferKey As Long, nBufferSubKey As Long
    
    'Check to see if it's the first time if this program is run
    Engine.CheckFirst
    
    'Set the Fail and EndTask count to zero
    iLogFail = 0
    iEndTask = 0
    
    'Set this program on top of other windows
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    'Save the applications path to the "run" key in the regidtry to enable the program to run at startup
    RegOpenKey HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\", nBufferKey
    RegOpenKey nBufferKey, "run", nBufferSubKey
    RegSetValueEx nBufferSubKey, "Login", 0, REG_SZ, App.Path & "\" & App.EXEName & ".exe", Len(App.Path & "\" & App.EXEName & ".scr")

    'Start the mouse jumping
    Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Disable program from exiting
    Cancel = 1
    
    'Prevent user from end tasking program!
    'If user tries to end task app then run program again!
    Call Shell(App.Path & "\" & App.EXEName & ".exe", vbNormalFocus)
    
    'Show message to warn user
    MsgBox "Your system is still locked for unauthorized entry!" & vbNewLine _
         & "Please enter the correct username and password to continue.", vbCritical + vbOKOnly, "Unauthorized attempt to exit"
    
    'Count the EndTask counter
    iEndTask = iEndTask + 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Disable the user from exiting the app
    Cancel = 1
    
    'Check the username and password
    cmdLogin_Click
    
    'Count the Fail counter
    iLogFail = iLogFail + 1
End Sub

Private Sub Timer1_Timer()
    'Set a random number to the mouse's X pos
    X = Int(Rnd * 800)
    'Set a random number to the mouse's y pos
    Y = Int(Rnd * 600)
    
    'Set the cursor's XY pos accordingly to the random X and Y pos
    SetCursorPos X, Y
    
    DoEvents
End Sub

Private Sub txtPassword_KeyPress(KeyCode As Integer)
    'If the <Enter> button was pressed move on to check the username and password
    If KeyCode = vbKeyReturn Then cmdLogin_Click
End Sub

Private Sub txtUser_KeyPress(KeyCode As Integer)
    'If the <enter> button was pressed move on to the password field
    If KeyCode = vbKeyReturn Then txtPassword.SetFocus

End Sub
