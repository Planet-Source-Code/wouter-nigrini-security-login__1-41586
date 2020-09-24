Attribute VB_Name = "Engine"
Option Explicit

Public Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetCursor Lib "User32" (ByVal hCursor As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

    Public Const REG_SZ = 1
    Public Const HKEY_LOCAL_MACHINE = &H80000002
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    Public Const HWND_TOPMOST = -1
    Public Const HWND_NOTOPMOST = -2

    Public bFirst As Boolean
    Public strUser As String, strPass As String
    Public strName As String
    Public iLogFail As Integer
    Public iEndTask As Integer
    
Public Sub CheckFirst()
On Error Resume Next
    
    'Get first run value
    bFirst = GetSetting(App.Path, "First", "Value")
    
    'Check if an previous instance of the program was found
    'If App.PrevInstance = True Then End
    
    'Check to see if it's the first run
    If bFirst = False Then
       'Display First run message
       MsgBox "This is the first time u run this utility logged on as this user." _
            & "Please enter a new Username and Password.", vbInformation + vbOKOnly _
             , "First Run"
             
       'Get the user's name
       strName = InputBox("Please enter your name:", "User's Name", "")
       
       'Get new username and password
       strUser = InputBox("Please enter a new Username:", "New Username", "")
       strPass = InputBox("Please enter a new Password:", "New Password", "")
       
       'Save new username and password
       SaveSetting App.Path, "Login", "Username", strUser
       SaveSetting App.Path, "Login", "Password", strPass
       SaveSetting App.Path, "User", "Name", strName
       
       'Set the bFirst value to False and save the value
       bFirst = True
       SaveSetting App.Path, "First", "Value", bFirst
       
       'Welcome user message
       MsgBox "Welcome " & strName & ", to unlock your computer, enter your Username and Password." & vbNewLine _
            & "NOTE: Username and Password is case sensitive!", vbInformation + vbOKOnly, "Welcome"
       
       'Display user's name in title
       strName = GetSetting(App.Path, "User", "Name")
       frmMain.Caption = strName & " Login"
    Else
       'If it's not the first run then get the saved username and password
       strUser = GetSetting(App.Path, "Login", "Username")
       strPass = GetSetting(App.Path, "Login", "Password")
       strName = GetSetting(App.Path, "User", "Name")
       
       'Display user's name in title
       frmMain.Caption = strName & " Login"
    End If
    
End Sub
