VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Menu"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuMenuChgUser 
         Caption         =   "Change Username"
      End
      Begin VB.Menu mnuMenuChgPass 
         Caption         =   "Change Password"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuMenuChgPass_Click()
    Dim strTempPass As String
        
    'Get password
    strPass = GetSetting(App.Path, "Login", "Password")
        
    'First give old Password to be able to change it
    strTempPass = InputBox("Enter old password:", "Change Password", "")
        
    If strTempPass = strPass Then
       'Change password
       strPass = InputBox("Please enter a new Password:", "New Password", "")
       
       'Save changed password
       SaveSetting App.Path, "Login", "Password", strPass
    Else
       'if the old password was entered incoreect then display error
       MsgBox "Invalid password!" & vbNewLine & "NOTE: Password is case sensitive." _
              , vbCritical + vbOKOnly, "Invalid Old Password"
    End If
End Sub

Private Sub mnuMenuChgUser_Click()
    Dim strTempUser As String
        
    'Get username
    strUser = GetSetting(App.Path, "Login", "Username")
        
    'First give old Username to be able to change it
    strTempUser = InputBox("Enter old username:", "Change Username", "")
        
    If strTempUser = strUser Then
       'Change username
       strUser = InputBox("Please enter a new Username:", "New Username", "")
       
       'Save changed username
       SaveSetting App.Path, "Login", "Username", strUser
    Else
       'if the old username was entered incoreect then display error
       MsgBox "Invalid username!" & vbNewLine & "NOTE: Username is case sensitive." _
              , vbCritical + vbOKOnly, "Invalid Old Username"
    End If
End Sub
