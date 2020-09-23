VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password Changer"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConfirmNewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   50
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   3480
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdChangePassword 
      Caption         =   "&Submit"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "&Confirm New Password"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "&New Password"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "&Old Password:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "&Username:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PasswordIP = "127.0.0.1"
Const PasswordPort = "10002"

Dim FilledUser As Boolean, FilledPassword As Boolean
Dim FilledNewPassword As Boolean, FilledConfirmPassword As Boolean

Private Sub cmdChangePassword_Click()
    
    If txtNewPassword.Text <> txtConfirmNewPassword.Text Then _
        MsgBox "New Passwords Don't Match": Exit Sub
        
    If txtNewPassword.Text <> txtConfirmNewPassword.Text Then _
        MsgBox "New Passwords Don't Match": Exit Sub

    With sckConnect
        .Close
        .Connect PasswordIP, PasswordPort
    
        Dim Start As Date
        
        Start = Now
        
        Do
            DoEvents
            If Second(Now - Start) >= 5 And .State <> 7 Then
                .Close
                .Connect PasswordIP, PasswordPort
            End If
        Loop Until .State = 7
    
    End With
    
End Sub


Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
    Dim DataRecieved As String
    
    sckConnect.GetData DataRecieved
    
    If DataRecieved = "Sorry, server is full!" Then
        MsgBox "The server is full, try again later"
    ElseIf DataRecieved = "Connection Accepted. Have a lot of fun." Then
        MsgBox "Connected"
        
        Dim User As String * 20
        Dim Password As String * 50
        Dim NewPassword As String * 50
        Dim StringToSend As String * 120
        
        User = txtUser.Text
        Password = txtOldPassword.Text
        NewPassword = txtNewPassword.Text
        
        StringToSend = User & Password & NewPassword
        
        sckConnect.SendData Encode$(StringToSend, "PASSWORD")
    ElseIf DataRecieved = "SUCCESFUL" Then
        MsgBox "Password Changed"
    Else
        MsgBox "The server is not running at the moment. Sorry :-("
    End If
    
End Sub

Private Sub CheckFilled()
    cmdChangePassword.Enabled = True
    
    If Not FilledConfirmPassword Then
        cmdChangePassword.Enabled = False
    ElseIf Not FilledNewPassword Then
        cmdChangePassword.Enabled = False
    ElseIf Not FilledPassword Then
        'cmdChangePassword.Enabled = False
    ElseIf Not FilledUser Then
        cmdChangePassword.Enabled = False
    End If
End Sub
 
Private Sub txtConfirmNewPassword_Change()
    If Len(txtConfirmNewPassword.Text) = 0 Then
        cmdChangePassword.Enabled = False
        FilledConfirmPassword = False
    Else
        FilledConfirmPassword = True
        CheckFilled
    End If
End Sub

Private Sub txtNewPassword_Change()
    If Len(txtNewPassword.Text) = 0 Then
        cmdChangePassword.Enabled = False
        FilledNewPassword = False
    Else
        FilledNewPassword = True
        CheckFilled
    End If
End Sub

Private Sub txtOldPassword_Change()
    If Len(txtOldPassword.Text) = 0 Then
        cmdChangePassword.Enabled = False
        FilledPassword = False
    Else
        FilledPassword = True
        CheckFilled
    End If
End Sub

Private Sub txtUser_Change()
    If Len(txtUser.Text) = 0 Then
        cmdChangePassword.Enabled = False
        FilledUser = False
    Else
        FilledUser = True
        CheckFilled
    End If
End Sub
