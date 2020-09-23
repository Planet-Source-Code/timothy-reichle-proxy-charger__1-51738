VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmProxy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proxy Logon"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogOut 
      Caption         =   "Log &Out"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   3135
   End
   Begin VB.Frame frmLogin 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CommandButton cmdLogon 
         Caption         =   "&Logon"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   3135
      End
      Begin VB.ComboBox cboHours 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Number Of Hours Extra"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Get Data"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3495
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblMoney 
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum enumState
    GetMoney = 0
    Logon = 1
End Enum

Const LoginIp = "127.0.0.1"
Const LoginPort = 10000
Const DataPort = 10001

Dim UserName As String * 20
Dim password As String * 50

Dim State As enumState

Dim TotalMoney As Currency
Dim HourlyCharge As Currency

Private Sub cmdConnect_Click()
    sckConnect.Close
    sckConnect.Connect LoginIp, DataPort
    
    
    Dim start As Date
    
    start = Now
    
    Do
        DoEvents
        
        If Second(Now - start) > 5 Then
            sckConnect.Close
            sckConnect.Connect LoginIp, DataPort
            start = Now
        End If
        
    Loop Until sckConnect.State = 7
    
    State = GetMoney
    
    UserName = txtUser
    password = txtPassword
    
    sckConnect.SendData Encode(UserName & password, "HTTP PROXY")
    DoEvents
End Sub

Private Sub cmdLogon_Click()
    sckConnect.Close
    sckConnect.Connect LoginIp, LoginPort
    
    
    Dim start As Date
    
    start = Now
    
    Do
        DoEvents
        
        If Second(Now - start) > 5 Then
            sckConnect.Close
            sckConnect.Connect LoginIp, LoginPort
            start = Now
        End If
        
    Loop Until sckConnect.State = 7
        
    State = GetMoney
        
    ConnectionStr = "I" & UserName & password & cboHours.Text
        
    sckConnect.SendData Encode(ConnectionStr, "HTTP PROXY")
End Sub

Private Sub cmdLogOut_Click()
    sckConnect.Close
    sckConnect.Connect LoginIp, LoginPort
    
    
    Dim start As Date
    
    start = Now
    
    Do
        DoEvents
        
        If Second(Now - start) > 5 Then
            sckConnect.Close
            sckConnect.Connect LoginIp, LoginPort
            start = Now
        End If
        
    Loop Until sckConnect.State = 7
        
    sckConnect.SendData Encode("O", "HTTP PROXY")
End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
    Dim Data As String
    
    sckConnect.GetData Data
    
    Dim PlaceOfComma1 As Integer, PlaceOfComma2 As Integer
     
    PlaceOfComma1 = InStr(1, Data, ",")
    PlaceOfComma2 = InStr(PlaceOfComma1 + 1, Data, ",")
    TotalMoney = Val(Left$(Data, PlaceOfComma1 - 1))
    HourlyCharge = Val(Mid$(Data, PlaceOfComma1, PlaceOfComma2 - PlaceOfComma1))
    CanConnectTill = Right$(Data, Len(Data) - PlaceOfComma2)
    State = Logon
    
    Dim NumOfHours As Integer
    
    lblMoney = "You have: " & Format(TotalMoney, "$0.00") & "." & vbCrLf & _
               "The hourly rate is: " & Format(HourlyCharge, "$0.00") & vbCrLf & _
               "You can Login Till: " & CanConnectTill
               
    If HourlyCharge = 0 Then
        NumOfHours = 24
    Else
GetNewCharge:
    If TotalMoney < HourlyCharge Then Exit Sub

        Select Case NumOfHours
            Case Is < 4
            Case Is < 8
                HourCharge = HourCharge * 0.9
            Case Is < 12
                HourCharge = HourCharge * 0.85
            Case Is < 16
                HourCharge = HourCharge * 0.8
            Case Is < 20
                HourCharge = HourCharge * 0.75
            Case Is < 24
                HourCharge = HourCharge * 0.7
        End Select
        
        MaxHours = Int(TotalAmount / HourCharge)
                   
        If NumOfHours > MaxHours Then
        
            If Int(MaxHours / 4) < Int(NumOfHours / 4) Then
                NumOfHours = MaxHours
                GoTo GetNewCharge
            End If
        End If
    End If
            
    For cnt = 0 To NumOfHours
        cboHours.AddItem cnt
    Next cnt
    
    If NumOfHours = 0 And CanConnectTill > Now Then
        Dim msg As String
    
     
        msg = "You do not have enough money" & _
            " to surf the net." & vbCrLf & "Please buy some more credits soon"
    
        Call MsgBox(msg, vbOKOnly, "Proxy Login")
    Else
        frmLogin.Visible = True
    End If
End Sub

Private Sub txtPassword_Change()
    frmLogin.Visible = False
End Sub

Private Sub txtUser_Change()
    frmLogin.Visible = False
End Sub
