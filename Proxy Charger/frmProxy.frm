VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmProxy 
   Caption         =   "HTTP Proxy"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sckDatabase 
      Left            =   3600
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10001
   End
   Begin VB.TextBox txtUsers 
      Height          =   3735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   1815
   End
   Begin VB.Timer tmrCheck 
      Interval        =   1000
      Left            =   2280
      Top             =   1320
   End
   Begin VB.Timer tmrAccept 
      Interval        =   100
      Left            =   1800
      Top             =   1320
   End
   Begin MSWinsockLib.Winsock sckToProxy 
      Index           =   0
      Left            =   2280
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckAccept 
      Index           =   0
      Left            =   1440
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckProxy 
      Left            =   600
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8080
   End
   Begin MSWinsockLib.Winsock sckLogin 
      Left            =   3840
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10000
   End
   Begin MSWinsockLib.Winsock sckAcceptLogin 
      Index           =   0
      Left            =   3120
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1000
   End
   Begin MSWinsockLib.Winsock sckAcceptDatabase 
      Index           =   0
      Left            =   3000
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Loged In Users"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FreeNumber As Integer
Dim LoginFreeNumber As Integer
Dim DataFreeNumber As Integer

Dim Ip() As String

Const ProxyIP = "127.0.0.1" '"Where is the actual Proxy
Const ProxyPort = 8079 '"What port is the actual Proxy

Private Declare Sub SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long)

Private Sub Form_Load()
    '*** Set Everything Up ***
    
    mdlDatabase.ConnectToDatabase
    
    sckProxy.Listen
    sckLogin.Listen
    sckDatabase.Listen
    
End Sub

Private Sub Form_Resize()

    '*** Maintain the sizes of the control in relation to form size
    Const MinWidth = 6150
    Const MinHeight = 4665

    If Me.WindowState = vbMinimized Then Exit Sub

    If Me.Width < MinWidth Then
        Me.Width = MinWidth
        MouseX = (Me.Left + MinWidth) / Screen.TwipsPerPixelX
        MouseY = (Me.Top + Me.Height) / Screen.TwipsPerPixelY
    End If
    If MouseX > 0 Then SetCursorPos MouseX, MouseY
    If Me.Height < MinHeight Then
        Me.Height = MinHeight
        MouseX = (Me.Left + Me.Width) / Screen.TwipsPerPixelX
        MouseY = (Me.Top + MinHeight) / Screen.TwipsPerPixelY
    End If
    If MouseX > 0 Then SetCursorPos MouseX, MouseY


    With txtUsers
        .Width = Me.Width / 3
        .Height = Me.Height - 900
    End With
End Sub

Private Sub sckAccept_Close(Index As Integer) 'Discconects the last socket
    
    If Index > 0 Then
        Unload sckAccept(Index)
        Unload sckToProxy(Index)
    End If
    
    If FreeNumber = Index + 1 Then FreeNumber = Index

End Sub


Private Sub sckAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim Temp As String
    
        Dim start As Date
    
        start = Now
        sckToProxy(Index).Connect ProxyIP, ProxyPort
    
        Do
            DoEvents
        
            If Second(Now - start) > 5 Then
                sckToProxy(Index).Close
                sckToProxy(Index).Connect ProxyIP, ProxyPort
                start = Now
            End If
        
        Loop Until sckToProxy(Index).State = 7

    
    sckAccept(Index).GetData Temp     'Get Requests from User
      
    Dim HTTP_Is_At  As Integer, lpos As Integer
    
    lpos = InStr(1, Temp, " ", vbTextCompare)
      
    HTTP_Is_At = InStr(1, Temp, "HTTP")
      
    ToSend = Left$(Temp, lpos) & Right$(Temp, Len(Temp) - HTTP_Is_At + 3)
    
    On Error GoTo NewLog
    
    Open App.Path & "\UserLogs\" & UserOfIpAdress(sckAccept(Index).RemoteHostIP) & ".Log" For Input As #1
    
    total = ""
    
    Do Until EOF(1)
        Line Input #1, lineStr
            
        total = total & lineStr & vbCrLf
    Loop
        
NewLog:
        
    Close #1
    
    Open App.Path & "\UserLogs\" & UserOfIpAdress(sckAccept(Index).RemoteHostIP) & ".Log" For Output As #1
    total = total & "---------------------------------" & vbCrLf & Now & vbCrLf & vbCrLf & ToSend
    
    Print #1, total
    
    Close #1
    sckToProxy(Index).SendData ToSend    'Send Requests to proxy
End Sub

Private Sub sckAcceptDatabase_Close(Index As Integer)
    '************************************************
    'Login Close
    '************************************************
    
    If Index > 0 Then Unload sckAcceptDatabase(Index)
    
    If DataFreeNumber = Index + 1 Then DataFreeNumber = Index
End Sub


Private Sub sckAcceptDatabase_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim UserName As String * 20
    Dim Password As String * 50
    Dim Hours As Integer
    
    Dim RecievedData As String
    
    sckAcceptDatabase(Index).GetData RecievedData
    
    RecievedData = mdlCode.Decode(RecievedData, "HTTP PROXY")
        
    UserName = Left$(RecievedData, 20)
    Password = Mid$(RecievedData, 21, 50)

    Dim Temp As LoginReturn

    Temp = mdlDatabase.LogIn(UserName, Password, 0)
    
    Dim ReturnStr As String
    
    ReturnStr = Temp.Money & "," & Temp.HourCharge & "," & Temp.ConnectTill
    
    sckAcceptDatabase(Index).SendData ReturnStr
End Sub

Private Sub sckDatabase_ConnectionRequest(ByVal requestID As Long)

    '***Someone wants to connect to the database***

    If DataFreeNumber > 0 Then Load sckAcceptDatabase(DataFreeNumber)
    sckAcceptDatabase(DataFreeNumber).Close
    sckAcceptDatabase(DataFreeNumber).Accept requestID
    DataFreeNumber = DataFreeNumber + 1
End Sub

Private Sub sckToProxy_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    'Send data from proxy to user
    
    sckToProxy(Index).GetData Temp     'Send Html etc from proxy
    sckAccept(Index).SendData Temp    'Send Html etc to user
End Sub

Private Sub sckAcceptLogin_Close(Index As Integer)
    '***Login Close***
    
    If LoginFreeNumber = Index + 1 Then LoginFreeNumber = Index
End Sub


Private Sub sckAcceptLogin_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    '***Login Here***
    
    Dim UserName As String * 20
    Dim Password As String * 50
    Dim Hours As Integer
    
    Dim DataReturned As String
    
    sckAcceptLogin(Index).GetData DataReturned
    
    DataReturned = mdlCode.Decode(DataReturned, "HTTP PROXY")
    
    LogType = Left$(DataReturned, 1)
    
    If LogType = "I" Then
    
        UserName = Mid$(DataReturned, 2, 20)
        Password = Mid$(DataReturned, 22, 50)
        Hours = Val(Right$(DataReturned, Len(DataReturned) - 70))

        Dim LR As LoginReturn

        LR = mdlIP.LogIn(sckAcceptLogin(Index).RemoteHostIP, UserName, Password, Hours)
    
        Dim ReturnStr As String
        
        ReturnStr = LR.Money & "," & LR.HourCharge & "," & LR.ConnectTill
    
        sckAcceptLogin(Index).SendData ReturnStr
        sckAcceptLogin(Index).Close
    Else
        Call mdlIP.LogOff("", "", sckAcceptLogin(Index).RemoteHostIP)
    
        sckAcceptLogin(Index).Close
    End If
    
End Sub

Private Sub sckLogin_ConnectionRequest(ByVal requestID As Long)
    
    '***Someone wants to login***

    If LoginFreeNumber > 0 Then Load sckAcceptLogin(LoginFreeNumber)
    sckAcceptLogin(LoginFreeNumber).Accept requestID
    LoginFreeNumber = LoginFreeNumber + 1
End Sub

Private Sub sckProxy_ConnectionRequest(ByVal requestID As Long)
    '***Someone wants a webpage***
    Dim start As Date
        
    If FreeNumber > 0 Then
    
        Load sckAccept(FreeNumber)
        
        sckAccept(FreeNumber).Accept requestID
        
        If mdlIP.IpAdressIndex(sckAccept(FreeNumber).RemoteHostIP) = 0 Then
            sckAccept(FreeNumber).SendData "Unauthorised IP. Please Logon"
            Exit Sub
        End If
                
        Load sckToProxy(FreeNumber)
        
        FreeNumber = FreeNumber + 1
        
        
    

    
    Else
        sckAccept(0).Close
        sckAccept(0).Accept requestID
        sckToProxy(0).Close
    
        If mdlIP.IpAdressIndex(sckAccept(0).RemoteHostIP) = 0 Then
            sckAccept(0).SendData "<html>Unauthorised IP. Please Logon</html>"
            DoEvents
            sckAccept(0).Close
            Exit Sub
        End If
        FreeNumber = 1
        
    End If
End Sub


Private Sub tmrAccept_Timer()

    '***Close Sockets Not in Use**
    
    'Login Close
    If LoginFreeNumber <= 1 Then GoTo redo2
redo:
    On Error GoTo errorhandle
    If sckAcceptLogin(LoginFreeNumber - 1).State = 0 Then
        Unload sckAcceptLogin(LoginFreeNumber - 1)
        LoginFreeNumber = LoginFreeNumber - 1
    End If
    
    'Proxy Close
redo2:
    On Error GoTo errorhandle2
    If FreeNumber <= 1 Then GoTo redo3
    If sckToProxy(FreeNumber - 1).State = 0 Then
        Unload sckToProxy(FreeNumber - 1)
        Unload sckAccept(FreeNumber - 1)
        FreeNumber = FreeNumber - 1
    End If
    
    'Data Close
redo3:
    On Error GoTo errorhandle3
    If DataFreeNumber <= 1 Then Exit Sub
    If sckToProxy(DataFreeNumber - 1).State = 0 Then
        Unload sckAcceptDatabase(DataFreeNumber - 1)
        DataFreeNumber = DataFreeNumber - 1
    End If

    Exit Sub
errorhandle:
        
    'The control is already unloaded
    If Err = 340 Then LoginFreeNumber = LoginFreeNumber - 1: GoTo redo
     GoTo redo2
     
errorhandle2:

    'The control is already unloaded
    If Err = 340 Then FreeNumber = FreeNumber - 1:    GoTo redo2
    GoTo redo2
errorhandle3:
    
    'The control is already unloaded
    If Err = 340 Then DataFreeNumber = DataFreeNumber - 1:    GoTo redo3
End Sub

Private Sub tmrCheck_Timer()
    Call mdlIP.LogOffUsersWhoRunOutTime
End Sub
