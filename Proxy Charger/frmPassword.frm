VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server to change passwords"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Quit"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock sckAccept 
      Index           =   0
      Left            =   6600
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckListen 
      Left            =   6120
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10002
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' a Multiconnection Winsock TCP Server by over
'
' If you have got any questions about this example contact me:
' overkillpage@gmx.net
'
' You can change the basic server settings like port and maximum connections
' in the modServer.bas module.
'
' Greetings fly out to AbsoluteB
'

Option Explicit

Dim dbData As Database
Dim rsUser As Recordset

' starting our server
Private Sub Form_Load()
    Set dbData = OpenDatabase("C:\Documents and Settings\Timothy\My Documents\Programing\Visual Basic\Proxy\users.mdb")
    Set rsUser = dbData.OpenRecordset("qryUser")
    
    
    ' Set sckListen to listen mode. It will accept all incoming connection requests.
    sckListen.LocalPort = ServerPort
    sckListen.Listen

    ' create some accept sockets
    If Not InitAcceptSockets Then
        MsgBox "ERROR Can't create accept sockets!", vbCritical, "Error"
        End
    End If
    
    ' debug some server information

End Sub



' a new client connected
Private Sub sckListen_ConnectionRequest(ByVal requestID As Long)

    Dim aFreeSocket As Integer
    
    ' Request the number of an unused socket
    aFreeSocket = GetFreeSocket
    
    If aFreeSocket = 0 Then
        
        ' Tell the new client that the server is full and close the connection
        sckAccept(0).Accept requestID
        DoEvents
        sckAccept(0).SendData "Sorry, server is full!"
        DoEvents
        sckAccept(0).Close
        
    Else
        
        ' accept the connection on a free socket. set status of this socket to true(used)
        bSocketStatus(aFreeSocket) = True
        sckAccept(aFreeSocket).Accept requestID
        DoEvents
        ' Send a welcome message to the new client
        sckAccept(aFreeSocket).SendData "Connection Accepted. Have a lot of fun."
        ' Refresh the combobox -> add our new client
        
        
    End If
    
End Sub



' One of the connected clients sent some data ...
' Add login function and additional commands here ...
Private Sub sckAccept_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    Dim sData As String
    Dim UserName As String * 20
    Dim Password As String * 50
    Dim NewPassword As String * 50
    
    Dim DataReturned As String
    
    sckAccept(Index).GetData sData
    
    DataReturned = Decode$(sData, "PASSWORD")
    ' output the incoming data to the debug textbox
    
    
    UserName = Mid$(DataReturned, 1, 20)
    Password = Mid$(DataReturned, 21, 50)
    NewPassword = Mid$(DataReturned, 71, 50)

    Dim Target As String
    
    
    Target = "Name='" & RTrim$(UserName) & "'"
    

    rsUser.FindFirst (Target)
    
    If RTrim$(Password) = Decode$(rsUser.Fields("Password"), "PASSWORD") Or IsNull(rsUser.Fields("Password")) Then
        With rsUser
            .Edit
            .Fields("Password") = Encode$(RTrim$(NewPassword), "PASSWORD")
            .Update
        End With
        sckAccept(Index).SendData "SUCCESFUL"
    End If
    
    
End Sub



' a client disconnected
Private Sub sckAccept_Close(Index As Integer)
    
    ' Free the used socket.
    bSocketStatus(Index) = False
    sckAccept(Index).Close
    
    
End Sub



