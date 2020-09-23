Attribute VB_Name = "mdlIP"
'Piggyback users will only be able to connect when a non-piggyback user is connected

Private Type User
    Ip As String
    PiggyBack As Boolean
    NamePasswordIndex As Integer
    TimeIndex As Integer
End Type

Private Type NamePasswordIndex
    UserName As String
    Password As String
    Index As Integer
End Type

Private Type TimeIndex
    LogOfAt As Date
    Index As Integer
End Type

Private User() As User 'Keep track of all the users
Private UserNameP() As NamePasswordIndex 'Indexed User name
Private UserLogOffTime() As TimeIndex 'Indexed LogOfTime

Private NumOfUsers As Integer 'How many users

Dim NumOfPiggyBack As Integer, NumOfNonPiggyBack As Integer 'How man users are piggybacking/not piggybacking



'Bookmarks are code that needs to be edited for user/password index

Public Function IpAdressIndex(IpAdress As String) As Integer
    '***Find the index of a particular IP adress***
    If NumOfUsers = 0 Then IpAdressIndex = 0: Exit Function
    
    
    '*** Sort using database index method ***
    
    
    Dim MidRange As Integer 'Where is the current index
    Dim MinIndex As Integer, MaxIndex As Integer  'For Range of Sort
    
    MinIndex = 1: MaxIndex = NumOfUsers 'Set the range of search
    
   Do
        MidRange = Int((MaxIndex - MinIndex) / 2)  'Where Is halfway
        If MidRange = 0 Then MidRange = MinIndex
        DoEvents
        If IpAdress > User(MidRange).Ip Then
            'In Top Half
            MaxIndex = MidRange - 1 'Set the new max index
        ElseIf IpAdress < User(MidRange).Ip Then
            'In Bottom Half
            MinIndex = MidRange + 1 'Set the new max index
        End If
        
    Loop Until IpAdress = User(MidRange).Ip Or MaxIndex = MinIndex
    
    If Not (IpAdress = User(MidRange).Ip) Then IpAdressIndex = 0 'No matches
    
    IpAdressIndex = MidRange 'Match found, Return the index

End Function

Private Function UserIndex(UserName As String, Password As String) As Integer
    '***Find the index of a particular user***
    
    '*** Sort using database index method ***
    
    If NumOfUsers = 0 Then UserIndex = 0: Exit Function
    
    
    Dim MidRange As Integer 'Where is the current index
    Dim MinIndex As Integer, MaxIndex As Integer  'For Range of Sort
    
    MinIndex = 1: MaxIndex = NumOfUsers 'Set the range of search
    
    If NumOfUsers = 1 Then
    
    End If
    
   Do
        MidRange = Int((MaxIndex - MinIndex) / 2)  'Where Is halfway
        If MaxIndex = MinIndex Then MidRange = MinIndex
        DoEvents
        If UserName > UserNameP(MidRange).UserName Then
            'In Top Half
            MaxIndex = MidRange - 1 'Set the new max index
        ElseIf UserName < UserNameP(MidRange).UserName Then
            'In Bottom Half
            MinIndex = MidRange + 1 'Set the new max index
        End If
        
    Loop Until (UserName = UserNameP(MidRange).UserName = UserName And _
    UserName = UserNameP(MidRange).UserName = Password) Or MaxIndex = MinIndex
    
    If Not (UserName = UserNameP(MidRange).UserName = UserName And _
        UserName = UserNameP(MidRange).UserName = Password) Then UserIndex = 0: Exit Function 'No matches
    
    UserIndex = MidRange 'Match found, Return the index

End Function


Public Function UserOfIpAdress(ByVal IpAdress As String) As String
    Dim IpIndex As Integer
    
    IpIndex = IpAdressIndex(IpAdress)
    
    UserOfIpAdress = RTrim$(UserNameP(User(IpIndex).NamePasswordIndex).UserName)
    
End Function

Public Function LogIn(IpAdress As String, UserName As String, Password As String, Hours As Integer) As LoginReturn
    '***Log in users***
    
    Dim Temp As LoginReturn
    
    Temp = mdlDatabase.LogIn(UserName, Password, Hours)
        
    Dim PiggyBack As Boolean
        
    PiggyBack = False
        
    Select Case Temp.UserType
        Case 1 'LogedIn
            LogIn.UserType = 1
        Case 2 'Incorect_Password
            LogIn.UserType = 2
            Exit Function
        Case 3 'PiggyBack
            If AreNonPiggyBack Then
                LogIn.UserType = 3
                PiggyBack = True
            Else
                LogIn.UserType = 5
                Exit Function
            End If
        Case 4 'NoMoney
            LogIn.UserType = 4
            Exit Function
    End Select
    
    Dim IpIndex As Integer
    
    IpIndex = IpAdressIndex(IpAdress)
    
    If IpIndex = 0 Then 'Not found
        ReDim Preserve User(NumOfUsers + 1) As User 'Main User
        ReDim Preserve UserNameP(NumOfUsers + 1) As NamePasswordIndex 'Indexed User Name
        ReDim Preserve UserLogOffTime(NumOfUsers + 1) As TimeIndex 'Index Time
        
        
        NumOfUsers = NumOfUsers + 1
        IpIndex = NumOfUsers
        UserNameIndex = NumOfUsers
        TimeIndex = NumOfUsers
    Else
        UserNameIndex = User(IpIndex).NamePasswordIndex
        TimeIndex = User(IpIndex).TimeIndex
    End If
    

    With User(IpIndex)
        .Ip = IpAdress
        .PiggyBack = PiggyBack
        .NamePasswordIndex = UserNameIndex
        .TimeIndex = TimeIndex
    End With
    
    With UserNameP(UserNameIndex)
        .UserName = UserName
        .Password = Password
        .Index = IpIndex
    End With
    
    With UserLogOffTime(TimeIndex)
        .LogOfAt = Temp.ConnectTill
        .Index = IpIndex
    End With
    
    With LogIn
        .ConnectTill = Temp.ConnectTill
        .HourCharge = Temp.HourCharge
        .HourCharge = Temp.Money
    End With
    
    If PiggyBack Then
        NumOfPiggyBack = NumOfPiggyBack + 1
    Else
        NumOfNonPiggyBack = NumOfNonPiggyBack + 1
    End If
    
    LogedInUsers
    
    Index
End Function

Public Function LogOff(ByVal UserName As String, ByVal Password As String, Optional ByVal IpAdress As String) As Boolean

    '***Log off a user***

    Dim UI As Integer
    Dim NameIndex As Integer, TimeIndex As Integer 'Index of array to remove
    Dim NP As Integer, TI As Integer 'Index of array to update
    
    'Get User Index
    
    If IpAdress = "" Then
        UI = UserIndex(UserName, Password)
        If UI = 0 Then LogOff = False: Exit Function
        NameIndex = UI
        UI = User(NameIndex).NamePasswordIndex
    Else
        UI = IpAdressIndex(IpAdress)
        If UI = 0 Then LogOff = False: Exit Function
        NameIndex = User(UI).NamePasswordIndex
    End If
    
    'The actual log off
    
    If UI > 0 Then
    
        If User(UI).PiggyBack Then
            NumOfPiggyBack = NumOfPiggyBack - 1
        Else
            NumOfNonPiggyBack = NumOfNonPiggyBack - 1
        End If
    
        
        TimeIndex = User(UI).TimeIndex
       
        For cnt = UI To NumOfUsers - 1 'Remove user from main array
            User(cnt) = User(cnt + 1)
            With User(cnt + 1)  'Update pointers '\
                NP = .NamePasswordIndex          '|
                TI = .TimeIndex                  '|
            End With                             '|
            UserNameP(NP).Index = cnt + 1        '|
            UserLogOffTime(TI).Index = cnt + 1   '/
        Next cnt
        ReDim Preserve User(NumOfUsers - 1) As User
        
        For cnt = NameIndex To NumOfUsers - 1 'Remove user from name index
            UserNameP(NameIndex) = UserNameP(NameIndex + 1)
            User(UserNameP(NameIndex + 1).Index).NamePasswordIndex = NameIndex + 1 'Update pointer
        Next cnt
        ReDim Preserve UserNameP(NameIndex - 1) As NamePasswordIndex
        
        For cnt = TimeIndex To NumOfUsers - 1 'Remove user from log off time index
            UserLogOffTime(TimeIndex) = UserLogOffTime(TimeIndex + 1)
            User(UserLogOffTime(TimeIndex + 1).Index).TimeIndex = TimeIndex + 1 'Update pointer
        Next cnt
        ReDim Preserve UserLogOffTime(TimeIndex - 1) As TimeIndex
        
        
        NumOfUsers = NumOfUsers - 1
        LogedInUsers
        LogOff = True
    Else
        LogOff = False 'Log-off Failed
    End If
End Function

Public Sub LogOffUsersWhoRunOutTime()

    '*** If the time is up they need to be logged off***

    For cnt = 1 To NumOfUsers
        If UserLogOffTime(cnt).LogOfAt < Now Then
            Call LogOff("", "", User(UserLogOffTime(cnt).Index).Ip)   'Log off the person who has run out of time
        Else
            Exit Sub 'No more users to log off
        End If
    Next cnt
End Sub

Public Function FindPiggyBackerIndex(Optional StartAt As Integer) As Integer
    
    '*** Find a piggybacker with their index above StartAt ****
    
    If StartAt < 0 Then StartAt = 1 'If nothing entered or negative then use default start
    
    For cnt = StartAt To NumOfUsers
        If User(cnt).PiggyBack Then FindPiggyBackerIndex = cnt: Exit Function
    Next cnt

End Function

Public Sub LogOffPiggyBack()

    '***Logoff the users that are piggybacking***


    Dim PiggyBacker As Integer

    If NumOfNonPiggyBack > 0 Then Exit Sub
    PiggyBacker = FindPiggyBackerIndex(0)
    Do Until PiggyBacker = 0
        DoEvents
        
        Call LogOff("", "", User(cnt).Ip)
        
        PiggyBacker = FindPiggyBackerIndex(PiggyBacker)
        
    Loop
End Sub


Public Sub LogedInUsers()
    '***List who is on in a text box***
    
    Dim Temp As String
    
    Temp = ""
    
    For cnt = 1 To NumOfUsers
        Temp = Temp & RTrim$(UserNameP(cnt).UserName) & ",IP: " & User(UserNameP(cnt).Index).Ip & vbCrLf
    Next cnt
    frmProxy.txtUsers = Temp
    
End Sub

Public Sub Index()
    '*** Fixed up the 3 indexs ***

    Dim tempLR As User, tempN As NamePasswordIndex, tempT As TimeIndex
    
    
    For OuterCnt = 0 To NumOfUsers - 1
        For cnt = OuterCnt + 1 To NumOfUsers - 1
            If User(cnt).Ip > User(cnt + 1).Ip Then
                tempLR = User(cnt)
                User(cnt) = User(cnt + 1)
                User(cnt) = tempLR
                
                UserNameP(cnt).Index = cnt + 1 'Update the pointers
                UserNameP(cnt + 1).Index = cnt
                UserLogOffTime(cnt).Index = cnt + 1
                UserLogOffTime(cnt + 1).Index = cnt
            End If
            
            If UserNameP(cnt).UserName > UserNameP(cnt + 1).UserName Then
                tempN = UserNameP(cnt)
                UserNameP(cnt) = UserNameP(cnt + 1)
                UserNameP(cnt) = tempN
                
                User(UserNameP(cnt).Index).NamePasswordIndex = cnt + 1 'Update the pointers
                User(UserNameP(cnt + 1).Index).NamePasswordIndex = cnt - 1
            End If
            
            If UserLogOffTime(cnt).LogOfAt > UserLogOffTime(cnt + 1).LogOfAt Then
                tempT = UserLogOffTime(cnt)
                UserLogOffTime(cnt) = UserLogOffTime(cnt + 1)
                UserLogOffTime(cnt) = tempT
                
                User(UserLogOffTime(cnt).Index).TimeIndex = cnt + 1 'Update the pointers
                User(UserLogOffTime(cnt + 1).Index).TimeIndex = cnt - 1
            End If
            
        Next cnt
    Next OuterCnt
End Sub
