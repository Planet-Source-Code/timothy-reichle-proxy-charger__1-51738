Attribute VB_Name = "mdlDatabase"
Public Enum enumUserType
    LogedIn = 1
    Incorect_Password = 2
    PiggyBack = 3
    NoMoney = 4
    NoNonPiggyBack = 5
End Enum

Public Type LoginReturn
    User As String
    ConnectTill As Date
    UserType As enumUserType
    HourCharge As Currency
    Money As Currency
End Type

Dim dbData As Database

Dim rsUser As Recordset
Dim rsTotalAmount As Recordset
Dim rsTransaction As Recordset

Public Sub ConnectToDatabase()

    '*** Connect to the database and set up the recordsets***

    Set dbData = OpenDatabase("C:\Documents and Settings\Timothy\My Documents\Programing\Visual Basic\Proxy\users.mdb")
    Set rsTotalAmount = dbData.OpenRecordset("qryAmount")
    Set rsTransaction = dbData.OpenRecordset("tblTransactions")
    Set rsUser = dbData.OpenRecordset("qryUser")
End Sub

Private Sub NewTransaction(User As String, Amount As Single)

    '*** A user has spent some money , this must be recorded here***

    With rsTransaction
        .Edit
        .AddNew
        .Fields("Name") = User
        .Fields("Amount") = Amount
        .Update
    End With
End Sub

Public Function LogIn(User As String, ByVal Password As String, ByVal NumOfHours As Byte) As LoginReturn
    
    '*** Check that the user can afford the number of hours and update accordingly
    
    Dim Target As String
    
    Target = "Name = '" & User & "'"
    
    rsUser.FindFirst Target
    
    Dim Time As Date
    Dim MaxHours As Integer
    Dim HourCharge As Currency, TotalAmount As Currency
    
    If Not rsUser.EOF Then
        
        If IsNull(rsUser.Fields("Password")) Then
            rsUser.Fields("Password") = Encode$(Decode$(rsUser.Fields("Password"), "PASSWORD"), "PASSWORD")
            If RTrim$(Password) = Decode$(rsUser.Fields("Password"), "PASSWORD") Then
                LogIn.UserType = Incorect_Password
                Exit Function
            End If
        Else
            If rsUser.Fields("PiggyBack") Then
                LogIn.UserType = PiggyBack
                Exit Function
            End If
        End If
        
        rsTotalAmount.FindFirst Target
        
        TotalAmount = rsTotalAmount.Fields("Amount")
        
GetNewCharge:

        HourCharge = rsUser.Fields("FeePerHour")
        LogIn.HourCharge = HourCharge
        
        If HourCharge > 0 Then
        
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
            
                If MaxHours = 0 Then
                    LogIn.UserType = NoMoney
                    Exit Function
                End If
                If Int(MaxHours / 4) < Int(NumOfHours / 4) Then
                    NumOfHours = MaxHours
                    GoTo GetNewCharge
                End If
                DoEvents
            End If
        End If
        

        If rsUser.Fields("AcessTill") < Now Or IsNull(rsUser.Fields("AcessTill")) Then
            With rsUser
                .Edit
                .Fields("AcessTill") = DateAdd("h", NumOfHours, Now)
                LogIn.ConnectTill = .Fields("AcessTill")
                .Update
            End With
            If HourCharge > 0 Then Call NewTransaction(User, -NumOfHours * HourCharge)
        Else
            With rsUser
                .Edit
                .Fields("AcessTill") = DateAdd("h", NumOfHours, .Fields("AcessTill"))
                LogIn.ConnectTill = .Fields("AcessTill")
                .Update
            End With
            If HourCharge > 0 Then Call NewTransaction(User, -NumOfHours * HourCharge)
        End If
        
        LogIn.UserType = LogedIn
        
        Set rsTotalAmount = dbData.OpenRecordset("qryAmount")
        
        rsTotalAmount.FindFirst Target
        
        LogIn.Money = rsTotalAmount.Fields("Amount")
        LogIn.HourCharge = HourCharge
        
    End If
End Function



