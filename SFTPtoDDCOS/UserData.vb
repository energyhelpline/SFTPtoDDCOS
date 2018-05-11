Imports System
Imports System.Security
Imports Microsoft.Exchange.WebServices.Data
Imports SFTPtoDDCOS

Public Interface IUserData

    Property Version As ExchangeVersion

    Property EmailAddress As String

    Property Password As SecureString

    Property AutodiscoverUrl As Uri

End Interface

'Public Class UserDataFromConsole
Public Class UserDataFromConfig
    Implements IUserData

    'Public Shared UserData As UserDataFromConsole
    Public Shared UserData As UserDataFromConfig

    Public Shared Function GetUserData() As IUserData
        If UserData Is Nothing Then
            ' GetUserDataFromConsole()
            GetUserDataFromConfig()
        End If

        Return UserData
        End Function

    'Private Shared Sub GetUserDataFromConsole()
    '    UserData = New UserDataFromConsole()
    '    Console.Write("Enter email address: ")
    '    UserData.EmailAddress = Console.ReadLine()
    '    UserData.Password = New SecureString()
    '    Console.Write("Enter password: ")
    '    While True
    '        Dim userInput As ConsoleKeyInfo = Console.ReadKey(True)
    '        If userInput.Key = ConsoleKey.Enter Then
    '            Exit While
    '        ElseIf userInput.Key = ConsoleKey.Escape Then
    '            Return
    '        ElseIf userInput.Key = ConsoleKey.Backspace Then
    '            If UserData.Password.Length <> 0 Then
    '                UserData.Password.RemoveAt(UserData.Password.Length - 1)
    '            End If
    '        Else
    '            UserData.Password.AppendChar(userInput.KeyChar)
    '            Console.Write("*")
    '        End If
    '    End While

    '    Console.WriteLine()
    '    UserData.Password.MakeReadOnly()
    'End Sub

    Private Shared Sub GetUserDataFromConfig()
        UserData = New UserDataFromConfig()
        UserData.EmailAddress = My.Settings.emailuseraddress
        UserData.Password = New SecureString()
        For Each ra As Char In My.Settings.email_password
            UserData.Password.AppendChar(ra)
        Next
        UserData.Password.MakeReadOnly()
        UserData.version = ExchangeVersion.Exchange2013_SP1
    End Sub

    Public Property Version As ExchangeVersion
    '    Get
    '        Return ExchangeVersion.Exchange2013_SP1
    '    End Get
    'End Property

    Public Property EmailAddress As String

    Public Property Password As SecureString

    Public Property AutodiscoverUrl As Uri

    Public Property IUserData_Version As ExchangeVersion Implements IUserData.Version
        Get
            Return ExchangeVersion.Exchange2013_SP1
        End Get
        Set(value As ExchangeVersion)
            'DirectCast(UserData, IUserData).Version = value
        End Set
    End Property

    Private Property IUserData_EmailAddress As String Implements IUserData.EmailAddress
        Get
            Return My.Settings.emailuseraddress
        End Get
        Set(value As String)
            DirectCast(UserData, IUserData).EmailAddress = value
        End Set
    End Property


    Private Property IUserData_Password As SecureString Implements IUserData.Password
        Get
            Dim pwd As New SecureString
            For Each ra As Char In My.Settings.email_password
                pwd.AppendChar(ra)
            Next
            pwd.MakeReadOnly()
            Return pwd
        End Get
        Set(value As SecureString)
            DirectCast(UserData, IUserData).Password = value
        End Set
    End Property

    Private Property IUserData_AutodiscoverUrl As Uri Implements IUserData.AutodiscoverUrl
        Get
            Return My.Settings.autodiscoveruri

        End Get
        Set(value As Uri)
            UserData.AutodiscoverUrl = value
        End Set
    End Property
End Class

