Imports System
Imports System.Net
Imports Microsoft.Exchange.WebServices.Data



Module Service

    Sub New()
        CertificateCallback.Initialize()
    End Sub

    Private Function RedirectionUrlValidationCallback(ByVal redirectionUrl As String) As Boolean
            Dim result As Boolean = False
            Dim redirectionUri As Uri = New Uri(redirectionUrl)
            If redirectionUri.Scheme = "https" Then
                result = True
            End If

            Return result
        End Function

    'Function ConnectToService(ByVal userData As IUserData) As ExchangeService
    '    Return ConnectToService(userData, Nothing)
    'End Function

    Function ConnectToService(ByVal userData As IUserData) As ExchangeService
        Dim service As ExchangeService = New ExchangeService(userData.Version)

        'If listener IsNot Nothing Then
        '        service.TraceListener = listener
        '        service.TraceFlags = TraceFlags.All
        '        service.TraceEnabled = True
        '    End If

        service.Credentials = New NetworkCredential(My.Settings.email_user, userData.Password, My.Settings.email_domain)
        'If userData.AutodiscoverUrl Is Nothing Then
        '    'Console.Write(String.Format("Using Autodiscover to find EWS URL for {0}. Please wait... ", userData.EmailAddress))
        '    service.AutodiscoverUrl(userData.EmailAddress, AddressOf RedirectionUrlValidationCallback)
        '    userData.AutodiscoverUrl = service.Url
        '    'Console.WriteLine("Autodiscover Complete")
        'Else
        '    service.Url = userData.AutodiscoverUrl
        'End If

        service.AutodiscoverUrl(My.Settings.emailuseraddress)
        userData.AutodiscoverUrl = service.Url


        Return service
    End Function

    Function ConnectToServiceWithImpersonation(ByVal userData As IUserData, ByVal impersonatedUserSMTPAddress As String) As ExchangeService
            Return ConnectToServiceWithImpersonation(userData, impersonatedUserSMTPAddress, Nothing)
        End Function

        Function ConnectToServiceWithImpersonation(ByVal userData As IUserData, ByVal impersonatedUserSMTPAddress As String, ByVal listener As ITraceListener) As ExchangeService
            Dim service As ExchangeService = New ExchangeService(userData.Version)
            If listener IsNot Nothing Then
                service.TraceListener = listener
                service.TraceFlags = TraceFlags.All
                service.TraceEnabled = True
            End If

            service.Credentials = New NetworkCredential(userData.EmailAddress, userData.Password)
            Dim impersonatedUserId As ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.SmtpAddress, impersonatedUserSMTPAddress)
            service.ImpersonatedUserId = impersonatedUserId
            If userData.AutodiscoverUrl Is Nothing Then
                service.AutodiscoverUrl(userData.EmailAddress, AddressOf RedirectionUrlValidationCallback)
                userData.AutodiscoverUrl = service.Url
            Else
                service.Url = userData.AutodiscoverUrl
            End If

            Return service
        End Function
    End Module

