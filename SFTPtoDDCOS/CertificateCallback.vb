Imports System.Net



Module CertificateCallback

    Sub New()
        ServicePointManager.ServerCertificateValidationCallback = AddressOf CertificateValidationCallBack
    End Sub

    Sub Initialize()
        End Sub

        Private Function CertificateValidationCallBack(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate, ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
            If sslPolicyErrors = System.Net.Security.SslPolicyErrors.None Then
                Return True
            End If

            If (sslPolicyErrors And System.Net.Security.SslPolicyErrors.RemoteCertificateChainErrors) <> 0 Then
                If chain IsNot Nothing AndAlso chain.ChainStatus IsNot Nothing Then
                    For Each status As System.Security.Cryptography.X509Certificates.X509ChainStatus In chain.ChainStatus
                        If (certificate.Subject = certificate.Issuer) AndAlso (status.Status = System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.UntrustedRoot) Then
                            Continue For
                        Else
                            If status.Status <> System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.NoError Then
                                Return False
                            End If
                        End If
                    Next
                End If

                Return True
            Else
                Return False
            End If
        End Function
    End Module

