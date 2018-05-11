Imports Microsoft.Exchange.WebServices.Data
Imports System.Net
Imports System.IO
Imports Renci.SshNet

Public Class Worker
    Private mMustStop As Boolean
    Private mMain As System.Threading.Thread
    Shared userdata As IUserData = UserDataFromConfig.GetUserData()
    Shared es As ExchangeService = Service.ConnectToService(userdata)


    Public Sub DoWork()
        'Create a thread to do some work
        mMain = System.Threading.Thread.CurrentThread
        mMain.Name = "OPSThread"
        'Loop to perform the main jobs you wish to run.
        While Not mMustStop
            MainLoop()
            'Period to wait 
            System.Threading.Thread.Sleep(My.Settings.timerinterval * 60000)
        End While

    End Sub

    Public Sub StopWork()
        mMustStop = True
        If Not mMain Is Nothing Then
            If Not mMain.Join(100) Then
                mMain.Abort()
            End If
        End If
        TearDown()
    End Sub

    Private Sub MainLoop()
        Try
            'Put your subs in here that you want the service to process
            job1()

        Catch ex As Exception
            My.Application.Log.WriteException(ex, TraceEventType.Error, ex.StackTrace)
            'sendmail("Exception has occured on " & My.Computer.Name, , , , ex.ToString)

        End Try

    End Sub

    Private Sub TearDown()

    End Sub

    Private Sub Job1()

        Dim view As ItemView = New ItemView(Integer.MaxValue)
        'Dim itempropertyset As PropertySet = New PropertySet(EmailMessageSchema.Attachments)
        Dim searchy As SearchFilter = New SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, False)
        'view.PropertySet = itempropertyset

        'view.SearchFilter = searchy
        Dim results As FindItemsResults(Of Item)
        results = es.FindItems(WellKnownFolderName.Inbox, searchy, view)

        If (results.TotalCount > 0) Then
            For Each mailitem As EmailMessage In results.Items
                mailitem.Load(New PropertySet(EmailMessageSchema.Attachments))
                Dim x As Integer = mailitem.Attachments.Count
                For Each attachy As Attachment In mailitem.Attachments
                    If TypeOf attachy Is FileAttachment Then
                        'is the file the one we want? check it matches the file extension
                        Dim dotposition = attachy.Name.LastIndexOf(".")
                        If attachy.Name.Substring(dotposition + 1) <> "txt" Then
                            Continue For

                        End If
                        'read attachment and load into memory
                        Dim fileattachy As FileAttachment = attachy
                        Dim filestream As IO.Stream = New IO.MemoryStream
                        fileattachy.Load(filestream)
                        filestream.Position = 0
                        'sftp the file
                        If Sftpfile(filestream, fileattachy.Name) Then
                            'sftp success - move email
                            Dim destfolder As FolderId = FindFolderIdByDisplayName(es, My.Settings.foldertomoveto, WellKnownFolderName.Inbox)
                            mailitem.IsRead = True
                            mailitem.Update(ConflictResolutionMode.AlwaysOverwrite)
                            Dim msg As Item = Item.Bind(es, mailitem.Id)
                            msg.Move(destfolder)
                            My.Application.Log.WriteEntry("Successfully uploaded " + attachy.Name)
                        Else
                            'something went wrong with the SFTP upload
                            Dim em As EmailMessage = New EmailMessage(es) With {
                                .Subject = "*** SFTPDDCOS upload Failure ***",
                                .Body = "Failed to upload " + attachy.Name
                                }
                            em.ToRecipients.Add(My.Settings.emailalerts)
                            em.From.Address = My.Settings.smtpmailfrom
                            em.SendAndSaveCopy()
                        End If
                        filestream.Dispose()

                    End If
                Next
            Next

        End If


    End Sub



    Public Function Sftpfile(ByVal inputfile As IO.Stream, ByVal fname As String) As Boolean


        Dim connectioninfo As New ConnectionInfo(My.Settings.sftphost, My.Settings.sftpusername, New PasswordAuthenticationMethod(My.Settings.sftpusername, My.Settings.sftppassword))
        Dim client As New SftpClient(connectioninfo)

        Try
            client.Connect()
            client.UploadFile(inputfile, "/" + fname)
            Sftpfile = True
        Catch ex As Exception

            Sftpfile = False
        Finally
            If client IsNot Nothing Then
                client.Dispose()
            End If
        End Try


    End Function



    Public Shared Function FindFolderIdByDisplayName(ByVal service As ExchangeService, ByVal DisplayName As String, ByVal SearchFolder As WellKnownFolderName) As FolderId
        Dim rootFolder As Folder = Folder.Bind(service, SearchFolder)
        For Each folder As Folder In rootFolder.FindFolders(New FolderView(100))
            If folder.DisplayName = DisplayName Then
                Return folder.Id
            End If
        Next

        Return Nothing
    End Function


End Class