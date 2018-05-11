Public Class Service1
    Private worker As New Worker()
    Protected Overrides Sub OnStart(ByVal args() As String)

        System.IO.Directory.SetCurrentDirectory(GetMyDir)
        Dim wt As System.Threading.Thread
        Dim ts As System.Threading.ThreadStart
        ts = AddressOf worker.DoWork
        wt = New System.Threading.Thread(ts)
        My.Application.Log.WriteEntry("Service Started", TraceEventType.Information, 60500)

        wt.Start()


    End Sub

    Protected Overrides Sub OnStop()
        worker.StopWork()
        My.Application.Log.WriteEntry("Service Stopped", TraceEventType.Information, 60501)
    End Sub


    Function GetMyDir() As String
        Dim fi As System.IO.FileInfo
        Dim di As System.IO.DirectoryInfo
        Dim pc As System.Diagnostics.Process
        Try
            pc = System.Diagnostics.Process.GetCurrentProcess
            fi = New System.IO.FileInfo(pc.MainModule.FileName)
            di = fi.Directory
            GetMyDir = di.FullName
        Finally
            fi = Nothing
            di = Nothing
            pc = Nothing
        End Try
    End Function
End Class
