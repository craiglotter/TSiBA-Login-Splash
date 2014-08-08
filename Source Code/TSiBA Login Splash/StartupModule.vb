Module StartupModule
    Sub main(ByVal args() As String)
        Try
            If args.Length = 2 Then
                Dim splash As New Splash_Screen(args(0), args(1))
                Application.Run(splash)
            Else
                MsgBox("Usage: <filename.exe> <heading> <message.rtf> where ""heading"" is the Title string to be displayed on the main form and ""message.rtf"" is the path of the required input message file for the splash screen to display.", MsgBoxStyle.Information, "Incorrect Startup Detected")
            End If
        Catch ex As Exception
            Application.Exit()
        End Try
    End Sub
End Module
