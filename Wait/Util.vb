Module Util
    '---- Wait ----------------------------------------------------------------
    Public Sub wait(ByVal interval As Integer)
        xtrace_root(Bullet & "Wait " & interval.ToString & " Sec...       (Space to continiue)", 1)
        interval = interval * 1000

        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval
            If Console.KeyAvailable Then
                If Strings.LCase(Console.ReadKey(True).KeyChar) = " " Then Exit Do
            End If
        Loop
        sw.Stop()

        xtrace_root(vbNewLine, 1)
    End Sub
End Module
