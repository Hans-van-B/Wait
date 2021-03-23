Module Log
    Sub xtrace_init()
        My.Computer.FileSystem.WriteAllText(LogFile, "xtrace_init" & vbNewLine, False)
        xtrace(AppName & " V" & AppVer)
        xtrace_i("AppRoot = " & AppRoot)
    End Sub

    '---- xtrace
    Sub xtrace_root(Msg As String, TV As Integer)
        If TV <= CTrace Then
            Console.Write(Msg)
        End If

        My.Computer.FileSystem.WriteAllText(LogFile, Msg, True)
    End Sub

    Sub xtrace(Msg As String)
        xtrace_root(Msg & vbNewLine, 2)
    End Sub

    Sub xtrace(Msg As String, TV As Integer)
        xtrace_root(Msg & vbNewLine, TV)
    End Sub

    '---- xtrace_i
    Sub xtrace_i(Msg As String)
        xtrace(" * " & Msg)
    End Sub

    Sub xtrace_i(Msg As String, TV As Integer)
        xtrace(" * " & Msg, TV)
    End Sub

    '---- xtrace_err
    Sub xtrace_err(Msg As String)
        xtrace("ERROR: " & Msg, 1)
        ErrorCount = ErrorCount + 1
    End Sub
End Module
