Module CommandLineArg
    Sub Read_Command_Line_Arg()
        xtrace_i("Read_Command_Line_Arg")

        Dim SwName As String = ""
        Dim SwString As String = ""
        Dim P1 As Integer
        Dim Name As String
        Dim ValS As String
        Dim MemArg As String = ""

        For Each argument As String In My.Application.CommandLineArgs
            xtrace_i("Argument = " & argument)

            '---- Double-dash arguments
            If Left(argument, 2) = "--" Then
                SwName = Mid(argument, 3)
                xtrace_i("DDA      : " & SwName)

                '---- Double-dash Assignment
                P1 = InStr(SwName, "=")
                If P1 > 0 Then
                    Name = Left(SwName, P1 - 1)
                    ValS = Mid(SwName, P1 + 1)
                    xtrace_i("DDA Name = " & Name)

                    If Name = "w" Then
                        SwWait = Val(ValS)
                        xtrace_i("Wait = " & ValS)
                    End If

                    Continue For
                End If

                '---- Double-dash No Assignment
                xtrace_i("DDN Name = " & Name)

                If SwName = "help" Then
                    ShowHelp()
                    ExitProgram = True
                End If

                If SwName = "xhelp" Then
                    ShowXHelp()
                    ExitProgram = True
                End If

                Continue For
            End If

            '---- Single-dash arguments
            If Left(argument, 1) = "-" Then
                ' Switch String = remaining switches
                SwString = Mid(argument, 2)

                ' for each switch in the string
                While Len(SwString) > 0
                    SwName = Left(SwString, 1)
                    SwString = Mid(SwString, 2)
                    xtrace_i("SDA:" & SwName & "," & SwString)

                    If SwName = "b" Then
                        Bullet = " * "
                        xtrace_i("Set bullet")
                    End If

                    If SwName = "h" Then
                        ShowHelp()
                        ExitProgram = True
                    End If

                    If SwName = "p" Then
                        SwPause = True
                        SwWait = 0
                        xtrace_i("Set Pause = True")
                    End If

                    If SwName = "v" Then
                        xtrace_i("Set verbose = True")
                        Console.WriteLine("Log file = " & LogFile)
                        Verbose = True
                    End If
                End While

                Continue For
            End If

            '---- Slash arguments
            If Left(argument, 1) = "/" Then
                SwName = Mid(argument, 2)
                xtrace_i("SA Name  = " & SwName)

                If SwName = "?" Or SwName = "h" Then
                    ShowHelp()
                    ExitProgram = True
                    Continue For
                End If

                MemArg = SwName
                xtrace_i("MemSA    = " & MemArg)
                Continue For
            End If

            If (MemArg <> "") Then
                ' Compatibility with timeout
                If (MemArg = "t") Then
                    Try
                        SwWait = Val(argument)
                        xtrace_i("Wait     = " & SwWait.ToString)
                    Catch ex As Exception
                        xtrace_err("Invalid argument")
                    End Try
                    Continue For
                End If

                MemArg = ""
                Continue For
            End If


            '---- Else (No dashes)
            P1 = InStr(argument, "=")
            If P1 > 0 Then
                Name = Left(argument, P1 - 1)
                ValS = Mid(argument, P1 + 1)
                xtrace_i("NDA:" & argument)

                If Name = "??" Then

                End If

                Continue For
            End If

            '---- Else (No dashes, No Assign)

            ' for compatibility with the older wait command
            Try
                SwWait = Val(argument)
                xtrace_i("Wait = " & SwWait.ToString)
            Catch ex As Exception
                xtrace_err("Invalid argument")
            End Try

            MemArg = ""

        Next

    End Sub
End Module
