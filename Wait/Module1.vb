Module Module1

    Sub Main()
        xtrace_init()
        Read_Command_Line_Arg()
        If ExitProgram Then
            xtrace_i("Exit program 0")
            Exit Sub
        End If


        If UCase(Environment.GetEnvironmentVariable("TA_BATCH")) = "TRUE" Then
            EnvNoWait = True
            wait(6)
        End If


        If Not EnvNoWait Then
            If SwWait > 0 Then
                wait(SwWait)
            End If

            If SwPause Then
                Console.Write(Bullet & "Type any key to continiue . . . ")
                Console.ReadKey()
                Console.Write(vbNewLine)
            End If
        End If

    End Sub

End Module
