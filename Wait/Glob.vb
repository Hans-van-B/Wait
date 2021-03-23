Module Glob
    Public AppName As String = "Wait"
    Public AppVer As String = "0.02"

    Public AppRoot As String = IO.Path.GetDirectoryName(Diagnostics.Process.GetCurrentProcess().MainModule.FileName)
    Public CD As String = My.Computer.FileSystem.CurrentDirectory
    Public Temp As String = Environment.GetEnvironmentVariable("TEMP")

    Public LogFile As String = Temp & "\" & AppName & ".log"
    Public Verbose As Boolean = False
    Public CTrace As Integer = 1                       ' Console Trace Level
    Public ErrorCount As Integer = 0
    Public ExitProgram As Boolean = False

    ' Command-line Arg
    Public SwWait As Integer = 6
    Public SwPause As Boolean = False
    Public Bullet As String = ""

    ' Environment
    Public EnvNoWait As Boolean = False
End Module
