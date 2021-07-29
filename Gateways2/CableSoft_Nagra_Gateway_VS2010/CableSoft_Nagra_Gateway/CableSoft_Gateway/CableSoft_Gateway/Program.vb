Module Program

    Public Sub Main()
        EurekaLogSystem.ExceptionHandler.Activate()
        EurekaLogSystem.ExceptionHandler.CurrentEurekaLogOptions.OutputPath = _
            Environment.CurrentDirectory
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
		Application.Run(New rfrmMain)
	End Sub

End Module