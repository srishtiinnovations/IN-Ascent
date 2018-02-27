Module modStartup
    Public objAddOn As clsAddOn

    Public Sub Main()
        objAddOn = New clsAddOn
        objAddOn.Intialize()
        System.Windows.Forms.Application.Run()
    End Sub
End Module
