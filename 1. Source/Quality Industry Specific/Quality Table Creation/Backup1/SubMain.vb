''' Author                     Created Date
''' Manimaran                   17/11/2010
Module SubMain
    Public objAddOn As Menus
    Public Sub Main()
        objAddOn = New Menus
        objAddOn.Intialize()
        System.Windows.Forms.Application.Run()
    End Sub
End Module
