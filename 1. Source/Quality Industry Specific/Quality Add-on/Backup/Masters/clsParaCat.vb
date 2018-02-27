' Author                     Created Date
' Manimaran                   19/11/2010
Public Class clsParaCat
    Private oForm As SAPbouiCOM.Form
    Public Const formtype As String = "SST_PARACAT"
    Dim strSQL As String

    Public Sub LoadScreen()
        Dim j As Integer
        Dim omenus As SAPbouiCOM.MenuItem
        omenus = objAddOn.SBO_Application.Menus.Item("47616")
        For j = 0 To omenus.SubMenus.Count - 1
            strSQL = omenus.SubMenus.Item(j).String
            If strSQL.StartsWith("SST_PARACAT") = True Then
                objAddOn.SBO_Application.ActivateMenuItem(omenus.SubMenus.Item(j).UID.ToString)
                Exit For
            End If
        Next
       
    End Sub
   
End Class
