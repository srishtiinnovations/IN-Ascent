' Author                     Created Date
' Manimaran                   24/11/2010
Public Class clsUOM
    Private oForm As SAPbouiCOM.Form
    Dim strSQL As String
    Public Const formtype As String = "SST_UOM"

    Public Sub LoadScreen()
        Dim j As Integer
        Dim omenus As SAPbouiCOM.MenuItem
        omenus = objAddOn.SBO_Application.Menus.Item("47616")
        For j = 0 To omenus.SubMenus.Count - 1
            strSQL = omenus.SubMenus.Item(j).String
            If strSQL.StartsWith("SST_UOM") = True Then
                objAddOn.SBO_Application.ActivateMenuItem(omenus.SubMenus.Item(j).UID.ToString)
                Exit For
            End If
        Next

    End Sub
End Class
