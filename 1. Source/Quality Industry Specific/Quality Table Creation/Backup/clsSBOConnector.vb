' Author                     Created Date
' Manimaran                   19/11/2010

Namespace SST
    Public Class SBOConnector
        Public Function GetApplication(ByVal ConnectionStr As String) As SAPbouiCOM.Application
            Dim objGUIAPI As SAPbouiCOM.SboGuiApi
            Dim objApp As SAPbouiCOM.Application

            Try
                objGUIAPI = New SAPbouiCOM.SboGuiApi
                objGUIAPI.Connect(ConnectionStr)
                objApp = objGUIAPI.GetApplication(-1)
                If Not objApp Is Nothing Then Return objApp
            Catch ex As Exception
                MsgBox(ex.Message)
                End
            End Try
            Return Nothing
        End Function

        Public Function GetCompany(ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
            Dim objCompany As New SAPbobsCOM.Company
            Dim strCookie As String
            Dim strConContext As String
            Try
                strCookie = objCompany.GetContextCookie()
                strConContext = SBOApplication.Company.GetConnectionContext(strCookie)
                objCompany.SetSboLoginContext(strConContext)
                objCompany.Connect()
                Return objCompany
            Catch ex As Exception
                MsgBox(ex.Message & vbLf & ex.StackTrace)
            End Try
            Return Nothing
        End Function
    End Class
End Namespace
