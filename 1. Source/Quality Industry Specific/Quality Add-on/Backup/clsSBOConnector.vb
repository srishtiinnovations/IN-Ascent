Imports System.Data.OleDb
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
                If objCompany.Connected = True Then
                    Return objCompany
                Else
                    Dim strqry As String
                    strqry = "select * from [@sst_login]"
                    Dim ds As New DataSet
                    Dim Splt As String()
                    Dim LogDet As String = ""
                    Dim XMLReader As Xml.XmlReader
                    XMLReader = New Xml.XmlTextReader(Application.StartupPath & "\CmpyDetails.xml")

                    While XMLReader.Read
                        Select Case XMLReader.NodeType
                            Case Xml.XmlNodeType.Element
                                Debug.WriteLine(XMLReader.Name)
                                If XMLReader.AttributeCount > 0 Then
                                    While XMLReader.MoveToNextAttribute
                                        Debug.WriteLine(XMLReader.Name & "->" & XMLReader.Value)
                                        If LogDet = "" Then
                                            LogDet = XMLReader.Value
                                        Else
                                            LogDet = LogDet + "," + XMLReader.Value

                                        End If
                                    End While
                                End If
                        End Select
                    End While
                    XMLReader.Close()

                    Splt = Strings.Split(LogDet, ",")
                    ds = GetDataSet(strqry, Splt(3), Splt(2), Splt(0), Splt(1))


                    If ds.Tables(0).Rows.Count > 0 Then
                        objCompany.UserName = ds.Tables(0).Rows(0).Item("U_UN").ToString
                        Dim alg As Crypto.Algorithm = 3
                        Crypto.EncryptionAlgorithm = alg
                        Crypto.Key = ""
                        Crypto.Encoding = Crypto.EncodingType.HEX
                        Crypto.Content = ds.Tables(0).Rows(0).Item("U_PWD").ToString
                        If Crypto.DecryptString = True Then
                            objCompany.Password = Crypto.Content
                            'objCompany.Password = "bone"
                        End If
                    End If

                    objCompany.DbUserName = Splt(0)
                    objCompany.DbPassword = Splt(1)
                    objCompany.CompanyDB = Splt(2)
                    objCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
                    If Splt(4) = "MSSQL2008" Then
                        objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                    Else
                        objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
                    End If
                    objCompany.Server = Splt(3)
                    If objCompany.Connect <> 0 Then
                        MessageBox.Show(objCompany.GetLastErrorDescription)
                    Else
                        Return objCompany
                    End If
                End If

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End Function

        Public Function GetDataSet(ByVal sQuery As String, ByVal Server As String, ByVal Company As String, ByVal DataBase As String, ByVal PWD As String) As DataSet

            Dim oCONNECTION_STRING As String
            Dim oConnection As New OleDb.OleDbConnection()
            Dim oQUERY_STRING As String = ""
            Dim oDataAdapter As OleDb.OleDbDataAdapter
            Dim oDataSet As New DataSet()
            Try
                oCONNECTION_STRING = "Provider=SQLOLEDB;Server=" & Server & ";Database=" & Company & ";User ID=" & DataBase & ";Password=" & PWD & ""
                oConnection = New OleDbConnection(oCONNECTION_STRING)
                oConnection.Open()

                oDataAdapter = New OleDbDataAdapter(sQuery, oConnection)
                oDataAdapter.Fill(oDataSet)

            Catch ObjEx As Exception
                Throw New Exception(ObjEx.Message)
            Finally
                oConnection.Close()
            End Try

            Return oDataSet
        End Function

        'for vortex

        'Public Function GetApplication(ByVal ConnectionStr As String) As SAPbouiCOM.Application
        '    Dim objGUIAPI As SAPbouiCOM.SboGuiApi
        '    Dim objApp As SAPbouiCOM.Application

        '    Try
        '        objGUIAPI = New SAPbouiCOM.SboGuiApi
        '        '26-4-11
        '        '5645523035496D706C656D656E746174696F6E3A4A323038333630323933336B5A28F64AC1DDA36B2DF993A2407A7C7FCF896A
        '        'for our purpose
        '        'objGUIAPI.AddonIdentifier = "5645523035496D706C656D656E746174696F6E3A4A323038333630323933336B5A28F64AC1DDA36B2DF993A2407A7C7FCF896A"
        '        'for vortex
        '        objGUIAPI.AddonIdentifier = "5645523035496D706C656D656E746174696F6E3A4B30333037393938373932DC679DDD04AFA644AE5E3C3D40C974A478A6411E"
        '        objGUIAPI.Connect(ConnectionStr)
        '        objApp = objGUIAPI.GetApplication()
        '        If Not objApp Is Nothing Then Return objApp
        '    Catch ex As Exception
        '        MsgBox(ex.Message)

        '        End
        '    End Try
        '    Return Nothing
        'End Function
        'Public Function GetCompany(ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
        '    Dim objCompany As New SAPbobsCOM.Company
        '    Dim strCookie As String
        '    Dim strConContext As String
        '    Dim LogDet As String = ""
        '    Dim Splt As String()
        '    Try
        '        Dim XMLReader As Xml.XmlReader
        '        XMLReader = New Xml.XmlTextReader(Application.StartupPath & "\CmpyDetails.xml")

        '        While XMLReader.Read
        '            Select Case XMLReader.NodeType
        '                Case Xml.XmlNodeType.Element
        '                    Debug.WriteLine(XMLReader.Name)
        '                    If XMLReader.AttributeCount > 0 Then
        '                        While XMLReader.MoveToNextAttribute
        '                            Debug.WriteLine(XMLReader.Name & "->" & XMLReader.Value)
        '                            If LogDet = "" Then
        '                                LogDet = XMLReader.Value
        '                            Else
        '                                LogDet = LogDet + "," + XMLReader.Value

        '                            End If
        '                        End While
        '                    End If
        '            End Select
        '        End While
        '        XMLReader.Close()

        '        Splt = Strings.Split(LogDet, ",")
        '        objCompany.UserName = Splt(0)
        '        Dim alg As Crypto.Algorithm = 3
        '        Crypto.EncryptionAlgorithm = alg
        '        Crypto.Key = ""
        '        Crypto.Encoding = Crypto.EncodingType.HEX
        '        Crypto.Content = Splt(1)
        '        If Crypto.DecryptString = True Then
        '            objCompany.Password = Crypto.Content
        '        End If
        '        objCompany.DbUserName = Splt(2)
        '        Crypto.Content = Splt(3)
        '        If Crypto.DecryptString = True Then
        '            objCompany.DbPassword = Crypto.Content
        '        End If
        '        objCompany.CompanyDB = Splt(4)
        '        objCompany.language = SAPbobsCOM.BoSuppLangs.ln_English
        '        If Splt(6) = "MSSQL2008" Then
        '            objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
        '        Else
        '            objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
        '        End If
        '        objCompany.Server = Splt(5)
        '        If objCompany.Connect <> 0 Then
        '            MessageBox.Show(objCompany.GetLastErrorDescription + "...Connection failed")
        '        Else
        '            Return objCompany
        '        End If
        '        Return Nothing
        '    Catch ex As Exception
        '        SBOApplication.SetStatusBarMessage("Give Company Details.....", SAPbouiCOM.BoMessageTime.bmt_Short, True)
        '    End Try
        '    Return objCompany
        'End Function

    End Class
End Namespace