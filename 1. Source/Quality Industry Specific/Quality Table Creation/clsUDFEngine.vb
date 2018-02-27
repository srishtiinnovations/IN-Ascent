Namespace SST

    Public Class UDFEngine
        Private objCompany As SAPbobsCOM.Company

        Public Sub New(ByVal Company As SAPbobsCOM.Company)
            objCompany = Company
        End Sub


#Region "Table Functions"
        Public Function CreateTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
            Dim objUserTableMD As SAPbobsCOM.UserTablesMD
            Dim ret As Integer
            Dim str As String = ""
            objUserTableMD = Nothing
            GC.Collect()
            objUserTableMD = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            Try
                If (Not objUserTableMD.GetByKey(TableName)) Then
                    objUserTableMD.TableName = TableName
                    objUserTableMD.TableDescription = TableDescription
                    objUserTableMD.TableType = TableType
                    If objUserTableMD.Add() = 0 Then
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
                        objUserTableMD = Nothing
                        objAddOn.SBO_Application.SetStatusBarMessage("Table" + TableName + "Created successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                        Return True
                    Else
                        ' Throw New Exception(objAddOn.oCompany.GetLastErrorDescription)
                        objAddOn.oCompany.GetLastError(ret, str)
                        MsgBox(str)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
                        objUserTableMD = Nothing
                        Return False
                    End If
                Else
                    objAddOn.SBO_Application.SetStatusBarMessage("Table" + TableName + "Created successfully", SAPbouiCOM.BoMessageTime.bmt_Short, False)
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            Finally
                objUserTableMD = Nothing
                GC.Collect()
            End Try
        End Function
#End Region

#Region "Loading Default form"
        Public Sub LoadDefaultForm(ByVal sFormUID As String)
            Dim i As Integer
            ' Link to the Default Forms menu
            Dim sboMenu As SAPbouiCOM.MenuItem = objAddOn.SBO_Application.Menus.Item("47616")
            Try
                ' Iterate through the submenus to find the correct UDO
                If sboMenu.SubMenus.Count > 0 Then
                    For i = 0 To sboMenu.SubMenus.Count - 1
                        If sboMenu.SubMenus.Item(i).String.Contains(sFormUID) Then
                            sboMenu.SubMenus.Item(i).Activate()
                        End If
                    Next
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
#End Region
        Public Sub AddCol(ByVal TableName As String, ByVal ColName As String, ByVal ColDesc As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, Optional ByVal EditSize As Integer = 10, Optional ByVal SubType As SAPbobsCOM.BoFldSubTypes = 0)
            Dim objUserFields As SAPbobsCOM.UserFieldsMD
            Dim intError As Integer

            objUserFields = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            objUserFields.TableName = TableName
            objUserFields.Name = ColName
            objUserFields.Type = FieldType
            objUserFields.SubType = SubType
            objUserFields.Description = ColDesc
            objUserFields.EditSize = EditSize
            intError = objUserFields.Add()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFields)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            If intError <> 0 Then
                Throw New Exception(objAddOn.oCompany.GetLastErrorDescription)
            End If
        End Sub

        Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "")
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue)
        End Sub

        Public Sub AddAlphaMemoField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        End Sub

        Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue)
        End Sub

        Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
            Dim intLoop As Integer
            Dim ret As Integer
            Dim str As String
            Dim strValue, strDesc As Array
            Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD

            strValue = ValidValues.Split(Convert.ToChar(","))
            'MsgBox(strValue(0))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Valid value Code and Descriptions mismatching")
            End If
            objUserFieldMD = Nothing
            GC.Collect()
            objUserFieldMD = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
            Try
                If (Not isColumnExist(TableName, ColumnName)) Then
                    objUserFieldMD.TableName = TableName
                    objUserFieldMD.Name = ColumnName
                    objUserFieldMD.Description = ColDescription
                    objUserFieldMD.Type = FieldType
                    If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                        objUserFieldMD.Size = Size
                    Else
                        objUserFieldMD.EditSize = Size
                    End If
                    objUserFieldMD.SubType = SubType
                    objUserFieldMD.DefaultValue = SetValidValue
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                    If objUserFieldMD.Add() <> 0 Then
                        objAddOn.oCompany.GetLastError(ret, str)
                        'MsgBox(Str)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
                        objUserFieldMD = Nothing
                        ' Throw New Exception(objAddOn.oCompany.GetLastErrorDescription)
                    Else
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
                        objUserFieldMD = Nothing
                    End If
                End If
            Catch ex As Exception

                Throw ex
            Finally
                objUserFieldMD = Nothing
                GC.Collect()
            End Try
        End Sub

        Public Sub AddNumericField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "")
        End Sub

        Public Sub AddNumericField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal DefultValue As String)
            Try
                addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, DefultValue)
            Catch ex As Exception
                MsgBox(ex.Message & TableName & ColumnName)
            End Try
        End Sub

        Public Sub AddFloatField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "")
        End Sub

        Public Sub AddDateField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes)
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "")
        End Sub

        Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
            Dim objRecordSet As SAPbobsCOM.Recordset
            objRecordSet = objAddOn.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE TableID = '" & TableName & "' AND AliasID = '" & ColumnName & "'")
                If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
                GC.Collect()
            End Try
        End Function
    End Class

End Namespace