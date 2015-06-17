Namespace _FormConfiguration
    Public Class Configuration
        Dim DataSet As DataSet
        Dim ConfigFileFullName As String

        Public Sub New()
            DataSet = New DataSet
        End Sub

        Public Sub Reload()
            Load(ConfigFileFullName)
        End Sub

        Public Function Load(ByVal ConfigFileFullName As String) As Boolean
            If Not My.Computer.FileSystem.FileExists(ConfigFileFullName) Then
                Return False
            End If

            With Me
                .ConfigFileFullName = ConfigFileFullName
                .DataSet.Clear()
                .DataSet.ReadXml(ConfigFileFullName)
            End With

            Return True
        End Function

        Public Sub Save()
            DataSet.WriteXml(ConfigFileFullName, XmlWriteMode.WriteSchema)
        End Sub

        Public Sub SaveAs(ByVal ConfigFileFullName As String)
            DataSet.WriteXml(ConfigFileFullName, XmlWriteMode.WriteSchema)
        End Sub

        Public Function GetConfig(ByVal TableName As String, ByRef DataTable As DataTable) As Boolean
            For Each Table As DataTable In DataSet.Tables
                If Table.TableName.Trim.ToLower = TableName.Trim.ToLower Then
                    DataTable = Table
                    Return True
                End If
            Next

            Return False
        End Function

        Public Sub SetConfig(ByVal DataTable As DataTable)
            For Each Table As DataTable In DataSet.Tables
                If Table.TableName.Trim.ToLower = DataTable.TableName.Trim.ToLower Then
                    DataSet.Tables.Remove(Table)
                    DataSet.Tables.Add(DataTable)
                    ' Me.Save()
                    Exit Sub
                End If
            Next

            DataSet.Tables.Add(DataTable)
            ' Me.Save()
        End Sub
    End Class
End Namespace