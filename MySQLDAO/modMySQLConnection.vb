Public Module modMySQLConnection
    Private Const _CONNECT As String = "server=Athena;user id=local_app;pwd=apppw;persistsecurityinfo=True;database=local_data"

    Public Function getConnection() As MySql.Data.MySqlClient.MySqlConnection
        Dim retval As New MySql.Data.MySqlClient.MySqlConnection

        Try
            retval.ConnectionString = _CONNECT
            retval.Open()

        Catch ex As Exception
            MsgBox("An error has occurred in clsMySQLConnection::getConnection" + vbCrLf + ex.Message, vbOKOnly, "Connection Error")
            retval = New MySql.Data.MySqlClient.MySqlConnection
        End Try

        Return retval
    End Function

End Module
