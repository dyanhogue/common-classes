Public Class daoAccounts

    Private _data As CommonClasses.clsAccount
    Private _conn As MySql.Data.MySqlClient.MySqlConnection
    Private _cmd As MySql.Data.MySqlClient.MySqlCommand
    Private _reader As MySql.Data.MySqlClient.MySqlDataReader
    Private _log As CommonClasses.clsLogHandler

    Public Sub New()
        _data = New CommonClasses.clsAccount
        _conn = New MySql.Data.MySqlClient.MySqlConnection
        _log = Nothing

    End Sub

    Public WriteOnly Property Conn As MySql.Data.MySqlClient.MySqlConnection
        Set(value As MySql.Data.MySqlClient.MySqlConnection)
            _conn = value
        End Set
    End Property

    Public WriteOnly Property Logger As CommonClasses.clsLogHandler
        Set(value As CommonClasses.clsLogHandler)
            _log = value
        End Set
    End Property

    Private Function getNextVal() As Integer
        Dim retval As Integer = 0
        Dim sql As String = ""

        Try
            sql = "SELECT MAX(account_id) AS id FROM accounts"

            _cmd = New MySql.Data.MySqlClient.MySqlCommand

            With _cmd
                .CommandType = CommandType.Text
                .Connection = _conn
                .CommandText = sql
                _reader = .ExecuteReader
            End With

            While _reader.Read
                retval = _reader.GetInt16(0) + 1
            End While

        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoAccounts::getNextVal", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function getRecord(ByVal id As Integer) As CommonClasses.clsAccount
        Dim sql As String = ""

        Try
            _data = New CommonClasses.clsAccount

            sql = "SELECT account_id" + vbCrLf
            sql += ", account_number" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", company_name" + vbCrLf
            sql += ", account_type_id" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", status " + vbCrLf
            sql += "FROM accounts " + vbCrLf
            sql += "WHERE account_id = " + id.ToString

            _cmd = New MySql.Data.MySqlClient.MySqlCommand

            With _cmd
                .CommandType = CommandType.Text
                .Connection = _conn
                .CommandText = sql
                _reader = .ExecuteReader
            End With

            With _reader
                While .Read
                    _data.Id = CInt(.GetValue(0))
                    _data.AccountNumber = .GetValue(1).ToString
                    _data.DisplayText = .GetValue(2).ToString
                    _data.Description = .GetValue(3).ToString
                    _data.AccountName = .GetValue(4).ToString
                    _data.TypeId = CInt(.GetValue(5))
                    _data.Remarks = .GetValue(6).ToString
                    _data.Status = .GetValue(7).ToString
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoAccounts::getRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return _data
    End Function

    Public Function getRecords(Optional ByVal psWhere As String = "") As Collection
        Dim retval As New Collection
        Dim sql As String = ""

        Try
            sql = "SELECT account_id" + vbCrLf
            sql += ", account_number" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", company_name" + vbCrLf
            sql += ", account_type_id" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", status " + vbCrLf
            sql += "FROM accounts " + vbCrLf
            If psWhere.Length > 0 Then
                sql += "WHERE " + psWhere + " " + vbCrLf
            End If
            sql += "ORDER BY display_text "

            _cmd = New MySql.Data.MySqlClient.MySqlCommand
            retval = New Collection

            With _cmd
                .CommandType = CommandType.Text
                .Connection = _conn
                .CommandText = sql
                _reader = .ExecuteReader
            End With

            With _reader
                While .Read
                    _data = New CommonClasses.clsAccount

                    _data.Id = CInt(.GetValue(0))
                    _data.AccountNumber = .GetValue(1).ToString
                    _data.DisplayText = .GetValue(2).ToString
                    _data.Description = .GetValue(3).ToString
                    _data.AccountName = .GetValue(4).ToString
                    _data.TypeId = CInt(.GetValue(5))
                    _data.Remarks = .GetValue(6).ToString
                    _data.Status = .GetValue(7).ToString

                    retval.Add(_data, _data.AccountNumber)
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoAccounts::getRecords", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function insertRecord(ByRef record As CommonClasses.clsAccount) As Boolean
        Dim retval As Boolean = False
        Dim sql As String = ""
        Dim values As String = ""
        Dim insertClause As String = ""
        Dim recordsAffected As Integer

        Try
            insertClause = "INSERT INTO accounts ("
            insertClause += "account_id" + vbCrLf
            insertClause += ", account_number" + vbCrLf
            insertClause += ", display_text" + vbCrLf
            insertClause += ", description" + vbCrLf
            insertClause += ", company_name" + vbCrLf
            insertClause += ", account_type_id" + vbCrLf
            insertClause += ", remarks" + vbCrLf
            insertClause += ", status " + vbCrLf
            insertClause += ")VALUES(" + vbCrLf
            values = getNextVal().ToString + vbCrLf

            With record
                values += ", '" + .AccountNumber + "'" + vbCrLf
                values += ", '" + .DisplayText + "'" + vbCrLf
                values += ", '" + .Description + "'" + vbCrLf
                values += ", '" + .AccountName + "'" + vbCrLf
                values += ", " + .TypeId.ToString + vbCrLf
                values += ", '" + .Remarks + "'" + vbCrLf
                values += ", '" + .Status + "'" + vbCrLf
            End With

            sql = insertClause + values + ")"

            _cmd = New MySql.Data.MySqlClient.MySqlCommand

            With _cmd
                .CommandType = CommandType.Text
                .Connection = _conn
                .CommandText = sql
                recordsAffected = .ExecuteNonQuery
            End With

            If recordsAffected > -1 Then
                retval = True
            Else
                retval = False
            End If
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoAccounts::getRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function updateRecord(ByRef record As CommonClasses.clsAccount) As Boolean
        Dim retval As Boolean = False
        Dim dbRecord As New CommonClasses.clsAccount
        Dim recordsAffected As Integer

        Dim sql As String = ""
        Dim values As String = ""
        Dim setClause As String = ""
        Dim sComma As String = ""

        Try
            dbRecord = getRecord(record.Id)

            sql = "UPDATE accounts SET " + vbCrLf

            '----- if the account_id or account number do not match, it's an insert not an update.
            If dbRecord.Id <> record.Id Or dbRecord.AccountNumber <> record.AccountNumber Then
                MsgBox("No matching record in the database to update.", MsgBoxStyle.OkOnly, "No Matching Record")
                Return False
            End If

            'setClause = "account_id = " + vbCrLf
            'setClause += ", account_number = " + vbCrLf
            If dbRecord.DisplayText.Equals(record.DisplayText) Then
                setClause += sComma + " display_text = '" + record.DisplayText + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Description.Equals(record.Description) Then
                setClause += sComma + " description = '" + record.Description + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.AccountName.Equals(record.AccountName) Then
                setClause += sComma + " company_name = '" + record.AccountName + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.TypeId.Equals(record.TypeId) Then
                setClause += sComma + " account_type_id = " + record.TypeId + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Remarks.Equals(record.Remarks) Then
                setClause += sComma + " remarks = '" + record.Remarks + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Status.Equals(record.Status) Then
                setClause += sComma + " status = '" + record.Status + "'" + vbCrLf
                sComma = ","
            End If

            If setClause.Length > 0 Then
                sql += setClause
                sql += " WHERE account_id = " + record.Id.ToString

                _cmd = New MySql.Data.MySqlClient.MySqlCommand

                With _cmd
                    .CommandType = CommandType.Text
                    .Connection = _conn
                    .CommandText = sql
                    recordsAffected = .ExecuteNonQuery
                End With

                If recordsAffected > -1 Then
                    retval = True
                Else
                    retval = False
                End If
            Else
                retval = True       '----- if there are no differences between the database record and the update values, return true
            End If

        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoAccounts::updateRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try


        Return retval
    End Function

    Public Function deleteRecord(ByVal id As Integer) As Boolean
        Dim retval As Boolean = False
        Dim sql As String = ""
        Dim recordsAffected As Integer

        Try
            sql = "DELETE FROM accounts WHERE account_id = " + id.ToString

            _cmd = New MySql.Data.MySqlClient.MySqlCommand

            With _cmd
                .CommandType = CommandType.Text
                .Connection = _conn
                .CommandText = sql
                recordsAffected = .ExecuteNonQuery
            End With

            If recordsAffected > -1 Then
                retval = True
            Else
                retval = False
            End If

        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoAccounts::deleteRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function
End Class
