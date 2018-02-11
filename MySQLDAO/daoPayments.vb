Public Class daoPayments
    Private _data As CommonClasses.clsPayment
    Private _conn As MySql.Data.MySqlClient.MySqlConnection
    Private _cmd As MySql.Data.MySqlClient.MySqlCommand
    Private _reader As MySql.Data.MySqlClient.MySqlDataReader
    Private _log As CommonClasses.clsLogHandler

    Public Sub New()
        _data = New CommonClasses.clsPayment
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
            sql = "SELECT MAX(payment_id) AS id FROM payments"

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
                _log.writeLog(ex, "daoPayments::getNextVal", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function getRecord(ByVal id As Integer) As CommonClasses.clsPayment
        Dim sql As String = ""

        Try
            _data = New CommonClasses.clsPayment

            sql = "SELECT payment_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", account_id" + vbCrLf
            sql += ", day_due" + vbCrLf
            sql += ", amount" + vbCrLf
            sql += "FROM payments " + vbCrLf
            sql += "WHERE payment_id = " + id.ToString

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
                    _data.DisplayText = .GetValue(1).ToString
                    _data.Description = .GetValue(2).ToString
                    _data.Remarks = .GetValue(3).ToString
                    _data.AccountId = .GetInt16(4)
                    _data.DueDay = .GetInt16(5)
                    _data.PaymentAmount = .GetDouble(6)

                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoPayments::getRecord", ex.Message, "error")
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
            sql = "SELECT payment_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", account_id" + vbCrLf
            sql += ", day_due" + vbCrLf
            sql += ", amount" + vbCrLf
            sql += "FROM payments " + vbCrLf
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
                    _data = New CommonClasses.clsPayment

                    _data.Id = CInt(.GetValue(0))
                    _data.DisplayText = .GetValue(1).ToString
                    _data.Description = .GetValue(2).ToString
                    _data.Remarks = .GetValue(3).ToString
                    _data.AccountId = .GetInt16(4)
                    _data.DueDay = .GetInt16(5)
                    _data.PaymentAmount = .GetDouble(6)

                    retval.Add(_data)
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoPayments::getRecords", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function insertRecord(ByRef record As CommonClasses.clsPayment) As Boolean
        Dim retval As Boolean = False
        Dim sql As String = ""
        Dim values As String = ""
        Dim insertClause As String = ""
        Dim recordsAffected As Integer

        Try
            insertClause = "INSERT INTO payments ("
            insertClause += "payment_id" + vbCrLf
            insertClause += ", display_text" + vbCrLf
            insertClause += ", description" + vbCrLf
            insertClause += ", remarks" + vbCrLf
            insertClause += ", account_id" + vbCrLf
            insertClause += ", day_due" + vbCrLf
            insertClause += ", amount" + vbCrLf
            insertClause += ")VALUES(" + vbCrLf
            values = getNextVal().ToString + vbCrLf

            With record
                values += ", '" + .DisplayText + "'" + vbCrLf
                values += ", '" + .Description + "'" + vbCrLf
                values += ", '" + .Remarks + "'" + vbCrLf
                values += ", " + .AccountId + vbCrLf
                values += ", " + .DueDay + vbCrLf
                values += ", " + .PaymentAmount + vbCrLf
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
                _log.writeLog(ex, "daoPayments::getRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function updateRecord(ByRef record As CommonClasses.clsPayment) As Boolean
        Dim retval As Boolean = False
        Dim dbRecord As New CommonClasses.clsPayment
        Dim recordsAffected As Integer

        Dim sql As String = ""
        Dim values As String = ""
        Dim setClause As String = ""
        Dim sComma As String = ""

        Try
            dbRecord = getRecord(record.Id)

            sql = "UPDATE payments SET " + vbCrLf

            '----- if the payment_id does not match, it's an insert not an update.
            If dbRecord.Id <> record.Id Then
                MsgBox("No matching record in the database to update.", MsgBoxStyle.OkOnly, "No Matching Record")
                Return False
            End If

            'setClause = "payment_id = " + vbCrLf
            'setClause += ", account_number = " + vbCrLf
            If dbRecord.DisplayText.Equals(record.DisplayText) Then
                setClause += sComma + " display_text = '" + record.DisplayText + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Description.Equals(record.Description) Then
                setClause += sComma + " description = '" + record.Description + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Remarks.Equals(record.Remarks) Then
                setClause += sComma + " remarks = '" + record.Remarks + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.AccountId <> record.AccountId Then
                setClause += sComma + " account_id = " + record.AccountId + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.DueDay <> record.DueDay Then
                setClause += sComma + " day_due = " + record.DueDay + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.PaymentAmount <> record.PaymentAmount Then
                setClause += sComma + " amount = " + record.PaymentAmount + "'" + vbCrLf
                sComma = ","
            End If

            If setClause.Length > 0 Then
                sql += setClause
                sql += " WHERE payment_id = " + record.Id.ToString

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
                _log.writeLog(ex, "daoPayments::updateRecord", ex.Message, "error")
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
            sql = "DELETE FROM payments WHERE payment_id = " + id.ToString

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
                _log.writeLog(ex, "daoPayments::deleteRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

End Class
