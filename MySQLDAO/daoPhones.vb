Public Class daoPhones
    Private _data As CommonClasses.clsPhone
    Private _conn As MySql.Data.MySqlClient.MySqlConnection
    Private _cmd As MySql.Data.MySqlClient.MySqlCommand
    Private _reader As MySql.Data.MySqlClient.MySqlDataReader
    Private _log As CommonClasses.clsLogHandler

    Public Sub New()
        _data = New CommonClasses.clsPhone
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
            sql = "SELECT MAX(phone_id) AS id FROM phone_numbers"

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
                _log.writeLog(ex, "daoPhones::getNextVal", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function getRecord(ByVal id As Integer) As CommonClasses.clsPhone
        Dim sql As String = ""

        Try
            _data = New CommonClasses.clsPhone

            sql = "SELECT phone_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", phone_type_id" + vbCrLf
            sql += ", country_code" + vbCrLf
            sql += ", area_code" + vbCrLf
            sql += ", prefix" + vbCrLf
            sql += ", pnumber" + vbCrLf
            sql += ", extension" + vbCrLf
            sql += ", modem_dial_string" + vbCrLf
            sql += "FROM phone_numbers " + vbCrLf
            sql += "WHERE phone_id = " + id.ToString

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
                    _data.TypeId = .GetInt16(4)
                    _data.CountryCode = .GetValue(5).ToString
                    _data.AreaCode = .GetValue(6).ToString
                    _data.Prefix = .GetValue(7).ToString
                    _data.PhoneNumber = .GetValue(8).ToString
                    _data.Ext = .GetValue(9).ToString

                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoPhones::getRecord", ex.Message, "error")
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
            sql = "SELECT phone_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", phone_type_id" + vbCrLf
            sql += ", country_code" + vbCrLf
            sql += ", area_code" + vbCrLf
            sql += ", prefix" + vbCrLf
            sql += ", pnumber" + vbCrLf
            sql += ", extension" + vbCrLf
            sql += ", modem_dial_string" + vbCrLf
            sql += "FROM phone_numbers " + vbCrLf
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
                    _data = New CommonClasses.clsPhone

                    _data.Id = CInt(.GetValue(0))
                    _data.DisplayText = .GetValue(1).ToString
                    _data.Description = .GetValue(2).ToString
                    _data.Remarks = .GetValue(3).ToString
                    _data.TypeId = .GetInt16(4)
                    _data.CountryCode = .GetValue(5).ToString
                    _data.AreaCode = .GetValue(6).ToString
                    _data.Prefix = .GetValue(7).ToString
                    _data.PhoneNumber = .GetValue(8).ToString
                    _data.Ext = .GetValue(9).ToString

                    retval.Add(_data)
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoPhones::getRecords", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function insertRecord(ByRef record As CommonClasses.clsPhone) As Boolean
        Dim retval As Boolean = False
        Dim sql As String = ""
        Dim values As String = ""
        Dim insertClause As String = ""
        Dim recordsAffected As Integer

        Try
            insertClause = "INSERT INTO phone_numbers ("
            insertClause += "phone_id" + vbCrLf
            insertClause += ", display_text" + vbCrLf
            insertClause += ", description" + vbCrLf
            insertClause += ", remarks" + vbCrLf
            insertClause += ", phone_type_id" + vbCrLf
            insertClause += ", country_code" + vbCrLf
            insertClause += ", area_code" + vbCrLf
            insertClause += ", prefix" + vbCrLf
            insertClause += ", pnumber" + vbCrLf
            insertClause += ", extension" + vbCrLf
            'insertClause += ", modem_dial_string" + vbCrLf
            insertClause += ")VALUES(" + vbCrLf
            values = getNextVal().ToString + vbCrLf

            With record
                values += ", '" + .DisplayText + "'" + vbCrLf
                values += ", '" + .Description + "'" + vbCrLf
                values += ", '" + .Remarks + "'" + vbCrLf
                values += ", " + .TypeId + vbCrLf
                values += ", " + .CountryCode + vbCrLf
                values += ", " + .AreaCode + vbCrLf
                values += ", " + .Prefix + vbCrLf
                values += ", " + .PhoneNumber + vbCrLf
                values += ", " + .Ext + vbCrLf
                'values += ", '" + .Remarks + "'" + vbCrLf
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
                _log.writeLog(ex, "daoPhones::getRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function updateRecord(ByRef record As CommonClasses.clsPhone) As Boolean
        Dim retval As Boolean = False
        Dim dbRecord As New CommonClasses.clsPhone
        Dim recordsAffected As Integer

        Dim sql As String = ""
        Dim values As String = ""
        Dim setClause As String = ""
        Dim sComma As String = ""

        Try
            dbRecord = getRecord(record.Id)

            sql = "UPDATE phone_numbers SET " + vbCrLf

            '----- if the phone_id does not match, it's an insert not an update.
            If dbRecord.Id <> record.Id Then
                MsgBox("No matching record in the database to update.", MsgBoxStyle.OkOnly, "No Matching Record")
                Return False
            End If

            'setClause = "phone_id = " + vbCrLf
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
            If dbRecord.TypeId <> record.TypeId Then
                setClause += sComma + " phone_type_id = " + record.TypeId.ToString + vbCrLf
                sComma = ","
            End If
            If dbRecord.CountryCode.Equals(record.CountryCode) Then
                setClause += sComma + " country_code = " + record.CountryCode + vbCrLf
                sComma = ","
            End If
            If dbRecord.AreaCode.Equals(record.AreaCode) Then
                setClause += sComma + " area_code = " + record.AreaCode + vbCrLf
                sComma = ","
            End If
            If dbRecord.Prefix.Equals(record.Prefix) Then
                setClause += sComma + " prefix = " + record.Prefix + vbCrLf
                sComma = ","
            End If
            If dbRecord.PhoneNumber.Equals(record.PhoneNumber) Then
                setClause += sComma + " pnumber = " + record.PhoneNumber + vbCrLf
                sComma = ","
            End If
            If dbRecord.Ext.Equals(record.Ext) Then
                setClause += sComma + " extension = " + record.Ext + vbCrLf
                sComma = ","
            End If

            If setClause.Length > 0 Then
                sql += setClause
                sql += " WHERE phone_id = " + record.Id.ToString

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
                _log.writeLog(ex, "daoPhones::updateRecord", ex.Message, "error")
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
            sql = "DELETE FROM phone_numbers WHERE phone_id = " + id.ToString

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
                _log.writeLog(ex, "daoPhones::deleteRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

End Class
