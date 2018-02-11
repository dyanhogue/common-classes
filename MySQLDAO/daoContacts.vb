Public Class daoContacts
    Private _data As CommonClasses.clsContact
    Private _conn As MySql.Data.MySqlClient.MySqlConnection
    Private _cmd As MySql.Data.MySqlClient.MySqlCommand
    Private _reader As MySql.Data.MySqlClient.MySqlDataReader
    Private _log As CommonClasses.clsLogHandler

    Public Sub New()
        _data = New CommonClasses.clsContact
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
            sql = "SELECT MAX(contact_id) AS id FROM contacts"

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
                _log.writeLog(ex, "daoContacts::getNextVal", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function getRecord(ByVal id As Integer) As CommonClasses.clsContact
        Dim sql As String = ""

        Try
            _data = New CommonClasses.clsContact

            sql = "SELECT contact_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", contact_type_id" + vbCrLf
            sql += ", first_name" + vbCrLf
            sql += ", last_name" + vbCrLf
            sql += ", middle_name" + vbCrLf
            sql += ", prefix" + vbCrLf
            sql += ", suffix " + vbCrLf
            sql += "FROM contacts " + vbCrLf
            sql += "WHERE contact_id = " + id.ToString

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
                    _data.ContactTypeId = CInt(.GetValue(4))
                    _data.FirstName = .GetValue(5).ToString
                    _data.LastName = .GetValue(6).ToString
                    _data.MiddleName = .GetValue(7).ToString
                    _data.Prefix = .GetValue(8).ToString
                    _data.Suffix = .GetValue(9).ToString
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoContacts::getRecord", ex.Message, "error")
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
            sql = "SELECT contact_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", contact_type_id" + vbCrLf
            sql += ", first_name" + vbCrLf
            sql += ", last_name" + vbCrLf
            sql += ", middle_name" + vbCrLf
            sql += ", prefix" + vbCrLf
            sql += ", suffix " + vbCrLf
            sql += "FROM contacts " + vbCrLf
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
                    _data = New CommonClasses.clsContact

                    _data.Id = CInt(.GetValue(0))
                    _data.DisplayText = .GetValue(1).ToString
                    _data.Description = .GetValue(2).ToString
                    _data.Remarks = .GetValue(3).ToString
                    _data.ContactTypeId = CInt(.GetValue(4))
                    _data.FirstName = .GetValue(5).ToString
                    _data.LastName = .GetValue(6).ToString
                    _data.MiddleName = .GetValue(7).ToString
                    _data.Prefix = .GetValue(8).ToString
                    _data.Suffix = .GetValue(9).ToString

                    retval.Add(_data)
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoContacts::getRecords", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function insertRecord(ByRef record As CommonClasses.clsContact) As Boolean
        Dim retval As Boolean = False
        Dim sql As String = ""
        Dim values As String = ""
        Dim insertClause As String = ""
        Dim recordsAffected As Integer

        Try
            insertClause = "INSERT INTO contacts ("
            insertClause += "contact_id" + vbCrLf
            insertClause += ", display_text" + vbCrLf
            insertClause += ", description" + vbCrLf
            insertClause += ", remarks" + vbCrLf
            insertClause += ", contact_type_id" + vbCrLf
            insertClause += ", first_name" + vbCrLf
            insertClause += ", last_name" + vbCrLf
            insertClause += ", middle_name" + vbCrLf
            insertClause += ", prefix" + vbCrLf
            insertClause += ", suffix " + vbCrLf
            insertClause += ")VALUES(" + vbCrLf
            values = getNextVal().ToString + vbCrLf

            With record
                values += ", '" + .DisplayText + "'" + vbCrLf
                values += ", '" + .Description + "'" + vbCrLf
                values += ", '" + .Remarks + "'" + vbCrLf
                values += ", '" + .ContactTypeId + vbCrLf
                values += ", '" + .FirstName + "'" + vbCrLf
                values += ", '" + .LastName + "'" + vbCrLf
                values += ", '" + .MiddleName + "'" + vbCrLf
                values += ", '" + .Prefix + "'" + vbCrLf
                values += ", '" + .Suffix + "'" + vbCrLf
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
                _log.writeLog(ex, "daoContacts::getRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function updateRecord(ByRef record As CommonClasses.clsContact) As Boolean
        Dim retval As Boolean = False
        Dim dbRecord As New CommonClasses.clsContact
        Dim recordsAffected As Integer

        Dim sql As String = ""
        Dim values As String = ""
        Dim setClause As String = ""
        Dim sComma As String = ""

        Try
            dbRecord = getRecord(record.Id)

            sql = "UPDATE contacts SET " + vbCrLf

            '----- if the contact_id does not match, it's an insert not an update.
            If dbRecord.Id <> record.Id Then
                MsgBox("No matching record in the database to update.", MsgBoxStyle.OkOnly, "No Matching Record")
                Return False
            End If

            'setClause = "contact_id = " + vbCrLf
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
            If dbRecord.ContactTypeId = record.ContactTypeId Then
                setClause += sComma + "contact_type_id = " + record.Remarks + vbCrLf
                sComma = ","
            End If
            If dbRecord.FirstName.Equals(record.FirstName) Then
                setClause += sComma + " first_name = '" + record.FirstName + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.LastName.Equals(record.LastName) Then
                setClause += sComma + " last_name = '" + record.LastName + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.MiddleName.Equals(record.MiddleName) Then
                setClause += sComma + " middle_name = '" + record.MiddleName + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Prefix.Equals(record.Prefix) Then
                setClause += sComma + " prefix = '" + record.Prefix + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Suffix.Equals(record.Suffix) Then
                setClause += sComma + " suffix = '" + record.Suffix + "'" + vbCrLf
                sComma = ","
            End If

            If setClause.Length > 0 Then
                sql += setClause
                sql += " WHERE contact_id = " + record.Id.ToString

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
                _log.writeLog(ex, "daoContacts::updateRecord", ex.Message, "error")
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
            sql = "DELETE FROM contacts WHERE contact_id = " + id.ToString

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
                _log.writeLog(ex, "daoContacts::deleteRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

End Class
