Public Class daoAddress
    Private _data As CommonClasses.clsAddress
    Private _conn As MySql.Data.MySqlClient.MySqlConnection
    Private _cmd As MySql.Data.MySqlClient.MySqlCommand
    Private _reader As MySql.Data.MySqlClient.MySqlDataReader
    Private _log As CommonClasses.clsLogHandler

    Public Sub New()
        _data = New CommonClasses.clsAddress
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
            sql = "SELECT MAX(address_id) AS id FROM addresses"

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
                _log.writeLog(ex, "daoAddress::getNextVal", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function getRecord(ByVal id As Integer) As CommonClasses.clsAddress
        Dim sql As String = ""

        Try
            _data = New CommonClasses.clsAddress

            sql = "SELECT address_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", address_type_id" + vbCrLf
            sql += ", line1" + vbCrLf
            sql += ", line2" + vbCrLf
            sql += ", line3" + vbCrLf
            sql += ", line4" + vbCrLf
            sql += ", city" + vbCrLf
            sql += ", state" + vbCrLf
            sql += ", state_code" + vbCrLf
            sql += ", country" + vbCrLf
            sql += ", country_code" + vbCrLf
            sql += ", postal_code " + vbCrLf
            sql += "FROM addresses " + vbCrLf
            sql += "WHERE address_id = " + id.ToString

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
                    _data.AddressTypeId = .GetInt16(4)
                    _data.Line1 = .GetValue(5).ToString
                    _data.Line2 = .GetValue(6).ToString
                    _data.Line3 = .GetValue(7).ToString
                    _data.Line4 = .GetValue(8).ToString
                    _data.City = .GetValue(9).ToString
                    _data.State = .GetValue(10).ToString
                    _data.StateCode = .GetValue(11).ToString
                    _data.Country = .GetValue(12).ToString
                    _data.CountryCode = .GetValue(13).ToString
                    _data.PostalCode = .GetValue(14).ToString
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoaddresses::getRecord", ex.Message, "error")
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
            sql = "SELECT address_id" + vbCrLf
            sql += ", display_text" + vbCrLf
            sql += ", description" + vbCrLf
            sql += ", remarks" + vbCrLf
            sql += ", address_type_id" + vbCrLf
            sql += ", line1" + vbCrLf
            sql += ", line2" + vbCrLf
            sql += ", line3" + vbCrLf
            sql += ", line4" + vbCrLf
            sql += ", city" + vbCrLf
            sql += ", state" + vbCrLf
            sql += ", state_code" + vbCrLf
            sql += ", country" + vbCrLf
            sql += ", country_code" + vbCrLf
            sql += ", postal_code " + vbCrLf
            sql += "FROM addresses " + vbCrLf
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
                    _data = New CommonClasses.clsAddress

                    _data.Id = CInt(.GetValue(0))
                    _data.DisplayText = .GetValue(1).ToString
                    _data.Description = .GetValue(2).ToString
                    _data.Remarks = .GetValue(3).ToString
                    _data.AddressTypeId = .GetInt16(4)
                    _data.Line1 = .GetValue(5).ToString
                    _data.Line2 = .GetValue(6).ToString
                    _data.Line3 = .GetValue(7).ToString
                    _data.Line4 = .GetValue(8).ToString
                    _data.City = .GetValue(9).ToString
                    _data.State = .GetValue(10).ToString
                    _data.StateCode = .GetValue(11).ToString
                    _data.Country = .GetValue(12).ToString
                    _data.CountryCode = .GetValue(13).ToString
                    _data.PostalCode = .GetValue(14).ToString

                    retval.Add(_data)
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoaddresses::getRecords", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function insertRecord(ByRef record As CommonClasses.clsAddress) As Boolean
        Dim retval As Boolean = False
        Dim sql As String = ""
        Dim values As String = ""
        Dim insertClause As String = ""
        Dim recordsAffected As Integer

        Try
            insertClause = "INSERT INTO addresses ("
            insertClause += "address_id" + vbCrLf
            insertClause += ", display_text" + vbCrLf
            insertClause += ", description" + vbCrLf
            insertClause += ", remarks" + vbCrLf
            insertClause += ", address_type_id" + vbCrLf
            insertClause += ", line1" + vbCrLf
            insertClause += ", line2" + vbCrLf
            insertClause += ", line3" + vbCrLf
            insertClause += ", line4" + vbCrLf
            insertClause += ", city" + vbCrLf
            insertClause += ", state" + vbCrLf
            insertClause += ", state_code" + vbCrLf
            insertClause += ", country" + vbCrLf
            insertClause += ", country_code" + vbCrLf
            insertClause += ", postal_code " + vbCrLf
            insertClause += ")VALUES(" + vbCrLf
            values = getNextVal().ToString + vbCrLf

            With record
                values += ", '" + .DisplayText + "'" + vbCrLf
                values += ", '" + .Description + "'" + vbCrLf
                values += ", '" + .Remarks + "'" + vbCrLf
                values += ", '" + .AddressTypeId + vbCrLf
                values += ", '" + .Line1 + "'" + vbCrLf
                values += ", '" + .Line2 + "'" + vbCrLf
                values += ", '" + .Line3 + "'" + vbCrLf
                values += ", '" + .Line4 + "'" + vbCrLf
                values += ", '" + .City + "'" + vbCrLf
                values += ", '" + .State + "'" + vbCrLf
                values += ", '" + .StateCode + "'" + vbCrLf
                values += ", '" + .Country + "'" + vbCrLf
                values += ", '" + .CountryCode + "'" + vbCrLf
                values += ", '" + .PostalCode + "'" + vbCrLf
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
                _log.writeLog(ex, "daoaddresses::getRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function updateRecord(ByRef record As CommonClasses.clsAddress) As Boolean
        Dim retval As Boolean = False
        Dim dbRecord As New CommonClasses.clsAddress
        Dim recordsAffected As Integer

        Dim sql As String = ""
        Dim values As String = ""
        Dim setClause As String = ""
        Dim sComma As String = ""

        Try
            dbRecord = getRecord(record.Id)

            sql = "UPDATE addresses SET " + vbCrLf

            '----- if the address_id does not match, it's an insert not an update.
            If dbRecord.Id <> record.Id Then
                MsgBox("No matching record in the database to update.", MsgBoxStyle.OkOnly, "No Matching Record")
                Return False
            End If

            'setClause = "address_id = " + vbCrLf
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
            If dbRecord.AddressTypeId <> record.AddressTypeId Then
                setClause += sComma + " address_type_id = " + record.AddressTypeId + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Line1.Equals(record.Line1) Then
                setClause += sComma + " line1 = '" + record.Line1 + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Line2.Equals(record.Line2) Then
                setClause += sComma + " line2 = '" + record.Line2 + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Line3.Equals(record.Line3) Then
                setClause += sComma + " line3 = '" + record.Line3 + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Line4.Equals(record.Line4) Then
                setClause += sComma + " line4 = '" + record.Line4 + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.City.Equals(record.City) Then
                setClause += sComma + " city = '" + record.City + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.State.Equals(record.State) Then
                setClause += sComma + " state = '" + record.State + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.StateCode.Equals(record.StateCode) Then
                setClause += sComma + " state_code = '" + record.StateCode + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.Country.Equals(record.Country) Then
                setClause += sComma + " country = '" + record.Country + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.CountryCode.Equals(record.CountryCode) Then
                setClause += sComma + " country_code = '" + record.CountryCode + "'" + vbCrLf
                sComma = ","
            End If
            If dbRecord.PostalCode.Equals(record.PostalCode) Then
                setClause += sComma + " postal_code = '" + record.PostalCode + "'" + vbCrLf
                sComma = ","
            End If

            If setClause.Length > 0 Then
                sql += setClause
                sql += " WHERE address_id = " + record.Id.ToString

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
                _log.writeLog(ex, "daoaddresses::updateRecord", ex.Message, "error")
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
            sql = "DELETE FROM addresses WHERE address_id = " + id.ToString

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
                _log.writeLog(ex, "daoaddresses::deleteRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

End Class
