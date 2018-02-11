Public Class daoState

    Private _data As CommonClasses.clsState
    Private _conn As MySql.Data.MySqlClient.MySqlConnection
    Private _cmd As MySql.Data.MySqlClient.MySqlCommand
    Private _reader As MySql.Data.MySqlClient.MySqlDataReader
    Private _log As CommonClasses.clsLogHandler

    Public Sub New()
        _data = New CommonClasses.clsState
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
            sql = "SELECT MAX(id) AS id FROM state"

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
                _log.writeLog(ex, "daoState::getNextVal", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function getRecord(ByVal id As Integer) As CommonClasses.clsState
        Dim sql As String = ""

        Try
            _data = New CommonClasses.clsState

            sql = "SELECT id" + vbCrLf
            sql += ", name" + vbCrLf
            sql += ", abbreviation" + vbCrLf
            sql += ", country" + vbCrLf
            sql += ", type" + vbCrLf
            sql += ", sort" + vbCrLf
            sql += ", status" + vbCrLf
            sql += ", occupied" + vbCrLf
            sql += ", notes" + vbCrLf
            sql += ", fips_state" + vbCrLf
            sql += ", assoc_press" + vbCrLf
            sql += ", standard_federal_region" + vbCrLf
            sql += ", census_region" + vbCrLf
            sql += ", census_region_name" + vbCrLf
            sql += ", census_division" + vbCrLf
            sql += ", census_division_name" + vbCrLf
            sql += ", circuit_court" + vbCrLf
            sql += "FROM state " + vbCrLf
            sql += "WHERE id = " + id.ToString

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
                    _data.StateName = .GetValue(1).ToString
                    _data.StateCode = .GetValue(2).ToString
                    _data.Country = .GetValue(3).ToString
                    _data.StateType = .GetValue(4).ToString
                    _data.Sort = .GetInt16(5).ToString
                    _data.Status = .GetValue(6).ToString
                    _data.Occupied = .GetValue(7).ToString
                    _data.Notes = .GetValue(8).ToString
                    _data.FipsState = .GetValue(9).ToString
                    _data.AssocPress = .GetValue(10).ToString
                    _data.StandardFedRegion = .GetValue(11).ToString
                    _data.CensusRegion = .GetValue(12).ToString
                    _data.CensusRegionName = .GetValue(13).ToString
                    _data.CensusDivision = .GetValue(14).ToString
                    _data.CensusDivisionName = .GetValue(15).ToString
                    _data.CircuitCourt = .GetValue(16).ToString
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoState::getRecord", ex.Message, "error")
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
            sql = "SELECT id" + vbCrLf
            sql += ", name" + vbCrLf
            sql += ", abbreviation" + vbCrLf
            sql += ", country" + vbCrLf
            sql += ", type" + vbCrLf
            sql += ", sort" + vbCrLf
            sql += ", status" + vbCrLf
            sql += ", occupied" + vbCrLf
            sql += ", notes" + vbCrLf
            sql += ", fips_state" + vbCrLf
            sql += ", assoc_press" + vbCrLf
            sql += ", standard_federal_region" + vbCrLf
            sql += ", census_region" + vbCrLf
            sql += ", census_region_name" + vbCrLf
            sql += ", census_division" + vbCrLf
            sql += ", census_division_name" + vbCrLf
            sql += ", circuit_court" + vbCrLf
            sql += "FROM state " + vbCrLf
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
                    _data = New CommonClasses.clsState

                    _data.Id = CInt(.GetValue(0))
                    _data.StateName = .GetValue(1).ToString
                    _data.StateCode = .GetValue(2).ToString
                    _data.Country = .GetValue(3).ToString
                    _data.StateType = .GetValue(4).ToString
                    _data.Sort = .GetInt16(5).ToString
                    _data.Status = .GetValue(6).ToString
                    _data.Occupied = .GetValue(7).ToString
                    _data.Notes = .GetValue(8).ToString
                    _data.FipsState = .GetValue(9).ToString
                    _data.AssocPress = .GetValue(10).ToString
                    _data.StandardFedRegion = .GetValue(11).ToString
                    _data.CensusRegion = .GetValue(12).ToString
                    _data.CensusRegionName = .GetValue(13).ToString
                    _data.CensusDivision = .GetValue(14).ToString
                    _data.CensusDivisionName = .GetValue(15).ToString
                    _data.CircuitCourt = .GetValue(16).ToString

                    retval.Add(_data)
                End While
            End With
        Catch ex As Exception
            If Not _log Is Nothing Then
                _log.writeLog(ex, "daoState::getRecords", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function insertRecord(ByRef record As CommonClasses.clsState) As Boolean
        Dim retval As Boolean = False
        Dim sql As String = ""
        Dim values As String = ""
        Dim insertClause As String = ""
        Dim recordsAffected As Integer

        Try
            insertClause = "INSERT INTO state ("
            insertClause += "id" + vbCrLf
            insertClause += ", name" + vbCrLf
            insertClause += ", abbreviation" + vbCrLf
            insertClause += ", country" + vbCrLf
            insertClause += ", type" + vbCrLf
            insertClause += ", sort" + vbCrLf
            insertClause += ", status" + vbCrLf
            insertClause += ", occupied" + vbCrLf
            insertClause += ", notes" + vbCrLf
            insertClause += ", fips_state" + vbCrLf
            insertClause += ", assoc_press" + vbCrLf
            insertClause += ", standard_federal_region" + vbCrLf
            insertClause += ", census_region" + vbCrLf
            insertClause += ", census_region_name" + vbCrLf
            insertClause += ", census_division" + vbCrLf
            insertClause += ", census_division_name" + vbCrLf
            insertClause += ", circuit_court" + vbCrLf
            insertClause += ")VALUES(" + vbCrLf
            values = getNextVal().ToString + vbCrLf

            With record
                'values += ", '" + .DisplayText + "'" + vbCrLf
                'values += ", '" + .Description + "'" + vbCrLf
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
                _log.writeLog(ex, "daoState::getRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

    Public Function updateRecord(ByRef record As CommonClasses.clsState) As Boolean
        Dim retval As Boolean = False
        Dim dbRecord As New CommonClasses.clsState
        Dim recordsAffected As Integer

        Dim sql As String = ""
        Dim values As String = ""
        Dim setClause As String = ""
        Dim sComma As String = ""

        Try
            dbRecord = getRecord(record.Id)

            sql = "UPDATE state SET " + vbCrLf

            '----- if the id does not match, it's an insert not an update.
            If dbRecord.Id <> record.Id Then
                MsgBox("No matching record in the database to update.", MsgBoxStyle.OkOnly, "No Matching Record")
                Return False
            End If

            'setClause = "id = " + vbCrLf
            'setClause += ", account_number = " + vbCrLf
            'If dbRecord.DisplayText.Equals(record.DisplayText) Then
            '    setClause += sComma + " display_text = '" + record.DisplayText + "'" + vbCrLf
            '    sComma = ","
            'End If
            'If dbRecord.Description.Equals(record.Description) Then
            '    setClause += sComma + " description = '" + record.Description + "'" + vbCrLf
            '    sComma = ","
            'End If
            'If dbRecord.Remarks.Equals(record.Remarks) Then
            '    setClause += sComma + " remarks = '" + record.Remarks + "'" + vbCrLf
            '    sComma = ","
            'End If

            If setClause.Length > 0 Then
                sql += setClause
                sql += " WHERE id = " + record.Id.ToString

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
                _log.writeLog(ex, "daoState::updateRecord", ex.Message, "error")
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
            sql = "DELETE FROM state WHERE id = " + id.ToString

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
                _log.writeLog(ex, "daoState::deleteRecord", ex.Message, "error")
                _log.writeLog(vbTab + "SQL:  " + sql, "error")
                Throw ex
            End If
        End Try

        Return retval
    End Function

End Class
