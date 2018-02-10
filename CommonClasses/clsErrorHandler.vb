Public Class clsErrorHandler
    Private _log_filename As String
    Private _warning_level As ErrorWarningLevel
    Private _log_handler As clsLogHandler

    Public Property LogFilename As String
        Set(ByVal value As String)
            _log_filename = value
            _log_handler.LogFilename = value
        End Set
        Get
            Return _log_filename
        End Get
    End Property

    Public Property WarningLevel As ErrorWarningLevel
        Set(ByVal value As ErrorWarningLevel)
            _warning_level = value
        End Set
        Get
            Return _warning_level
        End Get
    End Property

    Public WriteOnly Property LogHandler As clsLogHandler
        Set(ByVal value As clsLogHandler)
            _log_handler = value
        End Set
    End Property

    Public Enum ErrorWarningLevel As Integer
        INFO
        TRACE
        DEBUG
        WARN
        FATAL
    End Enum

    Public Sub New()
        _log_filename = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "\utillog.log"
        _warning_level = ErrorWarningLevel.INFO
        _log_handler = New clsLogHandler
        _log_handler.LogFilename = _log_filename
    End Sub

    Public Sub DisplaySQLError(ByRef ex As Exception, ByVal sql As String, Optional ByVal source As String = "", Optional ByVal warning_level As ErrorWarningLevel = ErrorWarningLevel.INFO, Optional ByVal title As String = "SQL Error")
        Dim msg As String = ""
        Dim log_level As String = ""
        Dim log_msg As String = ""
        Dim style As MsgBoxStyle

        Select Case warning_level
            Case ErrorWarningLevel.DEBUG
                log_level = "DEBUG"
            Case ErrorWarningLevel.FATAL
                log_level = "FATAL"
            Case ErrorWarningLevel.INFO
                log_level = "INFO"
            Case ErrorWarningLevel.TRACE
                log_level = "TRACE"
            Case ErrorWarningLevel.WARN
                log_level = "WARN"
        End Select

        If warning_level = ErrorWarningLevel.FATAL Then
            msg = " A fatal error has occured executing a query"
            style = MsgBoxStyle.Critical
        Else
            msg = "An error has occured executing a query"
            If warning_level = ErrorWarningLevel.WARN Then
                style = MsgBoxStyle.Exclamation
            Else
                style = MsgBoxStyle.Information
            End If
        End If

        If source.Length > 0 Then
            msg += " in " & source & vbCrLf
            log_msg += "Source:  " + source + vbCrLf
        Else
            msg += vbCrLf
        End If

        msg += "Exception Message:  " & ex.Message & vbCrLf
        msg += "SQL:  " & sql
        log_msg += "Exception Message:  " & ex.Message & vbCrLf
        log_msg += "SQL:  " & sql

        _log_handler.writeLog(log_msg, log_level)
        MsgBox(msg, style, title)

    End Sub

    Public Sub DisplayError(ByRef ex As Exception, Optional ByVal source As String = "", Optional ByVal warning_level As ErrorWarningLevel = ErrorWarningLevel.INFO, Optional ByVal title As String = "Error")
        Dim msg As String = ""
        Dim log_level As String = ""
        Dim log_msg As String = ""
        Dim style As MsgBoxStyle

        Select Case warning_level
            Case ErrorWarningLevel.DEBUG
                log_level = "DEBUG"
            Case ErrorWarningLevel.FATAL
                log_level = "FATAL"
            Case ErrorWarningLevel.INFO
                log_level = "INFO"
            Case ErrorWarningLevel.TRACE
                log_level = "TRACE"
            Case ErrorWarningLevel.WARN
                log_level = "WARN"
        End Select

        If warning_level = ErrorWarningLevel.FATAL Then
            msg = " A fatal error has occured"
            style = MsgBoxStyle.Critical
        Else
            msg = "An error has occured"
            If warning_level = ErrorWarningLevel.WARN Then
                style = MsgBoxStyle.Exclamation
            Else
                style = MsgBoxStyle.Information
            End If
        End If

        If source.Length > 0 Then
            msg += " in " & source & vbCrLf
            log_msg += "Source:  " + source + vbCrLf
        Else
            msg += vbCrLf
        End If

        msg += "Exception Message:  " & ex.Message
        log_msg += "Exception Message:  " & ex.Message & vbCrLf

        _log_handler.writeLog(log_msg, log_level)
        MsgBox(msg, style, title)

    End Sub

    Public Sub DisplayInfo(ByVal title As String, ByVal msg As String)
        MsgBox(msg, MsgBoxStyle.Information, title)
    End Sub

End Class
