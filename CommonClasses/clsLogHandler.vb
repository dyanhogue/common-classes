Public Class clsLogHandler
	Private _path As String
	Private _filename As String
    Private _status_label As Object
    Private _status_text As Object
	Private _display_type As String

	Public WriteOnly Property LogPath() As String
		Set(ByVal value As String)
			_path = value
		End Set
	End Property

	Public WriteOnly Property LogFilename() As String
		Set(ByVal value As String)
			_filename = value
		End Set
	End Property

    Public WriteOnly Property StatusLabel() As Object
        Set(ByVal value As Object)
            _status_label = value
        End Set
    End Property

    Public WriteOnly Property StatusText() As Object
        Set(ByVal value As Object)
            _status_text = value
        End Set
    End Property

	Public WriteOnly Property DisplayType() As String
		Set(ByVal value As String)
			_display_type = value
		End Set
	End Property

	Public Sub writeLog(ByVal msg As String, ByVal level As String)
		Dim lineout As String
		lineout = level & ":  " & msg & vbCrLf
		FileIO.FileSystem.WriteAllText(_path & _filename, lineout, True)
	End Sub

    Public Sub writeLog(ByRef ex As Exception, ByVal source As String, ByVal msg As String, ByVal level As String)
        Dim lineout As String
        lineout = "Error source:  " + source + vbCrLf
        lineout += "     message:  " + msg + vbCrLf
        lineout += "     level:  " + level + vbCrLf
        lineout += "     detail:  " + ex.StackTrace + vbCrLf
        lineout += "End error" + vbCrLf

        FileIO.FileSystem.WriteAllText(_path & _filename, lineout, True)

    End Sub

    Public Sub writeStatus(ByVal msg As String, Optional ByVal append As Boolean = False)
		Select Case _display_type
			Case "text"
				If append Then
					_status_text.AppendText(msg & vbCrLf)
				Else
					_status_text.Text = msg
				End If
			Case "label"
				_status_label.Text = msg
		End Select
	End Sub

	Public Sub New()
		_path = "\log\"
		_filename = CStr(Now.Month) & CStr(Now.Day) & CStr(Now.Year) & "_log.log"
		_display_type = "label"
        _status_label = New Object
        _status_text = New Object

	End Sub
End Class
