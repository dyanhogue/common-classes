Public Class clsFileInfo
	Private _id As Integer
	Private _file_name As String
	Private _file_location As String
	Private _file_type_id As Integer
	Private _description As String
    Private _display_order As Integer
    Private _remarks As String
	Private _dirty As Boolean

	Public Property Id() As Integer
		Get
			Return _id
		End Get
		Set(ByVal value As Integer)
			_id = value
		End Set
	End Property

	Public Property FileName() As String
		Get
			Return _file_name
		End Get
		Set(ByVal value As String)
			_file_name = value
		End Set
	End Property

	Public Property FileLocation() As String
		Get
			Return _file_location
		End Get
		Set(ByVal value As String)
			_file_location = value
		End Set
	End Property

	Public Property FileTypeId() As Integer
		Get
			Return _file_type_id
		End Get
		Set(ByVal value As Integer)
			_file_type_id = value
		End Set
	End Property

	Public Property Description() As String
		Get
			Return _description
		End Get
		Set(ByVal value As String)
			_description = value
		End Set
	End Property

    Public Property DisplayOrder() As Integer
        Get
            Return _display_order
        End Get
        Set(ByVal value As Integer)
            _display_order = value
        End Set
    End Property

    Public Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(ByVal value As String)
            _remarks = value
        End Set
    End Property

	Public Property Dirty() As Boolean
		Get
			Return _dirty
		End Get
		Set(ByVal value As Boolean)
			_dirty = value
		End Set
	End Property

	Public Overrides Function ToString() As String
		Return _file_name
	End Function

	Public Sub New()
		_id = 0
		_file_type_id = 0
        _display_order = 0
        _description = ""
		_file_name = ""
        _file_location = ""
        _remarks = ""
		_dirty = False

	End Sub

    Public Function Compare(ByVal record As clsFileInfo) As Boolean
        Dim retval As Boolean = True

        With record
            If Not .Description.Equals(_description) Then
                Return False
            End If

            If Not .FileName.Equals(_file_name) Then
                Return False
            End If

            If Not .FileLocation.Equals(_file_location) Then
                Return False
            End If

            If .FileTypeId <> _file_type_id Then
                Return False
            End If

            If .DisplayOrder <> _display_order Then
                Return False
            End If

            If Not .Remarks.Equals(_remarks) Then
                Return False
            End If
        End With
        Return retval
    End Function
End Class
