Public Class clsPhone
	Private _id As Integer
    Private _display_text As String
    Private _description As String
    Private _type_id As Integer
    Private _area_code As String
	Private _prefix As String
	Private _number As String
	Private _country_code As String
	Private _ext As String
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

    Public Property DisplayText() As String
        Get
            Return _display_text
        End Get
        Set(ByVal value As String)
            _display_text = value
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

    Public Property TypeId() As Integer
        Get
            Return _type_id
        End Get
        Set(ByVal value As Integer)
            _type_id = value
        End Set
    End Property

    Public Property AreaCode() As String
		Get
			Return _area_code
		End Get
		Set(ByVal value As String)
			_area_code = value
		End Set
	End Property

	Public Property Prefix() As String
		Get
			Return _prefix
		End Get
		Set(ByVal value As String)
			_prefix = value
		End Set
	End Property

	Public Property PhoneNumber() As String
		Get
			Return _number
		End Get
		Set(ByVal value As String)
			_number = value
		End Set
	End Property

	Public Property CountryCode() As String
		Get
			Return _country_code
		End Get
		Set(ByVal value As String)
			_country_code = value
		End Set
	End Property

	Public Property Ext() As String
		Get
			Return _ext
		End Get
		Set(ByVal value As String)
			_ext = value
		End Set
	End Property

	Public Property Remarks() As String
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
		Return "(" & _area_code & ") " & _prefix & "-" & _number
	End Function

	Public Sub New()
		_id = 0
        _display_text = ""
        _description = ""
        _type_id = 0
        _area_code = ""
		_prefix = ""
		_number = ""
		_country_code = ""
		_ext = ""
		_remarks = ""
		_dirty = False
	End Sub

	Public Function Compare(ByRef save As clsPhone) As Boolean

        With save
            If .TypeId <> _type_id Then
                Return False
            End If
            If .AreaCode <> _area_code Then
                Return False
            End If
            If .Prefix <> _prefix Then
                Return False
            End If
            If .PhoneNumber <> _number Then
                Return False
            End If
            If .CountryCode <> _country_code Then
                Return False
            End If
            If .Ext <> _ext Then
                Return False
            End If
            If .Remarks <> _remarks Then
                Return False
            End If
        End With

        Return True
    End Function
End Class
