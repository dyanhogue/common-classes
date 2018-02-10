Public Class clsAddress
    Private _id As Integer
    Private _display_text As String
    Private _description As String
	Private _address_type_id As Integer
	Private _line1 As String
    Private _line2 As String
    Private _line3 As String
    Private _line4 As String
    Private _city As String
    Private _state As String
    Private _state_code As String
    Private _postal_code As String
    Private _country As String
    Private _country_code As String
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

    Public Property DisplayText As String
        Get
            Return _display_text
        End Get
        Set(ByVal value As String)
            _display_text = value
        End Set
    End Property

    Public Property Description As String
        Get
            Return _description
        End Get
        Set(ByVal value As String)
            _description = value
        End Set
    End Property

	Public Property AddressTypeId() As Integer
		Get
			Return _address_type_id
		End Get
		Set(ByVal value As Integer)
			_address_type_id = value
		End Set
	End Property

	Public Property Line1() As String
		Get
			Return _line1
		End Get
		Set(ByVal value As String)
			_line1 = value
		End Set
	End Property

	Public Property Line2() As String
		Get
			Return _line2
		End Get
		Set(ByVal value As String)
			_line2 = value
		End Set
	End Property

    Public Property Line3 As String
        Get
            Return _line3
        End Get
        Set(value As String)
            _line3 = value
        End Set
    End Property

    Public Property Line4 As String
        Get
            Return _line4
        End Get
        Set(value As String)
            _line4 = value
        End Set
    End Property

    Public Property City() As String
		Get
			Return _city
		End Get
		Set(ByVal value As String)
			_city = value
		End Set
	End Property

	Public Property State() As String
		Get
			Return _state
		End Get
		Set(ByVal value As String)
			_state = value
		End Set
	End Property

    Public Property StateCode As String
        Get
            Return _state_code
        End Get
        Set(value As String)
            _state_code = value
        End Set
    End Property

    Public Property PostalCode() As String
		Get
			Return _postal_code
		End Get
		Set(ByVal value As String)
			_postal_code = value
		End Set
	End Property

	Public Property Country() As String
		Get
			Return _country
		End Get
		Set(ByVal value As String)
			_country = value
		End Set
	End Property

    Public Property CountryCode As String
        Get
            Return _country_code
        End Get
        Set(value As String)
            _country_code = value
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

    Public Function getLabel(ByVal bcountry As Boolean) As String
        Dim retval As String = ""

        retval = _line1 & vbCrLf
        If Line2.Length > 0 Then
            retval += _line2 & vbCrLf
        End If
        retval += _city & "  " & _state & "  " & _postal_code & vbCrLf
        If bcountry Then
            retval += _country
        End If

        Return retval
    End Function

    Public Overrides Function ToString() As String
        Return _display_text
    End Function

	Public Sub New()
        _id = 0
        _display_text = ""
        _description = ""
		_address_type_id = 0
		_line1 = ""
		_line2 = ""
		_city = ""
		_state = ""
		_postal_code = ""
		_country = ""
        _remarks = ""
        _dirty = False
	End Sub

	Public Function Compare(ByRef save As clsAddress) As Boolean
        Dim retval As Boolean = True
        With save
            If .AddressTypeId <> _address_type_id Then
                retval = False
            End If
            If Not .Line1.Equals(_line1) Then
                retval = False
            End If
            If Not .Line2.Equals(_line2) Then
                retval = False
            End If
            If Not City.Equals(_city) Then
                retval = False
            End If
            If Not State.Equals(_state) Then
                retval = False
            End If
            If Not .PostalCode.Equals(_postal_code) Then
                retval = False
            End If
            If Not .Country.Equals(_country) Then
                retval = False
            End If
            If Not .Remarks.Equals(_remarks) Then
                retval = False
            End If
        End With
		Return retval
	End Function
End Class
