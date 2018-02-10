Public Class clsCountry

    Private _id As Integer
    Private _iso As String
    Private _iso3 As String
    Private _iso_numeric As String
    Private _fips As String
    Private _full_name As String
    Private _capital As String
    Private _area As String
    Private _population As String
    Private _continent As String
    Private _internet As String
    Private _currency_code As String
    Private _currency_name As String
    Private _phone_code As String
    Private _postal_code_format As String
    Private _postal_code_regex As String

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Public Property ISO As String
        Get
            Return _iso
        End Get
        Set(ByVal value As String)
            _iso = value
        End Set
    End Property

    Public Property ISO3 As String
        Get
            Return _iso3
        End Get
        Set(ByVal value As String)
            _iso3 = value
        End Set
    End Property

    Public Property ISONumeric As String
        Get
            Return _iso_numeric
        End Get
        Set(ByVal value As String)
            _iso_numeric = value
        End Set
    End Property

    Public Property FIPS As String
        Get
            Return _fips
        End Get
        Set(ByVal value As String)
            _fips = value
        End Set
    End Property

    Public Property FullName As String
        Get
            Return _full_name
        End Get
        Set(ByVal value As String)
            _full_name = value
        End Set
    End Property

    Public Property Capital As String
        Get
            Return _capital
        End Get
        Set(ByVal value As String)
            _capital = value
        End Set
    End Property

    Public Property Area As String
        Get
            Return _area
        End Get
        Set(ByVal value As String)
            _area = value
        End Set
    End Property

    Public Property Population As String
        Get
            Return _population
        End Get
        Set(ByVal value As String)
            _population = value
        End Set
    End Property

    Public Property Continent As String
        Get
            Return _continent
        End Get
        Set(ByVal value As String)
            _continent = value
        End Set
    End Property

    Public Property Internet As String
        Get
            Return _internet
        End Get
        Set(ByVal value As String)
            _internet = value
        End Set
    End Property

    Public Property CurrencyCode As String
        Get
            Return _currency_code
        End Get
        Set(ByVal value As String)
            _currency_code = value
        End Set
    End Property

    Public Property CurrencyName As String
        Get
            Return _currency_name
        End Get
        Set(ByVal value As String)
            _currency_name = value
        End Set
    End Property

    Public Property PhoneCode As String
        Get
            Return _phone_code
        End Get
        Set(ByVal value As String)
            _phone_code = value
        End Set
    End Property

    Public Property PostalCodeFormat As String
        Get
            Return _postal_code_format
        End Get
        Set(ByVal value As String)
            _postal_code_format = value
        End Set
    End Property

    Public Property PostalCodeRegex As String
        Get
            Return _postal_code_regex
        End Get
        Set(ByVal value As String)
            _postal_code_regex = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _full_name
    End Function


    Public Sub New()
        _id = 0
        _iso = ""
        _iso3 = ""
        _iso_numeric = ""
        _fips = ""
        _full_name = ""
        _capital = ""
        _area = ""
        _population = ""
        _continent = ""
        _internet = ""
        _currency_code = ""
        _currency_name = ""
        _phone_code = ""
        _postal_code_format = ""
        _postal_code_regex = ""
    End Sub
End Class
