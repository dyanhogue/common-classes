Public Class clsState
    Private _id As Integer
    Private _name As String
    Private _abbreviation As String
    Private _country As String
    Private _type As String
    Private _sort As Integer
    Private _status As String
    Private _occupied As String
    Private _notes As String
    Private _fips_state As String
    Private _assoc_press As String
    Private _standard_federal_region As String
    Private _census_region As String
    Private _census_region_name As String
    Private _census_division As String
    Private _census_division_name As String
    Private _circuit_court As String

    Public Sub New()
        _id = 0
        _name = ""
        _abbreviation = ""
        _country = ""
        _type = ""
        _sort = 0
        _status = ""
        _occupied = ""
        _notes = ""
        _fips_state = ""
        _assoc_press = ""
        _standard_federal_region = ""
        _census_region = ""
        _census_region_name = ""
        _census_division = ""
        _census_division_name = ""
        _circuit_court = ""

    End Sub

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Public Property StateName As String
        Get
            Return _name
        End Get
        Set
            _name = Value
        End Set
    End Property

    Public Property StateCode As String
        Get
            Return _abbreviation
        End Get
        Set
            _abbreviation = Value
        End Set
    End Property

    Public Property Country As String
        Get
            Return _country
        End Get
        Set
            _country = Value
        End Set
    End Property

    Public Property StateType As String
        Get
            Return _type
        End Get
        Set
            _type = value
        End Set
    End Property


    Public Property Sort As Integer
        Get
            Return _sort
        End Get
        Set(ByVal value As Integer)
            _sort = value
        End Set
    End Property

    Public Property Status As String
        Get
            Return _status
        End Get
        Set
            _status = Value
        End Set
    End Property

    Public Property Occupied As String
        Get
            Return _occupied
        End Get
        Set
            _occupied = Value
        End Set
    End Property

    Public Property Notes As String
        Get
            Return _notes
        End Get
        Set
            _notes = Value
        End Set
    End Property

    Public Property FipsState As String
        Get
            Return _fips_state
        End Get
        Set
            _fips_state = Value
        End Set
    End Property

    Public Property AssocPress As String
        Get
            Return _assoc_press
        End Get
        Set
            _assoc_press = Value
        End Set
    End Property

    Public Property StandardFedRegion As String
        Get
            Return _standard_federal_region
        End Get
        Set
            _standard_federal_region = Value
        End Set
    End Property

    Public Property CensusRegion As String
        Get
            Return _census_region
        End Get
        Set
            _census_region = Value
        End Set
    End Property

    Public Property CensusRegionName As String
        Get
            Return _census_region_name
        End Get
        Set
            _census_region_name = Value
        End Set
    End Property

    Public Property CensusDivision As String
        Get
            Return _census_division
        End Get
        Set
            _census_division = Value
        End Set
    End Property

    Public Property CensusDivisionName As String
        Get
            Return _census_division_name
        End Get
        Set
            _census_division_name = Value
        End Set
    End Property

    Public Property CircuitCourt As String
        Get
            Return _circuit_court
        End Get
        Set
            _circuit_court = Value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _abbreviation + " - " + _name
    End Function
End Class
