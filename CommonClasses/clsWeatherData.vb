Public Class clsWeatherData
    Private _id As Integer
    Private _display_text As String
    Private _description As String
    Private _remarks As String
    Private _dirty As Boolean
    Private _datetime As String
    Private _dewpoint As String
    Private _heatindex As String
    Private _indoorhumidity As String
    Private _indoortemperature As String
    Private _outdoorhumidity As String
    Private _outdoortemperature As String
    Private _barometricpressure As String
    Private _rainevent As String
    Private _raineventdatetime As String
    Private _averagewindspeed As String
    Private _windchill As String
    Private _currentwindspeed As String
    Private _currentwinddirection As String
    Private _peakwindspeed As String

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

    Public Property DateTime As String
        Get
            Return _datetime
        End Get
        Set(ByVal value As String)
            _datetime = value
        End Set
    End Property

    Public Property Dewpoint As String
        Get
            Return _dewpoint
        End Get
        Set(ByVal value As String)
            _dewpoint = value
        End Set
    End Property

    Public Property HeatIndex As String
        Get
            Return _heatindex
        End Get
        Set(ByVal value As String)
            _heatindex = value
        End Set
    End Property

    Public Property IndoorHumidity As String
        Get
            Return _indoorhumidity
        End Get
        Set(ByVal value As String)
            _indoorhumidity = value
        End Set
    End Property

    Public Property IndoorTemperature As String
        Get
            Return _indoortemperature
        End Get
        Set(ByVal value As String)
            _indoortemperature = value
        End Set
    End Property

    Public Property OutdoorHumidity As String
        Get
            Return _outdoorhumidity
        End Get
        Set(ByVal value As String)
            _outdoorhumidity = value
        End Set
    End Property

    Public Property OutdoorTemperature As String
        Get
            Return _outdoortemperature
        End Get
        Set(ByVal value As String)
            _outdoortemperature = value
        End Set
    End Property

    Public Property BarometricPressure As String
        Get
            Return _barometricpressure
        End Get
        Set(ByVal value As String)
            _barometricpressure = value
        End Set
    End Property

    Public Property RainEvent As String
        Get
            Return _rainevent
        End Get
        Set(ByVal value As String)
            _rainevent = value
        End Set
    End Property

    Public Property RainEventDateTime As String
        Get
            Return _raineventdatetime
        End Get
        Set(ByVal value As String)
            _raineventdatetime = value
        End Set
    End Property

    Public Property AverageWindspeed As String
        Get
            Return _averagewindspeed
        End Get
        Set(ByVal value As String)
            _averagewindspeed = value
        End Set
    End Property

    Public Property WindChill As String
        Get
            Return _windchill
        End Get
        Set(ByVal value As String)
            _windchill = value
        End Set
    End Property

    Public Property CurrentWindspeed As String
        Get
            Return _currentwindspeed
        End Get
        Set(ByVal value As String)
            _currentwindspeed = value
        End Set
    End Property

    Public Property CurrentWindDirection As String
        Get
            Return _currentwinddirection
        End Get
        Set(ByVal value As String)
            _currentwinddirection = value
        End Set
    End Property

    Public Property PeakWindspeed As String
        Get
            Return _peakwindspeed
        End Get
        Set(ByVal value As String)
            _peakwindspeed = value
        End Set
    End Property

    Public Overrides Function ToString() As String
        Return _display_text
    End Function

    Public Sub New()
        _id = 0
        _display_text = ""
        _description = ""
        _remarks = ""
        _dirty = False
        _datetime = ""
        _dewpoint = ""
        _heatindex = ""
        _indoorhumidity = ""
        _indoortemperature = ""
        _outdoorhumidity = ""
        _outdoortemperature = ""
        _barometricpressure = ""
        _rainevent = ""
        _raineventdatetime = ""
        _averagewindspeed = ""
        _windchill = ""
        _currentwindspeed = ""
        _currentwinddirection = ""
        _peakwindspeed = ""
    End Sub

    Public Function Compare(ByRef save As clsTypeDef) As Boolean

        With save
            If .DisplayText <> _display_text Then
                Return False
            End If
            If .Description <> _description Then
                Return False
            End If
            If .Remarks <> _remarks Then
                Return False
            End If
        End With

        Return True
    End Function

End Class
