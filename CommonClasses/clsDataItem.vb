Public Class clsDataItem
    Private _data_type As DataType
    Private _value As String
    Private _bool_value As Boolean
    Private _int_value As Integer
    Private _double_value As Double
    Private _display_mask As String
    Private _data_length As Integer
    Private _data_precision As Integer


    Public Enum DataType As Integer
        VARCHAR
        NUMBER
        BOOLVAL
        FILEPATH
    End Enum

    Public Property ItemDataType As DataType
        Set(value As DataType)
            _data_type = value
        End Set
        Get
            Return _data_type
        End Get
    End Property

    Public Property ItemValue As String
        Set(value As String)
            _value = value
        End Set
        Get
            Return _value
        End Get
    End Property

    Public Property BoolValue As Boolean
        Set(value As Boolean)
            _bool_value = value
        End Set
        Get
            Return _bool_value
        End Get
    End Property

    Public Property IntValue As Integer
        Set(value As Integer)
            _int_value = value
        End Set
        Get
            Return _int_value
        End Get
    End Property

    Public Property DoubleValue As Double
        Set(value As Double)
            _double_value = value
        End Set
        Get
            Return _double_value
        End Get
    End Property

    Public Property DisplayMask As String
        Set(value As String)
            _display_mask = value
        End Set
        Get
            Return _display_mask
        End Get
    End Property

    Public Property DataLength As Integer
        Set(value As Integer)
            _data_length = value
        End Set
        Get
            Return _data_length
        End Get
    End Property

    Public Property DataPrecision As Integer
        Set(value As Integer)
            _data_precision = value
        End Set
        Get
            Return _data_precision
        End Get
    End Property

    Public Sub New()
        _data_type = DataType.VARCHAR
        _value = ""
        _bool_value = True
        _int_value = 0
        _double_value = 0
        _display_mask = ""
        _data_length = 0
        _data_precision = 0
    End Sub

    Public Overrides Function ToString() As String
        If Len(_display_mask) > 0 Then
            Return Format(_value, _display_mask)
        Else
            Return _value
        End If
    End Function

End Class
