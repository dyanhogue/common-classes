Public Class clsCyberdress
    Private _id As Integer
    Private _display_text As String
    Private _description As String
    Private _remarks As String
    Private _type_id As Integer
    Private _value As String
    Private _uid As String
    Private _pwd As String
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

    Public Property Remarks() As String
        Get
            Return _remarks
        End Get
        Set(ByVal value As String)
            _remarks = value
        End Set
    End Property

    Public Property TypeId As Integer
        Get
            Return _type_id
        End Get
        Set(value As Integer)
            _type_id = value
        End Set
    End Property

    Public Property Value As String
        Get
            Return _value
        End Get
        Set(value As String)
            _value = value
        End Set
    End Property

    Public Property UID As String
        Get
            Return _uid
        End Get
        Set(value As String)
            _uid = value
        End Set
    End Property

    Public Property Password As String
        Get
            Return _pwd
        End Get
        Set(value As String)
            _pwd = value
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
        Return _display_text
    End Function

    Public Sub New()
        _id = 0
        _display_text = ""
        _description = ""
        _remarks = ""
        _type_id = 0
        _value = ""
        _uid = ""
        _pwd = ""
        _dirty = False
    End Sub

    Public Function Compare(ByRef save As clsCyberdress) As Boolean

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
            If .TypeId <> _type_id Then
                Return False
            End If
            If .Value <> _value Then
                Return False
            End If
            If .UID <> _uid Then
                Return False
            End If
            If .Password <> _pwd Then
                Return False
            End If
        End With

        Return True
    End Function

End Class
