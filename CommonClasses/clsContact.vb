Public Class clsContact
    Private _id As Integer
    Private _display_text As String
    Private _description As String
    Private _remarks As String
    Private _contact_type_id As Integer
    Private _first_name As String
    Private _last_name As String
    Private _middle_name As String
    Private _prefix As String
    Private _suffix As String
    Private _dirty As Boolean

    Public Sub New()
        _id = 0
        _display_text = ""
        _description = ""
        _remarks = ""
        _contact_type_id = 0
        _first_name = ""
        _last_name = ""
        _middle_name = ""
        _prefix = ""
        _suffix = ""
        _dirty = False
    End Sub

    Public Property Id As Integer
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

    Public Property Remarks As String
        Get
            Return _remarks
        End Get
        Set(ByVal value As String)
            _remarks = value
        End Set
    End Property

    Public Property ContactTypeId As Integer
        Get
            Return _contact_type_id
        End Get
        Set(value As Integer)
            _contact_type_id = value
        End Set
    End Property

    Public Property FirstName As String
        Get
            Return _first_name
        End Get
        Set(value As String)
            _first_name = value
        End Set
    End Property

    Public Property LastName As String
        Get
            Return _last_name
        End Get
        Set(value As String)
            _last_name = value
        End Set
    End Property

    Public Property MiddleName As String
        Get
            Return _middle_name

        End Get
        Set(value As String)
            _middle_name = value
        End Set
    End Property

    Public Property Prefix As String
        Get
            Return _prefix
        End Get
        Set(value As String)
            _prefix = value
        End Set
    End Property

    Public Property Suffix As String
        Get
            Return _suffix
        End Get
        Set(value As String)
            _suffix = value
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

    Public Function Compare(ByRef save As clsContact) As Boolean
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
            If .FirstName <> _first_name Then
                Return False
            End If
            If .LastName <> _last_name Then
                Return False
            End If
            If .MiddleName <> _middle_name Then
                Return False
            End If
            If .Prefix <> _prefix Then
                Return False
            End If
            If .Suffix <> _suffix Then
                Return False
            End If
        End With

        Return True
    End Function
End Class
