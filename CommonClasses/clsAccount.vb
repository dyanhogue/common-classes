Public Class clsAccount
    Private _id As Integer
    Private _number As String
    Private _name As String
    Private _display_text As String
    Private _description As String
    Private _remarks As String
    Private _type_id As Integer
    Private _status As String

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Public Property AccountNumber As String
        Get
            Return _number
        End Get
        Set(ByVal value As String)
            _number = value
        End Set
    End Property

    Public Property AccountName As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
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

    Public Property TypeId As Integer
        Get
            Return _type_id
        End Get
        Set(ByVal value As Integer)
            _type_id = value
        End Set
    End Property

    Public Property Status As String
        Get
            Return _status
        End Get
        Set(ByVal value As String)
            _status = value
        End Set
    End Property

    Public Sub New()
        _id = 0
        _number = ""
        _name = ""
        _display_text = ""
        _description = ""
        _remarks = ""
        _type_id = 1
        _status = "A"
    End Sub

    Public Overrides Function ToString() As String
        Return _display_text
    End Function

    Public Function Compare(ByRef save As clsAccount) As Boolean

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
            If .AccountName <> _name Then
                Return False
            End If
            If .AccountNumber <> _number Then
                Return False
            End If
            If .TypeId <> _type_id Then
                Return False
            End If
            If .Status <> _status Then
                Return False
            End If
        End With

        Return True
    End Function
End Class
