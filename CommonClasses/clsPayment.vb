Public Class clsPayment
    Private _id As Integer
    Private _account_id As Integer
    Private _display_text As String
    Private _description As String
    Private _remarks As String
    Private _due_day As Integer
    Private _payment_amount As Double
    Private _dirty As Boolean

    Public Sub New()
        _id = 0
        _account_id = 0
        _display_text = ""
        _description = ""
        _remarks = ""
        _due_day = 1
        _payment_amount = 0.0
    End Sub

    Public Property Id As Integer
        Get
            Return _id
        End Get
        Set(ByVal value As Integer)
            _id = value
        End Set
    End Property

    Public Property AccountId As Integer
        Get
            Return _account_id
        End Get
        Set(value As Integer)
            _account_id = value
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

    Public Property DueDay As Integer
        Get
            Return _due_day
        End Get
        Set(ByVal value As Integer)
            _due_day = value
        End Set
    End Property

    Public Property PaymentAmount As Double
        Get
            Return _payment_amount
        End Get
        Set(ByVal value As Double)
            _payment_amount = value
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

    Public Function Compare(ByRef save As clsPayment) As Boolean

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
            If .AccountId <> _account_id Then
                Return False
            End If
            If .DueDay <> _due_day Then
                Return False
            End If
            If .PaymentAmount <> _payment_amount Then
                Return False
            End If
        End With

        Return True
    End Function

End Class
