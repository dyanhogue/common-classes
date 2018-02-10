Public Class clsLedgerItem
    Private _id As Integer
    Private _display_text As String
    Private _description As String
    Private _amount As Double
    Private _category_id As Integer
    Private _sub_category_id As Integer
    Private _remarks As String
    Private _is_deposit As Boolean
    Private _balance As Double
    Private _date_created As String
    Private _dirty As Boolean

    Public Property Id As Integer
        Set(ByVal value As Integer)
            _id = value
        End Set
        Get
            Return _id
        End Get
    End Property

    Public Property DisplayText As String
        Set(ByVal value As String)
            _display_text = value
        End Set
        Get
            Return _display_text
        End Get
    End Property

    Public Property Description As String
        Set(ByVal value As String)
            _description = value
        End Set
        Get
            Return _description
        End Get
    End Property

    Public Property Amount As Double
        Set(value As Double)
            _amount = value
        End Set
        Get
            Return _amount
        End Get
    End Property

    Public Property CategoryId As Integer
        Set(value As Integer)
            _category_id = value
        End Set
        Get
            Return _category_id
        End Get
    End Property

    Public Property SubCategoryId As Integer
        Set(value As Integer)
            _sub_category_id = value
        End Set
        Get
            Return _sub_category_id
        End Get
    End Property

    Public Property Remarks As String
        Set(ByVal value As String)
            _remarks = value
        End Set
        Get
            Return _remarks
        End Get
    End Property

    Public Property IsDeposit As Boolean
        Set(ByVal value As Boolean)
            _is_deposit = value
        End Set
        Get
            Return _is_deposit
        End Get
    End Property

    Public Property Balance As Double
        Set(ByVal value As Double)
            _balance = value
        End Set
        Get
            Return _balance
        End Get
    End Property

    Public Property DateCreated As String
        Get
            Return _date_created
        End Get
        Set(ByVal value As String)
            _date_created = value
        End Set
    End Property

    Public Property Dirty As Boolean
        Set(ByVal value As Boolean)
            _dirty = value
        End Set
        Get
            Return _dirty
        End Get
    End Property

    Public Sub New()
        _id = 0
        _display_text = ""
        _description = ""
        _amount = 0
        _category_id = 0
        _sub_category_id = 0
        _remarks = ""
        _is_deposit = False
        _balance = 0
        _date_created = Now.ToShortDateString
        _dirty = False
    End Sub

    Public Overrides Function ToString() As String
        Return _display_text + " $" + _amount
    End Function

    Public Function Comapre(ByVal record As clsLedgerItem) As Boolean
        Dim retval As Boolean = True

        With record
            If Not _display_text.Equals(.DisplayText) Then
                Return False
            End If

            If Not _description.Equals(.Description) Then
                Return False
            End If

            If Not _amount = .Amount Then
                Return False
            End If

            If Not _category_id = .CategoryId Then
                Return False
            End If

            If Not _sub_category_id = .SubCategoryId Then
                Return False
            End If

            If Not _remarks.Equals(.Remarks) Then
                Return False
            End If

            If Not _is_deposit = .IsDeposit Then
                Return False
            End If
        End With

        Return retval
    End Function
End Class
