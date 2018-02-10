Public Class clsListItem
    Private _id As String
    Private _display_text As String
    Private _checked As Boolean

    Public Property Id As String
        Get
            Return _id
        End Get
        Set(value As String)
            _id = value
        End Set
    End Property

    Public Property Display As String
        Get
            Return _display_text
        End Get
        Set(value As String)
            _display_text = value
        End Set
    End Property

    Public Property Checked As Boolean
        Get
            Return _checked
        End Get
        Set(ByVal value As Boolean)
            _checked = value
        End Set
    End Property

    Public Sub New()
        _id = "0"
        _display_text = ""
        _checked = False
    End Sub

    Public Overrides Function ToString() As String
        Return _display_text
    End Function
End Class
