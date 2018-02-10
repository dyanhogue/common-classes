Public Class clsComponent
	Private _id As Integer
	Private _display_text As String
	Private _description As String
	Private _unit_id As Integer
	Private _remarks As String
    Private _dirty As Boolean

	Public Property Id As Integer
		Set(ByVal value As Integer)
			_id = value
		End Set
		Get
			Return _id
		End Get
	End Property

	Public Property UnitId As Integer
		Set(ByVal value As Integer)
			_unit_id = value
		End Set
		Get
			Return _unit_id
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

	Public Property Remarks As String
		Set(ByVal value As String)
			_remarks = value
		End Set
		Get
			Return _remarks
		End Get
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
        _description = ""
        _display_text = ""
        _id = 0
        _remarks = ""
        _unit_id = 0
        _dirty = False
    End Sub

    Public Function Compare(ByVal component As clsComponent) As Boolean
        Dim retval As Boolean = True

        With component
            If Not .Description.Equals(_description) Then
                Return False
            End If

            If Not .DisplayText.Equals(_display_text) Then
                Return False
            End If

            If Not .Remarks.Equals(_remarks) Then
                Return False
            End If

            If .UnitId <> _unit_id Then
                Return False
            End If
        End With

        Return retval
    End Function
End Class
