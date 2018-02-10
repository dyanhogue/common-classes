Public Class clsUnit
	Private _id As Integer
	Private _display_text As String
	Private _desc As String
	Private _metric As Boolean
	Private _fractional As Boolean
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
			_desc = value
		End Set
		Get
			Return _desc
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

	Public Property isMetric As Boolean
		Set(ByVal value As Boolean)
			_metric = value
		End Set
		Get
			Return _metric
		End Get
	End Property

	Public Property isFractional As Boolean
		Set(ByVal value As Boolean)
			_fractional = value
		End Set
		Get
			Return _fractional
		End Get
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
		_desc = ""
		_remarks = ""
		_metric = False
		_dirty = False
		_fractional = False
	End Sub

	Public Overrides Function ToString() As String
		Return _display_text
    End Function

    Public Function Compare(ByVal unit As clsUnit) As Boolean
        Dim retval As Boolean = True

        With unit
            If Not .DisplayText.Equals(_display_text) Then
                Return False
            End If

            If Not .Description.Equals(_desc) Then
                Return False
            End If

            If Not .Remarks.Equals(_remarks) Then
                Return False
            End If

            If Not .isMetric.Equals(_metric) Then
                Return False
            End If

            If Not .isFractional.Equals(_fractional) Then
                Return False
            End If
        End With
        Return retval
    End Function
End Class
