Public Class clsEmailAddress
	Private _id As Integer
	Private _address As String
	Private _display_text As String
	Private _type_id As Integer
	Private _remarks As String
	Private _dirty As Boolean

	Public Property Id() As Integer
		Get
			Return _id
		End Get
		Set(ByVal value As Integer)
			_id = value
		End Set
	End Property

	Public Property EmailAddress() As String
		Get
			Return _address
		End Get
		Set(ByVal value As String)
			_address = value
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

	Public Property TypeId() As Integer
		Get
			Return _type_id
		End Get
		Set(ByVal value As Integer)
			_type_id = value
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

	Public Overrides Function ToString() As String
		Return _display_text
	End Function

	Public Sub New()
		_id = 0
		_type_id = 0
		_address = ""
		_display_text = ""
		_remarks = ""
		_dirty = False
	End Sub

    Public Function Compare(ByVal address As clsEmailAddress) As Boolean
        Dim retval As Boolean = True

        With address
            If Not .EmailAddress.Equals(_address) Then
                Return False
            End If

            If Not .DisplayText.Equals(_display_text) Then
                Return False
            End If

            If Not .Remarks.Equals(_remarks) Then
                Return False
            End If

            If .TypeId <> _type_id Then
                Return False
            End If
        End With
        Return retval
    End Function
End Class
