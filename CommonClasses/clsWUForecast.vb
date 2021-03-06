﻿Public Class clsWUForecast
    Private _id As Integer                      '----- <period>
    Private _display_text As String     '----- <title>
    Private _description As String      '----- <fcttest>
    Private _remarks As String
    Private _url_icon As String             '----- <icon_url>
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

    Public Property IconURL As String
        Get
            Return _url_icon
        End Get
        Set(value As String)
            _url_icon = value
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
        _url_icon = ""
        _dirty = False
    End Sub

    Public Function Compare(ByRef save As clsTypeDef) As Boolean

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
        End With

        Return True
    End Function

End Class
