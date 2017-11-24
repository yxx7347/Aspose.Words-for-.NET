Namespace ApiExamples.TestData
	Public Class SimpleDataSource
		Public Sub New(ByVal name As String, ByVal message As String)
			Me.Name = name
			Me.Message = message
		End Sub

		Public Property Name() As String

		Public Property Message() As String
	End Class
End Namespace