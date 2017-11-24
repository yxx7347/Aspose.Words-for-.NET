Namespace ApiExamples.TestData
	Public Class NumericDataSourceWithMethod
		Public Sub New(ByVal value1? As Integer, ByVal value2 As Double, ByVal value3 As Integer, ByVal value4? As Integer, ByVal logical As Boolean)
			Me.Value1 = value1
			Me.Value2 = value2
			Me.Value3 = value3
			Me.Value4 = value4
			Me.Logical = logical
		End Sub

		Public Property Value1() As Integer?

		Public Property Value2() As Double

		Public Property Value3() As Integer

		Public Property Value4() As Integer?

		Public Property Logical() As Boolean

		Public Function Sum(ByVal value1 As Integer, ByVal value2 As Integer) As Integer
			Dim result As Integer = value1 + value2

			Return result
		End Function
	End Class
End Namespace