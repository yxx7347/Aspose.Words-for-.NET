Imports System

Namespace ApiExamples.TestData
	Public Class NumericDataSource
		Public Sub New(ByVal value1 As Integer, ByVal value2 As Double, ByVal value3 As Integer, ByVal [date] As Date)
			Me.Value1 = value1
			Me.Value2 = value2
			Me.Value3 = value3
			Me.Date = [date]
		End Sub

		Public Property Value1() As Integer

		Public Property Value2() As Double

		Public Property Value3() As Integer

		Public Property [Date]() As Date
	End Class
End Namespace