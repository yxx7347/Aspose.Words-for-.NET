Imports System

Namespace ApiExamples.TestData
	Friend Module DataTables
		Friend Function AddClientsTestData() As DataSet
			Dim rnd As New Random()
			Dim ds As New DataSet()

			Dim age As Integer = 30
			Dim j As Integer = 1

			For i As Integer = 1 To 10
				ds.Clients.Rows.Add(i, "Name " & i)
			Next i

			For i As Integer = 1 To 3
				Dim y As Integer = j
				Do While y <= j + 2
					ds.Contracts.Rows.Add(y, y, i, 1000)
					y += 1
				Loop

				j = j + 3
			Next i

			For i As Integer = 1 To 3
			   ds.Managers.Rows.Add(i, "Name " & i, age)
			   age = age + 1
			Next i

			Return ds
		End Function
	End Module
End Namespace