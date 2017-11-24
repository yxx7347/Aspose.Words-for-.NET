Imports System.IO
Imports Aspose.Words

Namespace ApiExamples.TestData
	Public Class DocumentDataSource
		Public Sub New(ByVal doc As Document)
			Me.Document = doc
		End Sub

		Public Sub New(ByVal stream As Stream)
			Me.DocumentByStream = stream
		End Sub

		Public Sub New(ByVal byteDoc() As Byte)
			Me.DocumentByByte = byteDoc
		End Sub

		Public Sub New(ByVal uriToDoc As String)
			Me.DocumentByUri = uriToDoc
		End Sub

		Public Property Document() As Document

		Public Property DocumentByStream() As Stream

		Public Property DocumentByByte() As Byte()

		Public Property DocumentByUri() As String
	End Class
End Namespace