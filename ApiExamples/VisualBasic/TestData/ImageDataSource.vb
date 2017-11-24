Imports System.Drawing
Imports System.IO

Namespace ApiExamples.TestData
	Public Class ImageDataSource
		Public Sub New(ByVal stream As Stream)
			Me.Stream = stream
		End Sub

		Public Sub New(ByVal imageObject As Image)
			Me.Image = imageObject
		End Sub

		Public Sub New(ByVal imageBytes() As Byte)
			Me.Bytes = imageBytes
		End Sub

		Public Sub New(ByVal uriToImage As String)
			Me.Uri = uriToImage
		End Sub

		Public Property Stream() As Stream

		Public Property Image() As Image

		Public Property Bytes() As Byte()

		Public Property Uri() As String
	End Class
End Namespace