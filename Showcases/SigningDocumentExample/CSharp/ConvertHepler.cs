using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;

namespace SigningDocumentExample
{
    public class ConvertHepler : Base
    {
        public static byte[] ConverImageToByteArray(string pathToImage)
        {
            Image imageIn = Image.FromFile(pathToImage);

            MemoryStream stream = new MemoryStream();
            imageIn.Save(stream, ImageFormat.Png);

            return stream.ToArray();
        }

        public static Document ConvertByteArrayToDocument(byte[] documentArray)
        {
            MemoryStream stream = new MemoryStream(documentArray);
            Document document = new Document(stream);

            return document;
        }

        public static byte[] ConvertDocumentToByteArray(Document document)
        {
            MemoryStream documentArray = new MemoryStream();
            document.Save(documentArray, SaveFormat.Docx);

            return documentArray.ToArray();
        }
    }
}
