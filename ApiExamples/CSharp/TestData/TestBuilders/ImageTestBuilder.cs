using System;
using System.Drawing;
using System.IO;
using ApiExamples.TestData.TestClasses;

namespace ApiExamples.TestData.TestBuilders
{
    public class ImageTestBuilder : ApiExampleBase
    {
        private Image mImage;
        private Stream mImageStream;
        private byte[] mImageBytes;
        private string mImageUri;

        public ImageTestBuilder()
        {
            this.mImage = Image.FromFile(ImageDir + "Watermark.png");
            this.mImageStream = Stream.Null;
            this.mImageBytes = new byte[0];
            this.mImageUri = string.Empty;
        }

        public ImageTestBuilder WithImage(Image image)
        {
            this.mImage = image;
            return this;
        }

        public ImageTestBuilder WithImageStream(Stream imageStream)
        {
            this.mImageStream = imageStream;
            return this;
        }

        public ImageTestBuilder WithImageBytes(byte[] imageBytes)
        {
            this.mImageBytes = imageBytes;
            return this;
        }

        public ImageTestBuilder WithImageUri(string imageUri)
        {
            this.mImageUri = imageUri;
            return this;
        }

        public ImageTestClass Build()
        {
            return new ImageTestClass(mImage, mImageStream, mImageBytes, mImageUri);
        }
    }
}
