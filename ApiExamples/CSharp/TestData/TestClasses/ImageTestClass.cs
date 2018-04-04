using System.IO;

namespace ApiExamples.TestData.TestClasses
{
    public class ImageTestClass
    {
        private readonly Stream mStream;
        private readonly ImageTestClass mImage;
        private readonly byte[] mBytes;
        private readonly string mUri;

        public ImageTestClass(Stream stream)
        {
            this.mStream = stream;
        }

        public ImageTestClass(ImageTestClass imageObject)
        {
            this.mImage = imageObject;
        }

        public ImageTestClass(byte[] imageBytes)
        {
            this.mBytes = imageBytes;
        }

        public ImageTestClass(string uriToImage)
        {
            this.mUri = uriToImage;
        }
    }
}
