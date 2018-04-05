using System.Drawing;
using System.IO;

namespace ApiExamples.NetCore.TestData.TestClasses
{
    public class ImageTestClass
    {
        public Image Image { get; set; }
        public Stream ImageStream { get; set; }
        public byte[] ImageBytes { get; set; }
        public string ImageUri { get; set; }

        public ImageTestClass(Image image, Stream imageStream, byte[] imageBytes, string imageUri)
        {
            this.Image = image;
            this.ImageStream = imageStream;
            this.ImageBytes = imageBytes;
            this.ImageUri = imageUri;
        }
    }
}
