using System;
using Aspose.Words;

namespace SigningDocumentExample
{
    public class SignPerson
    {
        public Guid PersonId { get; set; }
        public string Name { get; set; }
        public string Position { get; set; }
        public SignDocument[] Documents { get; set; }
        public byte[] Image { get; set; }
    }
}
