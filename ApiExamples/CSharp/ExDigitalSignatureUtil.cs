// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDigitalSignatureUtil : ApiExampleBase
    {
        [Test]
        public void RemoveAllSignaturesEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(Stream, Stream)
            //ExFor:DigitalSignatureUtil.RemoveAllSignatures(String, String)
            //ExSummary:Shows how to remove every signature from a document.
            //By stream:
            Stream docStreamIn = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream docStreamOut = new FileStream(MyDir + @"\Artifacts\Document.NoSignatures.FromStream.doc", FileMode.Create);

            DigitalSignatureUtil.RemoveAllSignatures(docStreamIn, docStreamOut);

            docStreamIn.Close();
            docStreamOut.Close();

            //By string:
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outFileName = MyDir + @"\Artifacts\Document.NoSignatures.FromString.doc";

            DigitalSignatureUtil.RemoveAllSignatures(doc.OriginalFileName, outFileName);
            //ExEnd
        }

        [Test]
        public void LoadSignaturesEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.LoadSignatures(Stream)
            //ExFor:DigitalSignatureUtil.LoadSignatures(String)
            //ExSummary:Shows how to load signatures from a document by stream and by string.
            Stream docStream = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);

            // By stream:
            DigitalSignatureCollection digitalSignatures = DigitalSignatureUtil.LoadSignatures(docStream);
            docStream.Close();

            // By string:
            digitalSignatures = DigitalSignatureUtil.LoadSignatures(MyDir + "Document.DigitalSignature.docx");
            //ExEnd
        }

        [Test]
        public void SignEx()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, String, DateTime)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, String, DateTime)
            //ExSummary:Shows how to sign documents.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By String
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outputDocFileName = MyDir + @"\Artifacts\Document.DigitalSignature.docx";

            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "My comment", DateTime.Now);

            //By Stream
            Stream docInStream = new FileStream(MyDir + "Document.DigitalSignature.docx", FileMode.Open);
            Stream docOutStream = new FileStream(MyDir + @"\Artifacts\Document.DigitalSignature.docx", FileMode.OpenOrCreate);

            DigitalSignatureUtil.Sign(docInStream, docOutStream, ch, "My comment", DateTime.Now);
            //ExEnd

            docInStream.Dispose();
            docOutStream.Dispose();
        }

        [Test]
        [ExpectedException(typeof(IncorrectPasswordException), ExpectedMessage = "The document password is incorrect.")]
        public void IncorrectPasswordForDecrypring()
        {
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //ByDocument
            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));
            string outputDocFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "Comment", DateTime.Now, "docPassword1");
        }

        [Test]
        public void SingDocumentWithPasswordDecrypring()
        {
            //ExStart
            //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, String, DateTime)
            //ExFor:DigitalSignatureUtil.Sign(Stream, Stream, CertificateHolder, String, DateTime)
            //ExSummary:Shows how to sign encrypted documents
            // Create certificate holder from a file.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //ByDocument
            Document doc = new Document(MyDir + "Document.Encrypted.docx", new LoadOptions("docPassword"));
            string outputDocFileName = MyDir + @"\Artifacts\Document.Encrypted.docx";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, ch, "Comment", DateTime.Now, "docPassword");

            // Open encrypted document from a file.
            Document signedDoc = new Document(outputDocFileName, new LoadOptions("docPassword"));

            // Check that encrypted document was successfully signed.
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
            if (signatures.IsValid && (signatures.Count > 0))
            {
                Assert.Pass(); //The document was signed successfully
            }
        }

        [Test]
        public void SingStreamDocumentWithPasswordDecrypring()
        {
            // Create certificate holder from a file.
            CertificateHolder ch = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            //By Stream
            Stream docInStream = new FileStream(MyDir + "Document.Encrypted.docx", FileMode.Open);
            Stream docOutStream = new FileStream(MyDir + @"\Artifacts\Document.Encrypted.docx", FileMode.OpenOrCreate);

            // Digitally sign encrypted with "docPassword" document in the specified path.
            DigitalSignatureUtil.Sign(docInStream, docOutStream, ch, "Comment", DateTime.Now, "docPassword");

            // Open encrypted document from a file.
            Document signedDoc = new Document(docOutStream, new LoadOptions("docPassword"));

            // Check that encrypted document was successfully signed.
            DigitalSignatureCollection signatures = signedDoc.DigitalSignatures;
            if (signatures.IsValid && (signatures.Count > 0))
            {
                docInStream.Dispose();
                docOutStream.Dispose();

                Assert.Pass(); //The document was signed successfully
            }
        }

        [Test]
        public void NoArgumentsForSing()
        {
            Assert.That(() => DigitalSignatureUtil.Sign(String.Empty, String.Empty, null, String.Empty, DateTime.Now, String.Empty), Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void NoCertificateForSign()
        {
            //ByDocument
            Document doc = new Document(MyDir + "Document.DigitalSignature.docx");
            string outputDocFileName = MyDir + @"\Artifacts\Document.DigitalSignature.docx";

            // Digitally sign encrypted with "docPassword" document in the specified path.
            Assert.That(() => DigitalSignatureUtil.Sign(doc.OriginalFileName, outputDocFileName, null, "Comment", DateTime.Now, "docPassword"), Throws.TypeOf<NullReferenceException>());
        }


        #region Sign document
        //ExStart
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
        //ExFor:SignOptions
        //ExFor:SignOptions.SignatureLineId
        //ExFor:SignOptions.SignatureLineImage
        //ExSummary:Shows how to sign a document.

        static readonly DataTable SignPersonsTable = GetSignPersonsTable();
        static readonly DataTable SignDocumentsTable = GetSignDocumentsTable();

        [Test]
        public void FirstSignDocument()
        {
            // Load scanned document.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Get path to document wich we'll be sign
            string pathToSignDocument = MyDir + "SignDocument.doc";

            // Saves just scanned document into a data base. Signs it with document creator and initial signer certificates.
            SignDocument(builder, "Dhocs", pathToSignDocument);

            Document signedDocument = new Document(pathToSignDocument);

            // Write created and signed document into a data base.
            WriteDocumentToDb(signedDocument);

            // In next steps, the document should be signed with approval sides and CEO. It will be achieved the same way as 
            // initial document is created using Aspose.Words API above.
        }

        [Test]
        public void SignDocumentFromDb()
        {
            //Get path to document wich we'll be sign
            string pathToSignDocument = MyDir + "SignDocument.doc";

            // Load signed document from data base.
            Document signedDocument = GetDocumentFromDb(pathToSignDocument);
            DocumentBuilder builder = new DocumentBuilder(signedDocument);

            // Saves just scanned document into a data base. Signs it with document creator and initial signer certificates.
            SignDocument(builder, "Hocs", pathToSignDocument);

            signedDocument = new Document(pathToSignDocument);

            // Write created and signed document into a data base.
            WriteDocumentToDb(signedDocument);

            // In next steps, the document should be signed with approval sides and CEO. It will be achieved the same way as 
            // initial document is created using Aspose.Words API above.
        }

        /// <summary>
        /// Creates new document signed with sign person.
        /// </summary>
        public void SignDocument(DocumentBuilder builder, string signerPerson, string pathToSignDocument)
        {
            SignatureLine signLineDocSigner = builder.InsertSignatureLine(new SignatureLineOptions()).SignatureLine;
            
            signLineDocSigner.Id = GetIdByName(signerPerson);

            CertificateHolder certificateHolderSignerPerson = CertificateHolder.Create(MyDir + "certificate.pfx", "123456");

            SignOptions signOptions = new SignOptions();

            signOptions.SignatureLineId = signLineDocSigner.Id;
            signOptions.SignatureLineImage = GetSignImageFromDb(signerPerson);

            if (builder.Document.OriginalFileName != null)
            {
                DigitalSignatureUtil.Sign(builder.Document.OriginalFileName, pathToSignDocument, certificateHolderSignerPerson, signOptions);
            }
            else
            {
                DigitalSignatureUtil.Sign(pathToSignDocument, pathToSignDocument, certificateHolderSignerPerson, signOptions);
            }
        }

        /// <summary>
        /// Returns image bytes of the scanned signature for the specified user from the data base.
        /// </summary>
        private static byte[] GetSignImageFromDb(string userName)
        {
            // This is just an example.
            // Actually, it will return image bytes from a data base.
            return Encoding.UTF8.GetBytes(userName);
        }

        /// <summary>
        /// Returns identifier of the user from the data base by the specified <paramref name="userName"/>.
        /// </summary>
        private static Guid GetIdByName(string userName)
        {
            // This is just an example.
            // Actually, this method should return value from a data base.
            Guid signerPersonGuid = (from row in SignPersonsTable.AsEnumerable() 
                                      where row.Field<string>("Name") == userName 
                                      select row.Field<Guid>("Id")).FirstOrDefault();

            return signerPersonGuid;
        }

        /// <summary>
        /// Writes document to a data base.
        /// </summary>
        private static void WriteDocumentToDb(Document signDocument)
        {
            // This just an example.
            // Actually, it will save document bytes into a data base.
            MemoryStream stream = new MemoryStream();
            signDocument.Save(stream, SaveFormat.Docx);

            byte[] doc = stream.ToArray();

            DataRow dr = SignDocumentsTable.Select(string.Format("Id = '{0}'", Guid.Parse("2263104B-1083-4988-B571-B356A2C7F51D"))).FirstOrDefault();
            
            if (dr != null)
            {
                dr["Document"] = doc; //changes the document
            }
            else
            {
                SignDocumentsTable.Rows.Add(new object[] { Guid.Parse("2263104B-1083-4988-B571-B356A2C7F51D"), signDocument.OriginalFileName, doc }); 
            }
        }

        /// <summary>
        /// Get document from data base.
        /// </summary>
        private static Document GetDocumentFromDb(string signDocumentName)
        {
            // This just an example.
            // Actually, it will save document bytes into a data base.
            Byte[] signDocument = (from row in SignDocumentsTable.AsEnumerable()
                                   where row.Field<string>("FileName") == signDocumentName
                                     select row.Field<Byte[]>("Document")).FirstOrDefault();

            MemoryStream stream = new MemoryStream(signDocument);
            Document doc = new Document(stream);

            return doc;
        }

        private static DataTable GetSignPersonsTable()
        {
            // Create a test DataTable with two columns and a few rows.
            DataTable table = new DataTable("SignPersons");

            DataColumn column = new DataColumn("Id", typeof(System.Guid));
            column.Unique = true;
            table.Columns.Add(column);

            column = new DataColumn("Name", typeof(System.String));
            table.Columns.Add(column);

            column = new DataColumn("The official", typeof(System.String));
            table.Columns.Add(column);

            column = new DataColumn("Password", typeof(System.String));
            table.Columns.Add(column);

            table.Rows.Add(new object[] { Guid.Parse("CDAA3044-8017-4E07-BFF4-93EA14A3A6C9"), "Hocs", "Head of Corporate Services", "123" });
            table.Rows.Add(new object[] { Guid.Parse("1C22DFF1-B98E-4F65-888F-D55F9A968CD3"), "Dhocs", "Deputy Head of Corporate Services", "1234" });
            table.Rows.Add(new object[] { Guid.Parse("C1DE3C2C-B2F8-4952-96BE-D300FBB9D26B"), "Hofd", "Head of Finance department", "12345" });
            table.Rows.Add(new object[] { Guid.Parse("0A31FD51-46AF-4600-A188-64887881EC47"), "Hoad", "Head of Accounts department", "123456" });
            table.Rows.Add(new object[] { Guid.Parse("409DB46E-4172-4679-ACDE-D78DDB18214F"), "Hdoit", "Head Department of IT", "1234567" });

            return table;
        }

        private static DataTable GetSignDocumentsTable()
        {
            // Create a test DataTable with two columns and a few rows.
            DataTable table = new DataTable("SignDocuments");

            DataColumn column = new DataColumn("Id", typeof(System.Guid));
            column.Unique = true;
            table.Columns.Add(column);

            column = new DataColumn("FileName", typeof(System.String));
            table.Columns.Add(column);

            column = new DataColumn("Document", typeof(System.Byte[]));
            table.Columns.Add(column);

            table.AcceptChanges();

            return table;
        }
        //ExEnd
        #endregion
    }
}