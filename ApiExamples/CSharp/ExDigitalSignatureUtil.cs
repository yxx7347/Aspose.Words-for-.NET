// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
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

        //ExStart
        //ExFor:DigitalSignatureUtil.Sign(String, String, CertificateHolder, SignOptions)
        //ExFor:SignOptions
        //ExFor:SignOptions.SignatureLineId
        //ExFor:SignOptions.SignatureLineImage
        //ExSummary:Shows how to implement document signing by any person that you need. Here the initial stage of the stated above example 
        //in which the document is created and it is sent to the database with signatures of approving divisions.

        //Preparation test table with persons of approving divisions.
        static readonly DataTable SignPersonsTable = GetSignPersonsTable();

        //Preparation test table for saving signed documents.
        static readonly DataTable SignDocumentsTable = GetSignDocumentsTable();

        /// <summary>
        /// Creates new document, signed by any persons from responsible divisions.
        /// </summary>
        [Test]
        public void FirstSigningDocumentAndInsertInDb()
        {
            //Load scanned document.
            Document doc = new Document(MyDir + "Document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Set path to the document that will be signed.
            string pathToSignedDocument = MyDir + "SignDocument.doc";

            //Let it be signed by the 'Deputy Head of Corporate Services'.
            SignDocument(builder, "Dhocs", pathToSignedDocument);

            //Get a signed document.
            Document signedDocument = new Document(pathToSignedDocument);

            //Write signed document into a database.
            WriteDocumentToDb(signedDocument);

            // In next steps, the document should be signed with approval sides and CEO. It will be achieved the same way as 
            // initial document is created using Aspose.Words API above. See SignDocumentFromDbAndUpdate method for this.
        }

        /// <summary>
        /// Signing document that are in database and update it.
        /// </summary>
        [Test]
        public void SigningDocumentFromDbAndUpdate()
        {
            //Set path to the document that will be signed.
            string pathToSignedDocument = MyDir + "SignDocument.doc";

            //Load signed document from a data base.
            Document signedDocument = GetDocumentFromDb(pathToSignedDocument);
            DocumentBuilder builder = new DocumentBuilder(signedDocument);

            //Let it be signed by the 'Head of Corporate Services'.
            SignDocument(builder, "Hocs", pathToSignedDocument);

            //Get a signed document.
            signedDocument = new Document(pathToSignedDocument);

            //Write signed document into a database.
            WriteDocumentToDb(signedDocument);
        }

        /// <summary>
        /// Creates new document signed with responsible person.
        /// </summary>
        private static void SignDocument(DocumentBuilder builder, string signerPerson, string pathToSignedDocument)
        {
            SignatureLineOptions options = new SignatureLineOptions();
            options.Signer = signerPerson;
            options.SignerTitle = GetPositionByName(signerPerson);
            options.ShowDate = true;
            options.DefaultInstructions = false;
            options.Instructions = "This is the test signature line";
            options.AllowComments = true;

            //Add signature line for responsible person.
            SignatureLine signatureLine = builder.InsertSignatureLine(options).SignatureLine;
            signatureLine.Id = GetIdByName(signerPerson);

            builder.Document.Save(pathToSignedDocument);

            //Sign document with responsible person certificate and image of its scanned signature.
            CertificateHolder certificateHolderSignerPerson = GetCertificateHolder(MyDir + "morzal.pfx", "aw");

            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = signatureLine.Id;
            signOptions.SignatureLineImage = GetSignImageFromDb(signerPerson);

            DigitalSignatureUtil.Sign(builder.Document.OriginalFileName ?? pathToSignedDocument, pathToSignedDocument, certificateHolderSignerPerson, signOptions);
        }

        /// <summary>
        /// Creates certificate holder with responsible person certificate.
        /// </summary>
        private static CertificateHolder GetCertificateHolder(string pathToCertificate, string certificatePassword)
        {
            return CertificateHolder.Create(pathToCertificate, certificatePassword);
        }

        /// <summary>
        /// Returns image bytes of the scanned signature for the specified user from the database by the specified <paramref name="userName"/>.
        /// </summary>
        private static byte[] GetSignImageFromDb(string userName)
        {
            //Get signed document from a database by 'FileName'.
            Byte[] image = (from row in SignDocumentsTable.AsEnumerable()
                                     where row.Field<string>("Name") == userName
                                     select row.Field<Byte[]>("Image")).FirstOrDefault();

            return image;
        }

        /// <summary>
        /// Get image bytes for saving into a database.
        /// </summary>
        private static byte[] imageToByteArray(Image imageIn)
        {
            MemoryStream ms = new MemoryStream();
            imageIn.Save(ms, ImageFormat.Png);

            return ms.ToArray();
        }

        /// <summary>
        /// Returns identifier of the user from the database by the specified <paramref name="userName"/>.
        /// </summary>
        private static Guid GetIdByName(string userName)
        {
            Guid signerPersonGuid = (from row in SignPersonsTable.AsEnumerable() 
                                      where row.Field<string>("Name") == userName 
                                      select row.Field<Guid>("Id")).FirstOrDefault();

            return signerPersonGuid;
        }

        /// <summary>
        /// Returns position of the user from the database by the specified <paramref name="userName"/>.
        /// </summary>
        private static string GetPositionByName(string userName)
        {
            string signerPersonPosition = (from row in SignPersonsTable.AsEnumerable()
                                     where row.Field<string>("Name") == userName
                                     select row.Field<string>("Position")).FirstOrDefault();

            return signerPersonPosition;
        }

        /// <summary>
        /// Writes signed document into a database.
        /// </summary>
        private static void WriteDocumentToDb(Document signedDocument)
        {
            //Create stream from signed document.
            MemoryStream stream = new MemoryStream();
            signedDocument.Save(stream, SaveFormat.Docx);

            byte[] doc = stream.ToArray();

            DataRow dbSignedDocument = SignDocumentsTable.Select(string.Format("Id = '{0}'", Guid.Parse("2263104B-1083-4988-B571-B356A2C7F51D"))).FirstOrDefault();
            
            //If this signed document are already exists, then we just update this document on new, else create new row in table.
            if (dbSignedDocument != null)
            {
                dbSignedDocument["Document"] = doc; //Changes the document.
            }
            else
            {
                SignDocumentsTable.Rows.Add(new object[] { Guid.Parse("2263104B-1083-4988-B571-B356A2C7F51D"), signedDocument.OriginalFileName, doc }); //Add new row.
            }
        }

        /// <summary>
        /// Get signed document from a database.
        /// </summary>
        private static Document GetDocumentFromDb(string signedDocumentName)
        {
            //Get signed document from a database by 'FileName'.
            Byte[] signedDocument = (from row in SignDocumentsTable.AsEnumerable()
                                     where row.Field<string>("FileName") == signedDocumentName
                                     select row.Field<Byte[]>("Document")).FirstOrDefault();

            //Generate new document from a stream.
            MemoryStream stream = new MemoryStream(signedDocument);
            Document signedDocumentFromDb = new Document(stream);

            return signedDocumentFromDb;
        }

        /// <summary>
        /// Creates new table with responsible persons.
        /// </summary>
        private static DataTable GetSignPersonsTable()
        {
            DataTable table = new DataTable("SignPersons");

            DataColumn column = new DataColumn("Id", typeof(System.Guid));
            column.Unique = true;
            table.Columns.Add(column);

            column = new DataColumn("Name", typeof(System.String));
            table.Columns.Add(column);

            column = new DataColumn("Position", typeof(System.String));
            table.Columns.Add(column);

            column = new DataColumn("Image", typeof(System.Byte[]));
            table.Columns.Add(column);

            Image img = Image.FromFile(MyDir + @"\Images\LogoSmall.png");
            
            table.Rows.Add(new object[] { Guid.Parse("CDAA3044-8017-4E07-BFF4-93EA14A3A6C9"), "Hocs", "Head of Corporate Services", imageToByteArray(img) });
            table.Rows.Add(new object[] { Guid.Parse("1C22DFF1-B98E-4F65-888F-D55F9A968CD3"), "Dhocs", "Deputy Head of Corporate Services", imageToByteArray(img) });
            table.Rows.Add(new object[] { Guid.Parse("C1DE3C2C-B2F8-4952-96BE-D300FBB9D26B"), "Hofd", "Head of Finance department", imageToByteArray(img) });
            table.Rows.Add(new object[] { Guid.Parse("0A31FD51-46AF-4600-A188-64887881EC47"), "Hoad", "Head of Accounts department", imageToByteArray(img) });
            table.Rows.Add(new object[] { Guid.Parse("409DB46E-4172-4679-ACDE-D78DDB18214F"), "Hdoit", "Head Department of IT", imageToByteArray(img) });

            return table;
        }

        /// <summary>
        /// Creates table for saving signed documents.
        /// </summary>
        private static DataTable GetSignDocumentsTable()
        {
            DataTable table = new DataTable("SignDocuments");

            DataColumn column = new DataColumn("Id", typeof(System.Guid));
            column.Unique = true;
            table.Columns.Add(column);

            column = new DataColumn("FileName", typeof(System.String));
            table.Columns.Add(column);

            column = new DataColumn("Document", typeof(System.Byte[]));
            table.Columns.Add(column);

            return table;
        }
        //ExEnd
    }
}