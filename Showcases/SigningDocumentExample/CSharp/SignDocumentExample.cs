// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace SigningDocumentExample
{
    [TestFixture]
    public class SignDocumentExample : Base
    {
        [Test]
        public void Main()
        {
            // Saves just scanned document into a data base. Signs it with document creator and initial signer certificates.
            SignDocument("SignPerson 1", "Jane Doe_02_08_2017_N1.docx");
        }

        private void SignDocument(string signerPerson, string pathToDocument)
        {
            CreateTestData();

            Document baseDocument;
            DocumentBuilder builder;

            if (!CheckDocumentExist(pathToDocument))
            {
                //Load scanned document.
                baseDocument = new Document(pathToDocument);
                builder = new DocumentBuilder(baseDocument);
            }
            else
            {
                baseDocument = GetDocument(pathToDocument);
                builder = new DocumentBuilder(baseDocument);
            }

            //Add info about responsible person
            SignatureLineOptions signatureLineOptions = new SignatureLineOptions();
            signatureLineOptions.Signer = GetSignPersonByName(signerPerson).Name;
            signatureLineOptions.SignerTitle = GetSignPersonByName(signerPerson).Position;

            //Add signature line for responsible person.
            SignatureLine signatureLine = builder.InsertSignatureLine(signatureLineOptions).SignatureLine;
            signatureLine.Id = GetSignPersonByName(signerPerson).PersonId;

            //Save document with line signatures into temporary stream for future signing.
            MemoryStream stream = new MemoryStream();
            baseDocument.Save(stream, SaveFormat.Docx);

            CertificateHolder certificateHolderSignerPerson = CertificateHolder.Create(MyDir + "morzal.pfx", "aw");
            
            SignOptions signOptions = new SignOptions();
            signOptions.SignatureLineId = GetSignPersonByName(signerPerson).PersonId;
            signOptions.SignatureLineImage = GetSignPersonByName(signerPerson).Image;

            DigitalSignatureUtil.Sign(stream, stream, certificateHolderSignerPerson, signOptions);

            WriteDocumentToDb(stream, signerPerson);
        }

        public void WriteDocumentToDb(MemoryStream stream, string signerPerson)
        {
            // This an example.
            // Actually, it will save or update object with data base.
            Document signedDocument = new Document(stream);

            if (!CheckDocumentExist(signedDocument.OriginalFileName))
            {
                GetSignPersonByName(signerPerson).Documents = new[] { new SignDocument { DocumentId = Guid.NewGuid(), DocumentName = "SignDocument", Document = ConvertHepler.ConvertDocumentToByteArray(signedDocument) } };
            }
            else
            {
                GetSignPersonByName(signerPerson).Documents.FirstOrDefault(p => p.DocumentName == signedDocument.OriginalFileName).Document = ConvertHepler.ConvertDocumentToByteArray(signedDocument);
            }
        }

        public bool CheckDocumentExist(string documentName)
        {
            // This an example.
            // Actually, it will return object from a data base.
            return signDocumentList.AsEnumerable().Any(p => p.DocumentName == documentName);
        }

        public Document GetDocument(string documentName)
        {
            // This an example.
            // Actually, it will return object from a data base.
            Byte[] signDocumentBytes = signDocumentList.AsEnumerable().FirstOrDefault(p => p.DocumentName == documentName)?.Document;
            Document signDocument = ConvertHepler.ConvertByteArrayToDocument(signDocumentBytes);

            return signDocument;
        }
        
        public SignPerson GetSignPersonByName(string name)
        {
            // This an example.
            // Actually, it will return object from a data base.
            return signPersonList.AsEnumerable().FirstOrDefault(p => p.Name == name);
        }

        private void CreateTestData()
        {
            signPersonList = new List<SignPerson>
            {
                new SignPerson { PersonId = Guid.NewGuid(), Name = "SignPerson 1", Position = "Head of Department", Image = ConvertHepler.ConverImageToByteArray(PathToImage) },
                new SignPerson { PersonId = Guid.NewGuid(), Name = "SignPerson 2", Position = "Deputy Head of Department", Image = ConvertHepler.ConverImageToByteArray(PathToImage) }
            };
        }

        private List<SignPerson> signPersonList;
        private List<SignDocument> signDocumentList;
        private readonly string PathToSignedDocument = MyDir + "SignDocument.doc";
        private readonly string PathToImage = MyDir + @"Images\LogoSmall.png";
    }
}