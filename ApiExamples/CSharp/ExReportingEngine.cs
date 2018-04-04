// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using ApiExamples.TestData;
using ApiExamples.TestData.TestBuilders;
using ApiExamples.TestData.TestClasses;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExReportingEngine : ApiExampleBase
    {
        private readonly String _image = ImageDir + "Test_636_852.gif";
        private readonly String _doc = MyDir + "ReportingEngine.TestDataTable.docx";

        [Test]
        public void SimpleCase()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateSimpleDocument("<<[s.Name]>> says: <<[s.Message]>>");

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");
            BuildReport(doc, sender, "s", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("LINQ Reporting Engine says: Hello World\f", doc.GetText());
        }

        [Test]
        public void StringFormat()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateSimpleDocument("<<[s.Name]:lower>> says: <<[s.Message]:upper>>, <<[s.Message]:caps>>, <<[s.Message]:firstCap>>");

            MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "hello world");
            BuildReport(doc, sender, "s");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("linq reporting engine says: HELLO WORLD, Hello World, Hello world\f", doc.GetText());
        }

        [Test]
        public void NumberFormat()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateSimpleDocument("<<[s.Value1]:alphabetic>> : <<[s.Value2]:roman:lower>>, <<[s.Value3]:ordinal>>, <<[s.Value1]:ordinalText:upper>>" + ", <<[s.Value2]:cardinal>>, <<[s.Value3]:hex>>, <<[s.Value3]:arabicDash>>");

            NumericTestClass sender = new NumericTestClass(1, 2.2, 200, DateTime.Parse("10.09.2016 10:00:00"));
            BuildReport(doc, sender, "s");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("A : ii, 200th, FIRST, Two, C8, - 200 -\f", doc.GetText());
        }

        [Test]
        public void DataTableTest()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.TestDataTable.docx");

            DataSet ds = DataSet.AddTestData();
            BuildReport(doc, ds, "ds");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestDataTable.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestDataTable.docx", MyDir + @"\Golds\ReportingEngine.TestDataTable Gold.docx"));
        }

        [Test]
        public void ProgressiveTotal()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.Total.docx");

            DataSet ds = DataSet.AddTestData();
            BuildReport(doc, ds, "ds");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.Total.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.Total.docx", MyDir + @"\Golds\ReportingEngine.Total Gold.docx"));
        }

        [Test]
        public void NestedDataTableTest()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.TestNestedDataTable.docx");

            DataSet ds = DataSet.AddTestData();
            BuildReport(doc, ds, "ds");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestNestedDataTable.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestNestedDataTable.docx", MyDir + @"\Golds\ReportingEngine.TestNestedDataTable Gold.docx"));
        }

        [Test]
        public void ChartTest()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.TestChart.docx");
            DataSet ds = DataSet.AddTestData();

            BuildReport(doc, ds.Managers, "managers");
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestChart.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestChart.docx", MyDir + @"\Golds\ReportingEngine.TestChart Gold.docx"));
        }

        [Test]
        public void BubbleChartTest()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.TestBubbleChart.docx");
            DataSet ds = DataSet.AddTestData();

            BuildReport(doc, ds.Managers, "managers");
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestBubbleChart.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestBubbleChart.docx", MyDir + @"\Golds\ReportingEngine.TestBubbleChart Gold.docx"));
        }

        [Test]
        public void SetChartSeriesColorsDynamically()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.SetColorDinamically.docx");
            DataSet ds = DataSet.AddTestData();

            BuildReport(doc, ds.Managers, "managers");
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.SetColorDinamically.docx");
        }

        [Test]
        public void ConditionalExpressionForLeaveChartSeries()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.TestRemoveChartSeries.docx");

            DataSet ds = DataSet.AddTestData();

            int condition = 3;

            BuildReport(doc, new object[] { ds.Managers, condition }, new[] { "managers", "condition" });
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestLeaveChartSeries.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestLeaveChartSeries.docx", MyDir + @"\Golds\ReportingEngine.TestLeaveChartSeries Gold.docx"));
        }

        [Test]
        public void ConditionalExpressionForRemoveChartSeries()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.TestRemoveChartSeries.docx");

            DataSet ds = DataSet.AddTestData();

            int condition = 2;

            BuildReport(doc, new object[] { ds.Managers, condition }, new[] { "managers", "condition" });
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.TestRemoveChartSeries.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.TestRemoveChartSeries.docx", MyDir + @"\Golds\ReportingEngine.TestRemoveChartSeries Gold.docx"));
        }

        [Test]
        public void IndexOf()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.TestIndexOf.docx");
            DataSet ds = DataSet.AddTestData();

            BuildReport(doc, ds, "ds");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("The names are: Name 1, Name 2, Name 3\f", doc.GetText());
        }

        [Test]
        public void IfElse()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.IfElse.docx");
            DataSet ds = DataSet.AddTestData();

            BuildReport(doc, ds.Managers, "m");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("You have chosen 3 item(s).\f", doc.GetText());
        }

        [Test]
        public void IfElseWithoutData()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.IfElse.docx");
            DataSet ds = new DataSet();

            BuildReport(doc, ds.Managers, "m");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            Assert.AreEqual("You have chosen no items.\f", doc.GetText());
        }

        [Test]
        public void ExtensionMethods()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.ExtensionMethods.docx");
            DataSet ds = DataSet.AddTestData();

            BuildReport(doc, ds, "ds");
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.ExtensionMethods.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.ExtensionMethods.docx", MyDir + @"\Golds\ReportingEngine.ExtensionMethods Gold.docx"));
        }

        [Test]
        public void Operators()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.Operators.docx");
            NumericDataSourceWithMethod testData = new NumericDataSourceWithMethod(1, 2.0, 3, null, true);

            ReportingEngine report = new ReportingEngine();
            report.KnownTypes.Add(typeof(NumericDataSourceWithMethod));
            report.BuildReport(doc, testData, "ds");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.Operators.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.Operators.docx", MyDir + @"\Golds\ReportingEngine.Operators Gold.docx"));
        }

        [Test]
        public void ContextualObjectMemberAccess()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.ContextualObjectMemberAccess.docx");
            DataSet ds = DataSet.AddTestData();

            BuildReport(doc, ds, "ds");
            doc.Save(MyDir + @"\Artifacts\ReportingEngine.ContextualObjectMemberAccess.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.ContextualObjectMemberAccess.docx", MyDir + @"\Golds\ReportingEngine.ContextualObjectMemberAccess Gold.docx"));
        }

        [Test]
        public void InsertDocumentDinamically()
        {
            Aspose.Words.Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.Document]>>");

            DocumentTestClass doc = new DocumentTestBuilder().WithDocument(new Document(MyDir + "ReportingEngine.TestDataTable.docx")).Build();

            BuildReport(template, doc, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by document");
        }

        [Test]
        public void InsertDocumentDinamicallyByStream()
        {
            // By stream
            Aspose.Words.Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentStream]>>");

            DocumentTestClass docStream = new DocumentTestBuilder().WithDocumentStream(new FileStream(this._doc, FileMode.Open, FileAccess.Read)).Build();

            BuildReport(template, docStream, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by stream");
        }

        [Test]
        public void InsertDocumentDinamicallyByBytes()
        {
            Aspose.Words.Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentBytes]>>");

            DocumentTestClass docBytes = new DocumentTestBuilder().WithDocumentBytes(File.ReadAllBytes(MyDir + "ReportingEngine.TestDataTable.docx")).Build();

            BuildReport(template, docBytes, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
        }

        [Test]
        public void InsertDocumentDinamicallyByUri()
        {
            Aspose.Words.Document template = DocumentHelper.CreateSimpleDocument("<<doc [src.DocumentUri]>>");

            DocumentTestClass docUri = new DocumentTestBuilder().WithDocumentUri("http://www.snee.com/xml/xslt/sample.doc").Build();

            BuildReport(template, docUri, "src", ReportBuildOptions.None);
            template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertDocumentDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertDocumentDinamically(uri) Gold.docx"), "Fail inserting document by uri");
        }

        //[Test]
        //public void InsertImageDinamically()
        //{
        //    // By stream
        //    Aspose.Words.Document template = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Stream]>>", ShapeType.TextBox);
        //    ImageTestClass docByStream = new ImageTestClass(new FileStream(this._image, FileMode.Open, FileAccess.Read));

        //    BuildReport(template, docByStream, "src", ReportBuildOptions.None);
        //    template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

        //    Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");

        //    // By image
        //    template = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Image]>>", ShapeType.TextBox);
        //    Image docByImg = new Image(Image.FromFile(this._image, true));

        //    BuildReport(template, docByImg, "src", ReportBuildOptions.None);
        //    template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

        //    Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");

        //    // By Uri
        //    template = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Uri]>>", ShapeType.TextBox);
        //    ImageTestClass docByUri = new ImageTestClass("http://joomla-aspose.dynabic.com/templates/aspose/App_Themes/V3/images/customers/americanexpress.png");

        //    BuildReport(template, docByUri, "src", ReportBuildOptions.None);
        //    template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

        //    Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(uri) Gold.docx"), "Fail inserting document by bytes");

        //    // By bytes
        //    template = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Bytes]>>", ShapeType.TextBox);
        //    ImageTestClass docByBytes = new ImageTestClass(File.ReadAllBytes(this._image));

        //    BuildReport(template, docByBytes, "src", ReportBuildOptions.None);
        //    template.Save(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx");

        //    Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.InsertImageDinamically.docx", MyDir + @"\Golds\ReportingEngine.InsertImageDinamically(stream,doc,bytes) Gold.docx"), "Fail inserting document by bytes");
        //}

        [Test]
        public void WithoutKnownType()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<[new DateTime()]:”dd.MM.yyyy”>>");

            ReportingEngine engine = new ReportingEngine();
            Assert.That(() => engine.BuildReport(doc, ""), Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void WorkWithKnownTypes()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”dd.MM.yyyy”>>");
            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”dd”>>");
            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”MM”>>");
            builder.Writeln("<<[new DateTime(2016, 1, 20)]:”yyyy”>>");

            builder.Writeln("<<[new DateTime(2016, 1, 20).Month]>>");

            ReportingEngine engine = new ReportingEngine();
            engine.KnownTypes.Add(typeof(DateTime));
            engine.BuildReport(doc, "");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.KnownTypes.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.KnownTypes.docx", MyDir + @"\Golds\ReportingEngine.KnownTypes Gold.docx"));
        }


        [Test]
        [Ignore("WORDSNET-16258")]
        public void StretchImagefitHeight()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Stream] -fitHeight>>", ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestClass(new FileStream(this._image, FileMode.Open, FileAccess.Read));
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            doc.Save(MyDir + "123.docx");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Aspose.Words.Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                // Assert that width is keeped and height is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        [Ignore("WORDSNET-16258")]
        public void StretchImagefitWidth()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Stream] -fitWidth>>", ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestClass(new FileStream(this._image, FileMode.Open, FileAccess.Read));
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Aspose.Words.Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox and 
                Assert.IsTrue(shape.ImageData.HasImage);

                // Assert that height is keeped and width is changed
                Assert.AreNotEqual(431.5, shape.Width);
                Assert.AreEqual(346.35, shape.Height);
            }

            dstStream.Dispose();
        }

        [Test]
        [Ignore("WORDSNET-16258")]
        public void StretchImagefitSize()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Stream] -fitSize>>", ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestClass(new FileStream(this._image, FileMode.Open, FileAccess.Read));
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Aspose.Words.Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                // Assert that height is changed and width is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreNotEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        [Ignore("WORDSNET-16258")]
        public void StretchImagefitSizeLim()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateTemplateDocumentWithDrawObjects("<<image [src.Stream] -fitSizeLim>>", ShapeType.TextBox);

            ImageTestClass imageStream = new ImageTestClass(new FileStream(this._image, FileMode.Open, FileAccess.Read));
            BuildReport(doc, imageStream, "src", ReportBuildOptions.None);

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Aspose.Words.Document(dstStream);
            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                // Assert that textbox size are equal image size
                Assert.AreEqual(346.35, shape.Height);
                Assert.AreEqual(258.54, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        public void WithoutMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder, new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            //Assert that build report failed without "ReportBuildOptions.AllowMissingMembers"
            Assert.That(() => BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.None), Throws.TypeOf<InvalidOperationException>());
        }

        [Test]
        public void WithMissingMembers()
        {
            DocumentBuilder builder = new DocumentBuilder();

            //Add templete to the document for reporting engine
            DocumentHelper.InsertBuilderText(builder, new[] { "<<[missingObject.First().id]>>", "<<foreach [in missingObject]>><<[id]>><</foreach>>" });

            BuildReport(builder.Document, new DataSet(), "", ReportBuildOptions.AllowMissingMembers);

            //Assert that build report success with "ReportBuildOptions.AllowMissingMembers"
            Assert.AreEqual(ControlChar.ParagraphBreak + ControlChar.ParagraphBreak + ControlChar.SectionBreak, builder.Document.GetText());
        }

        [Test]
        public void SetBackgroundColor()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "ReportingEngine.BackColor.docx");

            List<Colors> colors = new List<Colors>();
            colors.Add(new Colors
            {
                ColorCode = Color.Black,
                ColorName = "Black",
                Description = "Black color"
            });
            colors.Add(new Colors
            {
                ColorCode = Color.FromArgb(255, 0, 0),
                ColorName = "Red",
                Description = "Red color"
            });
            colors.Add(new Colors
            {
                ColorCode = Color.Empty,
                ColorName = "Empty",
                Description = "Empty color"
            });

            BuildReport(doc, colors, "Colors");

            doc.Save(MyDir + @"\Artifacts\ReportingEngine.BackColor.docx");

            Assert.IsTrue(DocumentHelper.CompareDocs(MyDir + @"\Artifacts\ReportingEngine.BackColor.docx", MyDir + @"\Golds\ReportingEngine.BackColor Gold.docx"));
        }

        private static void BuildReport(Aspose.Words.Document document, object dataSource, string dataSourceName, ReportBuildOptions reportBuildOptions)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.Options = reportBuildOptions;

            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Aspose.Words.Document document, object[] dataSource, String[] dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Aspose.Words.Document document, object dataSource, String dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }

        private static void BuildReport(Aspose.Words.Document document, object dataSource)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource);
        }
    }
}
