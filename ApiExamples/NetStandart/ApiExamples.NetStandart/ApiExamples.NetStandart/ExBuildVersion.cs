// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;

namespace ApiExamples.NetStandart
{
    [TestFixture]
    public class ExBuildVersion : ApiExampleBase
    {
        [Test]
        public void ShowBuildVersionInfo()
        {
            //ExStart
            //ExFor:BuildVersionInfo
            //ExSummary:Shows how to use BuildVersionInfo to obtain information about this product.
            Console.WriteLine("I am currently using {0}, version number {1}.", BuildVersionInfo.Product, BuildVersionInfo.Version);
            //ExEnd
        }
    }
}