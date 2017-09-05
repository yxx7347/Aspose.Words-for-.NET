// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExXpsSaveOptions : ApiExampleBase
    {
        //ToDo: See dev changes
        [Test]
        public void OptimizeOutput()
        {
            Document doc = new Document(MyDir + "XPSOutputOptimize.docx");

            XpsSaveOptions saveOptions = new XpsSaveOptions();
            saveOptions.OptimizeOutput = true;

            doc.Save(MyDir + @"\Artifacts\XPSOutputOptimize.docx");
        }

        [Test]
        public void GetAllSpecificClasses()
        {
            Assembly assembly = Assembly.LoadFrom("Aspose.Words.dll");

            foreach (Type type in assembly.ExportedTypes)
            {
                Console.WriteLine("\nClass: " + type.FullName);

                IEnumerable<MethodInfo> methodInfos = type.GetMethods(BindingFlags.Public | BindingFlags.Instance).Where(p => p.DeclaringType.FullName == type.FullName);

                Console.WriteLine("\nMethods:");

                foreach (MethodInfo info in methodInfos)
                {
                    Console.WriteLine(info.Name);
                }
            }
        }

        [Test]
        public void GetAllPrivateClasses()
        {
            Assembly assembly = Assembly.LoadFrom("Aspose.Words.dll");

            foreach (Type type in assembly.ExportedTypes)
            {
                Console.WriteLine("\nClass: " + type.FullName);

                IEnumerable<MethodInfo> methodInfos = type.GetMethods(BindingFlags.NonPublic);

                Console.WriteLine("\nMethods:");

                foreach (MethodInfo info in methodInfos)
                {
                    if (info != null)
                    {
                        Console.WriteLine(info.Name);
                        Assert.Fail();
                    }
                }
            }
        }
    }
}