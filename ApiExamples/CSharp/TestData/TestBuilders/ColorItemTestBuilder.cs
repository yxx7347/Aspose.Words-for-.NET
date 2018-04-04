using ApiExamples.TestData.TestClasses;

namespace ApiExamples.TestData.TestBuilders
{
    public class ColorItemTestBuilder
    {
        public string Name { get; set; }
        public int Color { get; set; }
        public double Value1 { get; set; }
        public double Value2 { get; set; }
        public double Value3 { get; set; }

        public ColorItemTestBuilder()
        {
            Name = "DefaultName";
            Color = System.Drawing.Color.Black.ToArgb();
            Value1 = 1.0;
            Value2 = 1.0;
            Value3 = 1.0;
        }

        public ColorItemTestClass Build()
        {
            return new ColorItemTestClass(Name, Color, Value1, Value2, Value3);
        }
    }
}
