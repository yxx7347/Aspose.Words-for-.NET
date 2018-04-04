namespace ApiExamples.TestData.TestClasses
{
    public class ColorItemTestClass
    {
        private readonly string mName;
        private readonly int mColor;
        private readonly double mValue1;
        private readonly double mValue2;
        private readonly double mValue3;

        public ColorItemTestClass(string name, int color, double value1, double value2, double value3)
        {
            mName = name;
            mColor = color;
            mValue1 = value1;
            mValue2 = value2;
            mValue3 = value3;
        }
    }
}
