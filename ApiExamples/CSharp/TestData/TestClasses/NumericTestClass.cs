using System;

namespace ApiExamples.TestData.TestClasses
{
    public class NumericTestClass
    {
        private readonly int mValue1;
        private readonly double mValue2;
        private readonly int mValue3;
        private readonly DateTime mDate;

        public NumericTestClass(int value1, double value2, int value3, DateTime date)
        {
            mValue1 = value1;
            mValue2 = value2;
            mValue3 = value3;
            mDate = date;
        }
    }
}
